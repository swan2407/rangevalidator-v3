"""User-supplied excel_io flow adapted to V3; uncertain reads never certify edits."""
from contextlib import contextmanager
import gc
import threading
from time import perf_counter
from .models import Workbook, Sheet, Cell, ReaderLimits, ReadFailure, address
from .com_read import read_call, ReadBudgetExceeded, FindBoundsError

_local = threading.local()


def emit(diag, stage, status='OK', exc=None, **metadata):
    if diag:
        failure = (exc if isinstance(exc, ReadFailure) else ReadFailure(stage, exc)) if exc else None
        if failure:
            metadata.setdefault('api', failure.api)
            for name in ('attempt', 'max_attempts'):
                metadata.setdefault(name, failure.context.get(name, 0))
        diag.emit('COM', stage, status, failure.code if failure else None,
                  exception_type=failure.exception_type if failure else '',
                  inner_hresult=failure.inner_hresult if failure else None, **metadata)


def _excel(runtime=None, diag=None):
    state = getattr(_local, 'excel', None)
    if state is not None:
        if runtime is not None and runtime is not state['runtime']:
            raise RuntimeError('Different runtime in an active Excel session')
        return state['app']
    from .readers import Win32Runtime
    runtime = runtime if runtime is not None else Win32Runtime()
    state = dict(runtime=runtime, app=None, initialized=False, owned=False, pid=None, diag=diag)
    _local.excel = state
    stage = 'com.STA.initialize'
    try:
        emit(diag, stage, 'ENTER'); runtime.initialize(); state['initialized'] = True
        stage = 'com.DispatchEx'; emit(diag, stage, 'ENTER')
        if hasattr(runtime, 'dispatch_owned'):
            app, state['pid'] = runtime.dispatch_owned()
            emit(diag, 'com.instance.verify')
        else:
            # Explicitly injected runtimes supply their own isolated instances.
            app = runtime.dispatch()
            try: state['pid'] = runtime.pid(app)
            except Exception: state['pid'] = None
        state['app'] = app
        state['owned'] = True
        stage = 'com.security'; emit(diag, stage, 'ENTER')
        for prop, value in (('Visible', False), ('DisplayAlerts', False),
                            ('ScreenUpdating', False), ('EnableEvents', False)):
            setattr(app, prop, value)
        for prop, value in (('AskToUpdateLinks', False), ('AutomationSecurity', 3)):
            try: setattr(app, prop, value)
            except Exception: pass
        emit(diag, stage)
        return app
    except Exception as exc:
        failure = exc if isinstance(exc, ReadFailure) else ReadFailure(stage, exc)
        try: close_excel(diag)
        except ReadFailure as cleanup: failure.cleanup_errors.extend(cleanup.cleanup_errors)
        raise failure from None


def close_excel(diag=None):
    state = getattr(_local, 'excel', None)
    if state is None: return
    _local.excel = None
    diag = diag or state['diag']
    errors = []
    app = state.pop('app')
    if app is not None and state['owned']:
        try:
            emit(diag, 'com.Quit', 'ENTER'); app.Quit(); emit(diag, 'com.Quit')
        except Exception as exc:
            errors.append(ReadFailure('com.Quit', exc).as_dict()); emit(diag, 'com.Quit', 'FAILED', exc)
    app = None
    gc.collect()
    if state['initialized']:
        try:
            state['runtime'].uninitialize(); emit(diag, 'com.STA.uninitialize')
        except Exception as exc: errors.append(ReadFailure('com.STA.uninitialize', exc).as_dict())
    if state['owned'] and state['pid'] is not None:
        try:
            state['runtime'].wait_exit(state['pid']); emit(diag, 'com.process.exit')
        except Exception as exc:
            errors.append(ReadFailure('com.process.exit', exc).as_dict()); emit(diag, 'com.process.exit', 'FAILED', exc)
    if errors: raise ReadFailure('cleanup', RuntimeError('Excel cleanup failed'), errors)


@contextmanager
def excel_session(diag=None):
    outer = getattr(_local, 'session', False)
    _local.session = True
    primary = None
    try: yield
    except BaseException as exc:
        primary = exc
        raise
    finally:
        _local.session = outer
        if not outer:
            try: close_excel(diag)
            except ReadFailure as exc:
                if isinstance(primary, ReadFailure): primary.cleanup_errors.extend(exc.cleanup_errors)
                elif primary is None: raise


def _open(path, runtime=None, diag=None):
    app = _excel(runtime, diag)
    emit(diag, 'com.Workbooks.Open', 'ENTER')
    book = app.Workbooks.Open(Filename=str(path), UpdateLinks=0, ReadOnly=True,
        IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False)
    if book is None:
        # Match the proven adapter: DRM add-ins can consume Open's return.
        for i in range(1, app.Workbooks.Count + 1):
            candidate = app.Workbooks.Item(i)
            try:
                if str(candidate.Name).lower() == path.name.lower():
                    book = candidate
                    break
            except Exception:
                pass
        if book is None: raise RuntimeError('Excel did not return the opened workbook')
    return book


def _find_last(sheet, order, diag=None, failures=None, *, look_in=-4123):
    """Find(xlPrevious). What='?*' = cells with at least one character; empty-string phantom cells that DRM
    files carry up to the sheet end are not matched (they made the old '*' search return 1,048,576)."""
    phase = 'last_row' if order == 1 else 'last_column'
    try:
        return read_call('com.Find.' + phase, lambda: sheet.Cells.Find(
            What='?*', After=sheet.Cells.Item(1, 1), LookIn=look_in, LookAt=2,
            SearchOrder=order, SearchDirection=2, MatchCase=False), diag, search_phase=phase)
    except ReadFailure as exc:
        if failures is not None: failures.append(exc)
        return None


MAX_COLS = 256        # design documents are never wider; cap when the column Find fails
ROW_CAP = 200000      # cap when the row Find fails; block reading stops at the real end anyway


def _used_size(sheet, diag=None):
    def size():
        ur = sheet.UsedRange
        return int(ur.Row) + int(ur.Rows.Count) - 1, int(ur.Column) + int(ur.Columns.Count) - 1
    try:
        return read_call('com.bounds.usedrange', size, diag)
    except Exception:
        return 0, 0


def _limit_failure(stage, reason, row, col):
    failure = ReadFailure(stage, ReadBudgetExceeded(reason))
    failure.context.update(attempted_bounds=(1, 1, row, col), limit_reason=reason)
    return failure


def _last_cell(sheet, diag=None, *, strict=False, look_in=-4123):
    """Upper bound of the content rectangle. UsedRange is used as a cap only (never read as a whole):
    DRM files report the whole sheet as used, which made Value2 on it run out of memory."""
    row_cell = _find_last(sheet, 1, diag, look_in=look_in)
    col_cell = _find_last(sheet, 2, diag, look_in=look_in)
    last_row = read_call('com.Find.coordinates', lambda: int(row_cell.Row), diag,
                         search_phase='last_row') if row_cell is not None else 0
    last_col = read_call('com.Find.coordinates', lambda: int(col_cell.Column), diag,
                         search_phase='last_column') if col_cell is not None else 0
    emit(diag, 'com.bounds.find', row=last_row, column=last_col)
    if not last_row or not last_col:
        ur_row, ur_col = _used_size(sheet, diag)
        emit(diag, 'com.bounds.usedrange', row=ur_row, column=ur_col)
        if strict and ((not last_row and ur_row > ROW_CAP) or (not last_col and ur_col > MAX_COLS)):
            raise _limit_failure('com.bounds.limit', 'UsedRange exceeds fallback bounds',
                                 last_row or ur_row, last_col or ur_col)
        if not last_row:
            last_row = max(min(ur_row, ROW_CAP), 4096)      # UsedRange 가 작게 잡혀도 최소 4096행은 블록으로 확인
        if not last_col:
            last_col = max(min(ur_col, MAX_COLS), 64)       # 열은 최소 64열
    if strict and last_col > MAX_COLS:
        raise _limit_failure('com.bounds.limit', f'Content extends beyond column {MAX_COLS}', last_row, last_col)
    return last_row, min(last_col, MAX_COLS)


def _is_empty(v):
    return v is None or v == ''


def _count_a(sheet, r0, r1, c1):
    try:
        rng = sheet.Range(sheet.Cells.Item(r0, 1), sheet.Cells.Item(r1, c1))
        return int(sheet.Application.WorksheetFunction.CountA(rng))
    except Exception:
        return None


def _lines(raw, rows, cols):
    if rows == cols == 1 and not isinstance(raw, tuple):
        return [[raw]]
    if not isinstance(raw, tuple) or len(raw) != rows:
        raise ValueError('Incomplete COM range row count')
    lines = [list(line) if isinstance(line, tuple) else [line] for line in raw]
    if any(len(line) != cols for line in lines):
        raise ValueError('Incomplete COM range column count')
    return lines


def _read_blocks(ws, last_row, last_col, diag=None, *, remaining=None, sparse=False, values_only=False):
    """Read ~130k-cell blocks through the observed extent, skipping empty blocks."""
    block = max(1 if sparse else 256, 131072 // max(last_col, 1))
    values, formulas, data_row, data_col = {}, {}, 0, 0
    content_cells = 0
    for r0 in range(1, last_row + 1, block):
        r1 = min(last_row, r0 + block - 1)
        if _count_a(ws, r0, r1, last_col) != 0:
            try:
                rng = read_call('com.cell', lambda: ws.Range(ws.Cells.Item(r0, 1), ws.Cells.Item(r1, last_col)),
                                diag, row=r0, column=1)
                vals = read_call('com.Value2', lambda: _lines(rng.Value2, r1-r0+1, last_col),
                                 diag, row=r1, column=last_col)
                fmls = None
                if not values_only:
                    fmls = read_call('com.formula.text', lambda: _lines(rng.Formula, r1-r0+1, last_col),
                                     diag, row=r1, column=last_col)
            except ReadFailure as exc:
                exc.context['attempted_bounds'] = (r0, 1, r1, last_col)
                raise
            for i, line in enumerate(vals):
                fline = fmls[i] if fmls is not None else ()
                if sparse:
                    row = {c: value for c, value in enumerate(line, 1)
                           if not _is_empty(value) or (fline and not _is_empty(fline[c - 1]))}
                    if not row:
                        continue
                    content_cells += len(row)
                    if remaining is not None and content_cells > remaining:
                        raise _limit_failure('com.bounds.limit', f'Content cell budget remaining: {remaining}',
                                             last_row, last_col)
                    data_row = r0 + i
                    values[data_row] = row
                    if fline:
                        formulas[data_row] = {column: fline[column - 1] for column in row}
                    continue
                n = len(line)
                while n and _is_empty(line[n - 1]) and (not fline or _is_empty(fline[n - 1])):
                    n -= 1
                if n:
                    data_row = r0 + i
                    data_col = max(n, data_col)
                    if remaining is not None and data_row * data_col > remaining:
                        raise _limit_failure('com.bounds.limit', f'Workbook cell budget remaining: {remaining}',
                                             last_row, last_col)
                    values[r0 + i] = line[:n]
                    if fline:
                        formulas[r0 + i] = fline[:n]
    return values, formulas, data_row


def _merge_state(ws, r0, r1, c1, c0=1, diag=None):
    try:
        return read_call('com.merge.flag',
                         lambda: ws.Range(f'{address(r0, c0)}:{address(r1, c1)}').MergeCells,
                         diag, row=r0, column=c0)
    except Exception:
        return None


def _merge_map(ws, last_row, last_col, diag=None, max_cells=200000, *, sparse=False):
    """Prune explicitly unmerged rectangles and cache confirmed MergeArea geometry."""
    from .readers import _merge_coordinates
    out, regions, geometry = {}, set(), {}
    if not last_row or not last_col or _merge_state(ws, 1, last_row, last_col, diag=diag) is False:
        return out, regions
    if not sparse and last_row * last_col > max_cells:
        raise _limit_failure('com.bounds.limit', f'Merge scan budget: {max_cells}', last_row, last_col)

    pending = [(1, 1, last_row, last_col, True)]
    probes = 0
    while pending:
        probes += 1
        if sparse and probes > min(50_000, max_cells):
            raise _limit_failure('com.bounds.limit', 'Merge inspection budget exceeded', last_row, last_col)
        r0, c0, r1, c1, checked = pending.pop()
        known = out.get((r0, c0))
        if known:
            # Remove the confirmed top-left intersection. The two remainder
            # rectangles partition the unverified cells without overlap.
            _, _, end_row, end_col = geometry[known[0]]
            end_row, end_col = min(r1, end_row), min(c1, end_col)
            if end_row < r1:
                pending.append((end_row + 1, c0, r1, c1, False))
            if end_col < c1:
                pending.append((r0, end_col + 1, end_row, c1, False))
            continue
        if r0 != r1 or c0 != c1:
            if not checked and _merge_state(ws, r0, r1, c1, c0, diag) is False:
                continue
            if r1 - r0 >= c1 - c0:
                middle = (r0 + r1) // 2
                pending.extend(((middle + 1, c0, r1, c1, False), (r0, c0, middle, c1, False)))
            else:
                middle = (c0 + c1) // 2
                pending.extend(((r0, middle + 1, r1, c1, False), (r0, c0, r1, middle, False)))
            continue

        try:
            cell = read_call('com.cell', lambda: ws.Range(address(r0, c0)), diag, row=r0, column=c0)
            merged = read_call('com.merge.flag', lambda: cell.MergeCells, diag, row=r0, column=c0)
            if merged is False:
                continue
            if merged is not True:
                raise ValueError('Single-cell merge state is unknown')
            region = read_call('com.merge.address', lambda: str(cell.MergeArea.Address).replace('$', ''),
                               diag, row=r0, column=c0)
            ar, ac, br, bc = _merge_coordinates(region)
            if not (ar <= r0 <= br and ac <= c0 <= bc):
                raise ValueError('MergeArea does not contain the queried cell')
            geometry[region] = ar, ac, br, bc
            anchor = address(ar, ac)
            regions.add(region)
            area = (min(last_row, br) - ar + 1) * (min(last_col, bc) - ac + 1)
            if sparse and len(out) + area > max_cells:
                raise _limit_failure('com.bounds.limit', f'Merged cell budget: {max_cells}', last_row, last_col)
            for row in range(ar, min(last_row, br) + 1):
                for column in range(ac, min(last_col, bc) + 1):
                    out[(row, column)] = (region, anchor)
        except Exception as exc:
            failure = exc if isinstance(exc, ReadFailure) else ReadFailure('com.merge', exc)
            failure.context.setdefault('attempted_bounds', (r0, c0, r0, c0))
            raise failure from None
    return out, regions


def _copy_sheet(ws, limits, remaining, diag=None, merge_limit=200000, *, sparse=False, values_only=False):
    row, col = _last_cell(ws, diag, strict=True)
    name = read_call('com.sheets.name', lambda: str(ws.Name), diag)
    visible = read_call('com.sheets.visibility', lambda: int(ws.Visible), diag)
    sheet = Sheet(name, {-1:'visible', 0:'hidden', 2:'veryHidden'}.get(visible, 'unknown'),
                  bounds=(1, 1, row, col) if row and col else None, merge_scope='content_bounds')
    emit(diag, 'com.bounds.final', row=row, column=col)
    if not row or not col:
        return sheet, 0, 0, []
    values, formulas, data_row = _read_blocks(ws, row, col, diag, remaining=remaining,
                                             sparse=sparse, values_only=values_only)
    data_col = max((max(v) if sparse else len(v) for v in values.values()), default=0)
    sheet.bounds = (1, 1, data_row, data_col) if data_row else None
    merges, regions = _merge_map(ws, data_row, data_col, diag, max_cells=merge_limit, sparse=sparse)
    text = 0
    for ri, line in values.items():
        fline = formulas.get(ri)
        for ci, value in (line.items() if sparse else enumerate(line, 1)):
            formula = None
            if fline:
                formula = fline[ci] if sparse else fline[ci - 1]
            has_formula = isinstance(formula, str) and formula.startswith('=')
            if has_formula and isinstance(value, str) and value == formula:
                has_formula = False                           # literal text that starts with '='
            formula = formula if has_formula else None
            if _is_empty(value) and not has_formula and (ri, ci) not in merges:
                continue
            a = address(ri, ci)
            region, anchor = merges.get((ri, ci), (None, None))
            if region and a != anchor:
                value, has_formula, formula = None, False, None
                if _is_empty(value):
                    pass
            kind = 'formula' if has_formula else 'value'
            if isinstance(value, int) and not isinstance(value, bool) and value < 0 and (value & 0xffff0000) == 0x800a0000:
                kind = 'error'
                value = {2000:'#NULL!',2007:'#DIV/0!',2015:'#VALUE!',2023:'#REF!',2029:'#NAME?',2036:'#NUM!',2042:'#N/A',2043:'#GETTING_DATA'}.get(value & 0xffff, '<Excel error>')
            if isinstance(value, bool):
                value = 'TRUE' if value else 'FALSE'
            text += (len(value) if isinstance(value, str) else 0) + len(formula or '')
            sheet.cells[a] = Cell(ri, ci, a, value, has_formula, formula, region, anchor, kind)
    for (ri, ci), (region, anchor) in merges.items():               # covered (non-anchor) merge cells keep their region
        a = address(ri, ci)
        if a not in sheet.cells:
            sheet.cells[a] = Cell(ri, ci, a, None, False, None, region, anchor, 'value')
    sheet.merged_ranges = sorted(regions)
    emit(diag, 'com.merge.restored' if regions else 'com.merge.skipped', row=data_row, column=data_col)
    cell_count = len(sheet.cells) if sparse else data_row * data_col
    if sparse and remaining is not None and cell_count > remaining:
        raise _limit_failure('com.bounds.limit', f'Content and merged cell budget remaining: {remaining}',
                             row, col)
    return sheet, cell_count, text, []


def read_workbook_values(path, limits=ReaderLimits(), *, _runtime=None, _trace=None, function_sheets_only=False):
    from .readers import input_path, check_limits
    from .engine import is_function_sheet
    stage = 'path'; book = com_sheet = None; failure = None; result = None
    own_session = not getattr(_local, 'session', False)
    timings = {'open_seconds': 0.0, 'sheet_count_seconds': 0.0,
               'sheet_names_seconds': 0.0, 'sheet_selection_seconds': 0.0, 'body_read_seconds': 0.0}
    try:
        emit(_trace, stage, 'ENTER'); check_limits(limits); path = input_path(path)
        stage = 'com.Workbooks.Open'; tick = perf_counter(); book = _open(path, _runtime, _trace)
        timings['open_seconds'] += perf_counter() - tick
        stage = 'com.readonly.verify'
        if not read_call('com.readonly.verify', lambda: book.ReadOnly, _trace):
            raise RuntimeError('Workbook did not open read-only')
        result = Workbook(str(path), 'com')
        stage = 'com.sheets'; tick = perf_counter()
        count = read_call('com.sheets.count', lambda: int(book.Worksheets.Count), _trace)
        timings['sheet_count_seconds'] += perf_counter() - tick
        result.source_sheet_count = count
        if count > limits.max_sheets:
            raise ReadFailure('com.sheet.limit', ReadBudgetExceeded(f'Worksheet count {count} exceeds {limits.max_sheets}'))
        result.read_timings = timings
        total = text = 0
        inventory = []
        # Complete the name-only inventory before any sheet body access.
        for index in range(1, count+1):
            if _trace: _trace.set_context(sheet_index=index,row=0,column=0,sheet_name='')
            tick = perf_counter()
            selection = 0.0
            try:
                com_sheet = read_call('com.sheets.item', lambda: book.Worksheets.Item(index), _trace)
                sheet_name = read_call('com.sheets.name', lambda: str(com_sheet.Name), _trace)
                select_started = perf_counter()
                included = not function_sheets_only or is_function_sheet(sheet_name)
                selection = perf_counter() - select_started
                observation = dict(index=index, sheet=sheet_name, included=included,
                                   states=['SHEET_NAME_DISCOVERED', 'SHEET_INCLUDED' if included else 'SHEET_EXCLUDED'])
                result.sheet_observations.append(observation)
                if _trace: _trace.set_context(sheet_name=sheet_name)
                emit(_trace, 'SHEET_NAME_DISCOVERED')
                emit(_trace, 'SHEET_INCLUDED' if included else 'SHEET_EXCLUDED')
                if included: inventory.append((index, sheet_name, observation))
            except Exception as exc:
                result.partial = True
                sf = exc if isinstance(exc, ReadFailure) else ReadFailure('com.sheets', exc)
                result.read_errors.append(dict(stage=sf.stage,api=sf.api,code=sf.code,
                    exception_type=sf.exception_type,error_class=sf.error_class,
                    inner_hresult=sf.inner_hresult,attempt=sf.context.get('attempt', 0),
                    max_attempts=sf.context.get('max_attempts', 0),
                    sheet_index=index,sheet_name='',search_phase=''))
                emit(_trace, 'com.sheet.error', 'FAILED', sf)
                del sf
            finally:
                timings['sheet_selection_seconds'] += selection
                timings['sheet_names_seconds'] += perf_counter() - tick - selection
                com_sheet = None
        for index, sheet_name, observation in inventory:
            if _trace: _trace.set_context(sheet_index=index,row=0,column=0,sheet_name=sheet_name)
            stage = 'com.extract'
            try:
                com_sheet = read_call('com.sheets.item', lambda: book.Worksheets.Item(index), _trace)
                observation['states'].append('SHEET_BODY_READ_STARTED')
                emit(_trace, 'SHEET_BODY_READ_STARTED')
                if _trace: _trace.set_context(sheet_name=sheet_name)
                tick = perf_counter()
                try:
                    sheet, cells, chars, cell_errors = _copy_sheet(com_sheet, limits, limits.max_cells-total, _trace)
                finally:
                    timings['body_read_seconds'] += perf_counter() - tick
                if text + chars > limits.max_text:
                    _, _, last_row, last_col = sheet.bounds
                    raise _limit_failure('com.text.limit', f'Workbook text budget: {limits.max_text}', last_row, last_col)
                if cell_errors:
                    result.partial=True
                    for item in cell_errors:
                        item.update(sheet_index=index, sheet_name=sheet_name, search_phase='')
                        result.read_errors.append(item)
                        emit(_trace,'com.cell.error','FAILED',code=item['code'],row=item['row'],column=item['column'],sheet_index=index,api='com.Value2',exception_type=item['exception_type'],error_class=item['error_class'])
                total += cells; text += chars
                result.sheets.append(sheet)
                observation['states'].append('SHEET_BODY_READ_COMPLETED')
                emit(_trace, 'SHEET_BODY_READ_COMPLETED')
            except Exception as sheet_exc:
                observation['states'].append('SHEET_BODY_READ_FAILED')
                emit(_trace, 'SHEET_BODY_READ_FAILED', 'FAILED', sheet_exc)
                result.partial = True
                sf = sheet_exc if isinstance(sheet_exc, ReadFailure) else ReadFailure('com.extract', sheet_exc)
                sf.backend = 'com'
                if _trace: sf.context.update(_trace.context())
                error = dict(stage=sf.stage,api=sf.api,code=sf.code,exception_type=sf.exception_type,
                             inner_hresult=sf.inner_hresult,
                             attempt=sf.context.get('attempt', 0),max_attempts=sf.context.get('max_attempts', 0),
                             error_class=sf.error_class,sheet_index=index,sheet_name=sheet_name,
                             search_phase=sf.context.get('search_phase',''))
                error.update({key:sf.context[key] for key in ('attempted_bounds','limit_reason') if key in sf.context})
                result.read_errors.append(error)
                if sf.error_class == 'READ_BUDGET_EXCEEDED':
                    # Preserve the sheet identity without exposing truncated cells.
                    result.sheets.append(Sheet(sheet_name, 'unknown'))
                bounds = error.get('attempted_bounds')
                emit(_trace,'com.sheet.error','FAILED',sf,sheet_index=index,
                     row=bounds[2] if bounds else 0,column=bounds[3] if bounds else 0,
                     detail=error.get('limit_reason',''))
                # Only primitive error evidence should survive until Quit.
                # The traceback can retain Range and Worksheet proxies.
                del sf
        emit(_trace, 'com.extract', 'OK')
    except Exception as exc:
        failure = exc if isinstance(exc,ReadFailure) else ReadFailure('com.Find' if isinstance(exc,FindBoundsError) else stage,exc)
        failure.backend = 'com'
        if _trace: failure.context.update(_trace.context())
        emit(_trace, failure.stage, 'FAILED', failure)
        emit(_trace, 'com.bounds.blocked', 'FAILED', message='EXTRACTION_FAILED_EDITS_BLOCKED')
    finally:
        com_sheet = None
        if book is not None:
            try: book.Close(SaveChanges=False); emit(_trace, 'com.Close(False)')
            except Exception as exc:
                close = ReadFailure('com.Close(False)',exc)
                emit(_trace, 'com.Close(False)', 'FAILED', close)
                failure = failure or ReadFailure('cleanup', RuntimeError('Workbook cleanup failed'))
                failure.cleanup_errors.append(close.as_dict())
        book = None
        if own_session or failure:
            try: close_excel(_trace)
            except ReadFailure as exc:
                if failure: failure.cleanup_errors.extend(exc.cleanup_errors)
                else: failure = exc
    if failure:
        failure.backend = 'com'
        raise failure
    emit(_trace, 'com.complete', 'OK')
    return result
