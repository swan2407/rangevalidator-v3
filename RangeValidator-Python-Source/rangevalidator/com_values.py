"""User-supplied excel_io flow adapted to V3; uncertain reads never certify edits."""
from contextlib import contextmanager
import gc
import threading
from .models import Workbook, Sheet, Cell, ReaderLimits, ReadFailure, address
from .com_read import read_call, ReadBudgetExceeded, FindBoundsError

_local = threading.local()


def emit(diag, stage, status='OK', exc=None, **metadata):
    if diag:
        failure = (exc if isinstance(exc, ReadFailure) else ReadFailure(stage, exc)) if exc else None
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
        app = runtime.dispatch(); state['app'] = app
        state['owned'] = True
        try: state['pid'] = runtime.pid(app)
        except Exception: state['pid'] = None
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
        failure = ReadFailure(stage, exc)
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


def _find_last(sheet, order, diag=None, failures=None):
    """Find(xlPrevious). What='?*' = cells with at least one character; empty-string phantom cells that DRM
    files carry up to the sheet end are not matched (they made the old '*' search return 1,048,576)."""
    phase = 'last_row' if order == 1 else 'last_column'
    try:
        return read_call('com.Find.' + phase, lambda: sheet.Cells.Find(
            What='?*', After=sheet.Cells.Item(1, 1), LookIn=-4123, LookAt=2,
            SearchOrder=order, SearchDirection=2, MatchCase=False), diag, search_phase=phase)
    except ReadFailure as exc:
        if failures is not None: failures.append(exc)
        return None


MAX_COLS = 256        # design documents are never wider; cap when the column Find fails
ROW_CAP = 200000      # cap when the row Find fails; block reading stops at the real end anyway


def _used_size(sheet):
    try:
        ur = sheet.UsedRange
        return int(ur.Row) + int(ur.Rows.Count) - 1, int(ur.Column) + int(ur.Columns.Count) - 1
    except Exception:
        return 0, 0


def _last_cell(sheet, diag=None):
    """Upper bound of the content rectangle. UsedRange is used as a cap only (never read as a whole):
    DRM files report the whole sheet as used, which made Value2 on it run out of memory."""
    row_cell = _find_last(sheet, 1, diag)
    col_cell = _find_last(sheet, 2, diag)
    last_row = int(row_cell.Row) if row_cell is not None else 0
    last_col = int(col_cell.Column) if col_cell is not None else 0
    emit(diag, 'com.bounds.find', row=last_row, column=last_col)
    if not last_row or not last_col:
        ur_row, ur_col = _used_size(sheet)
        emit(diag, 'com.bounds.usedrange', row=ur_row, column=ur_col)
        if not last_row:
            last_row = max(min(ur_row, ROW_CAP), 4096)      # UsedRange 가 작게 잡혀도 최소 4096행은 블록으로 확인
        if not last_col:
            last_col = max(min(ur_col, MAX_COLS), 64)       # 열은 최소 64열
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
    if raw is None:
        return [[None] * cols for _ in range(rows)]
    if not isinstance(raw, tuple):
        return [[raw]]
    return [list(line) if isinstance(line, tuple) else [line] for line in raw]


def _read_blocks(ws, last_row, last_col, diag=None, stop_after=2):
    """Values and formula text in row blocks (about 130k cells per block). A block is skipped when CountA says it
    is empty; after `stop_after` consecutive empty blocks the rest is treated as phantom range. Memory is bounded
    by one block whatever the reported sheet size is."""
    block = max(256, 131072 // max(last_col, 1))
    values, formulas, empty_run, data_row = {}, {}, 0, 0
    for r0 in range(1, last_row + 1, block):
        r1 = min(last_row, r0 + block - 1)
        got = False
        if _count_a(ws, r0, r1, last_col) != 0:
            rng = ws.Range(ws.Cells.Item(r0, 1), ws.Cells.Item(r1, last_col))
            vals = _lines(read_call('com.Value2', lambda: rng.Value2, diag), r1 - r0 + 1, last_col)
            fmls = _lines(read_call('com.formula.text', lambda: rng.Formula, diag), r1 - r0 + 1, last_col)
            for i, line in enumerate(vals):
                n = len(line)
                while n and _is_empty(line[n - 1]) and (i >= len(fmls) or n > len(fmls[i]) or _is_empty(fmls[i][n - 1])):
                    n -= 1
                if n:
                    values[r0 + i] = line[:n]
                    formulas[r0 + i] = (fmls[i] if i < len(fmls) else [])[:n]
                    data_row = r0 + i
                    got = True
        if got:
            empty_run = 0
        else:
            empty_run += 1
            if empty_run >= stop_after:
                break
    return values, formulas, data_row


def _merge_state(ws, r0, r1, c1):
    try:
        return ws.Range(ws.Cells.Item(r0, 1), ws.Cells.Item(r1, c1)).MergeCells
    except Exception:
        return None


def _merge_map(ws, last_row, last_col, diag=None, chunk=64, max_cells=200000):
    """{(row, col): (region, anchor)} by asking MergeCells per range and narrowing (sheet → 64-row chunk → row → cell).
    Sheets without merges cost one call; the old per-cell scan cost one call per cell."""
    from .readers import _merge_coordinates
    out, regions = {}, set()
    if not last_row or not last_col or last_row * last_col > max_cells:
        emit(diag, 'com.merge.skipped', row=last_row, column=last_col)
        return out, regions
    if _merge_state(ws, 1, last_row, last_col) is False:
        return out, regions
    for r0 in range(1, last_row + 1, chunk):
        r1 = min(last_row, r0 + chunk - 1)
        if _merge_state(ws, r0, r1, last_col) is False:
            continue
        for r in range(r0, r1 + 1):
            if _merge_state(ws, r, r, last_col) is False:
                continue
            c = 1
            while c <= last_col:
                try:
                    cell = ws.Cells.Item(r, c)
                    if cell.MergeCells:
                        area = cell.MergeArea
                        region = str(area.Address).replace('$', '')
                        ar, ac, br, bc = _merge_coordinates(region)
                        anchor = address(ar, ac)
                        regions.add(region)
                        for mr in range(ar, min(last_row, br) + 1):
                            for mc in range(ac, min(last_col, bc) + 1):
                                out[(mr, mc)] = (region, anchor)
                        c = bc + 1
                        continue
                except Exception:
                    pass
                c += 1
    return out, regions


def _copy_sheet(ws, limits, remaining, diag=None, merge_limit=200000):
    row, col = _last_cell(ws, diag)
    sheet = Sheet(str(ws.Name), {-1:'visible', 0:'hidden', 2:'veryHidden'}.get(int(ws.Visible), 'unknown'),
                  bounds=(1, 1, row, col) if row and col else None, merge_scope='content_bounds')
    emit(diag, 'com.bounds.final', row=row, column=col)
    if not row or not col:
        return sheet, 0, 0, []
    values, formulas, data_row = _read_blocks(ws, row, col, diag)
    data_col = max((len(v) for v in values.values()), default=0)
    sheet.bounds = (1, 1, data_row, data_col) if data_row else None
    merges, regions = _merge_map(ws, data_row, data_col, diag)
    text = 0
    for ri, line in values.items():
        fline = formulas.get(ri, [])
        for ci, value in enumerate(line, 1):
            formula = fline[ci - 1] if ci - 1 < len(fline) else None
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
    return sheet, data_row * data_col, text, []


def read_workbook_values(path, limits=ReaderLimits(), *, _runtime=None, _trace=None):
    from .readers import input_path, check_limits
    stage = 'path'; book = com_sheet = None; failure = None; result = None
    own_session = not getattr(_local, 'session', False)
    try:
        emit(_trace, stage, 'ENTER'); check_limits(limits); path = input_path(path)
        stage = 'com.Workbooks.Open'; book = _open(path, _runtime, _trace)
        stage = 'com.readonly.verify'
        if not book.ReadOnly: raise RuntimeError('Workbook did not open read-only')
        result = Workbook(str(path), 'com')
        stage = 'com.sheets'; count = int(book.Worksheets.Count)
        total = text = 0
        for index in range(1, count+1):
            if _trace: _trace.set_context(sheet_index=index,row=0,column=0)
            stage = 'com.extract'
            sheet_name = ''
            try:
                com_sheet=book.Worksheets.Item(index)
                sheet_name = str(com_sheet.Name)
                if _trace: _trace.set_context(sheet_name=sheet_name)
                sheet, cells, chars, cell_errors = _copy_sheet(com_sheet, limits, limits.max_cells-total, _trace)
                if cell_errors:
                    result.partial=True
                    for item in cell_errors:
                        item.update(sheet_index=index, sheet_name=sheet_name, search_phase='')
                        result.read_errors.append(item)
                        emit(_trace,'com.cell.error','FAILED',code=item['code'],row=item['row'],column=item['column'],sheet_index=index,api='com.Value2',exception_type=item['exception_type'],error_class=item['error_class'])
                total += cells; text += chars
                result.sheets.append(sheet)
            except Exception as sheet_exc:
                result.partial = True
                sf = sheet_exc if isinstance(sheet_exc, ReadFailure) else ReadFailure('com.extract', sheet_exc)
                sf.backend = 'com'
                if _trace: sf.context.update(_trace.context())
                result.read_errors.append(dict(stage=sf.stage,api=sf.api,code=sf.code,exception_type=sf.exception_type,error_class=sf.error_class,sheet_index=index,sheet_name=sheet_name,search_phase=sf.context.get('search_phase','')))
                emit(_trace,'com.sheet.error','FAILED',sf,sheet_index=index)
                if isinstance(sheet_exc, ReadBudgetExceeded):
                    raise
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
