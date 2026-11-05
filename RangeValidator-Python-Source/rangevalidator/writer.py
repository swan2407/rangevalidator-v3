"""Selected certified MISMATCH only. Excel is the only detailed-document Writer."""
from contextlib import contextmanager
import gc
import os
import re
from pathlib import Path
import uuid
import zipfile
from .workflow import canonical
from .engine import is_blank_range
from .manual_edits import AUTO_MISMATCH, AUTO_CREATE, MANUAL_REVIEW, ManualEdit, is_review_target, manual_value
from .readers import Win32Runtime, read_com, _com_bounds
from .preservation import manifest, verify
from .report import create_report
from .models import ReadFailure


@contextmanager
def input_locks(paths):
    import win32file
    handles = []
    try:
        for path in paths:
            handles.append(win32file.CreateFile(str(path),win32file.GENERIC_READ,
                win32file.FILE_SHARE_READ,None,win32file.OPEN_EXISTING,win32file.FILE_ATTRIBUTE_NORMAL,None))
        yield
    finally:
        for h in reversed(handles):
            h.Close()


@contextmanager
def excel_book(path,diag):
    runtime = Win32Runtime()
    app = books = book = None
    initialized = owned = False
    pid = None
    stage = 'com.STA.initialize'
    errors = []
    primary = None
    def mark(s):
        nonlocal stage
        stage=s
        diag.emit('COM',s)
    try:
        mark(stage)
        runtime.initialize()
        initialized = True
        before = runtime.process_ids()
        mark('com.DispatchEx')
        app = runtime.dispatch()
        mark('com.instance.verify')
        pid = runtime.pid(app)
        if not pid or pid in before:
            raise RuntimeError('Cannot prove new Excel ownership')
        owned = True
        mark('com.security')
        app.Visible=False
        app.DisplayAlerts=False
        app.ScreenUpdating=False
        app.EnableEvents=False
        app.AskToUpdateLinks=False
        app.AutomationSecurity=3
        books=app.Workbooks
        mark('com.Workbooks.Open')
        book=books.Open(Filename=str(Path(path).resolve()),UpdateLinks=0,ReadOnly=True,
            IgnoreReadOnlyRecommended=True,Notify=False,AddToMru=False,Password='',WriteResPassword='')
        if not book.ReadOnly:
            raise RuntimeError('Workbook did not open ReadOnly')
        yield book,mark
    except Exception as exc:
        primary = exc if isinstance(exc,ReadFailure) else ReadFailure(stage,exc)
        diag.emit('COM',primary.stage,'FAILED',primary.code)
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception as exc:
                errors.append(ReadFailure('com.Close(False)',exc).as_dict())
        book=books=None
        if app is not None and owned:
            try:
                app.Quit()
            except Exception as exc:
                errors.append(ReadFailure('com.Quit',exc).as_dict())
        app=None
        gc.collect()
        if initialized:
            try:
                runtime.uninitialize()
            except Exception as exc:
                errors.append(ReadFailure('com.STA.uninitialize',exc).as_dict())
        if owned:
            try:
                runtime.wait_exit(pid)
                diag.emit('COM','com.process.exit','OK')
            except Exception as exc:
                errors.append(ReadFailure('com.process.exit',exc).as_dict())
                diag.emit('COM','com.process.exit','FAILED')
    if primary or errors:
        err=primary or ReadFailure('cleanup',RuntimeError('Excel cleanup failed'))
        err.cleanup_errors.extend(errors)
        raise err




def com_manifest(book):
    """In-memory COM-only preservation evidence; no decrypted workbook copy.
    Image bytes/formatting wholly outside Find bounds cannot be certified here.
    """
    from .models import address
    result=dict(sheets=[],media=None,charts=[])
    total=0
    ws=cell=area=comment=links=None
    try:
        sheets=book.Worksheets
        if sheets.Count>512:
            raise ValueError('Sheet preservation budget exceeded')
        for index in range(1,sheets.Count+1):
            ws=sheets.Item(index)
            bounds=_com_bounds(ws)
            cells={};merges=set()
            if bounds:
                r1,c1,r2,c2=bounds
                total+=(r2-r1+1)*(c2-c1+1)
                if total>250000:
                    raise ValueError('COM preservation cell budget exceeded')
                for row in range(r1,r2+1):
                    for col in range(c1,c2+1):
                        cell=ws.Cells.Item(row,col)
                        v=cell.Formula if cell.HasFormula else cell.Value2
                        if cell.MergeCells:
                            area=cell.MergeArea
                            merges.add(str(area.Address).replace('$',''))
                            area=None
                        note=None
                        comment=cell.Comment
                        if comment is not None:
                            note=(comment.Text(),comment.Author)
                        comment=None
                        link=None
                        links=cell.Hyperlinks
                        if links.Count:
                            h=links.Item(1)
                            link=(h.Address,h.SubAddress)
                            h=None
                        links=None
                        cells[address(row,col)]=(v,None,link,note)
                        cell=None
            result['sheets'].append(dict(name=str(ws.Name),state=ws.Visible,cells=cells,rows={},columns={},
                merges=sorted(merges),freeze='',autofilter='',images=ws.Shapes.Count,charts=0,shapes=None))
            ws=None
        return result
    finally:
        links=comment=area=cell=ws=sheets=None

def effective_formats(book, native):
    """Excel-rendered format values, avoiding locale/theme XML serialization artifacts."""
    result={}
    ws=cell=font=interior=borders=border=None
    try:
        for sheet in native['sheets']:
            ws=book.Worksheets.Item(sheet['name'])
            shapes=[]
            collection=shape=None
            try:
                collection=ws.Shapes
                for i in range(1,collection.Count+1):
                    shape=collection.Item(i)
                    if shape.Type!=13:
                        raise ValueError('Non-picture drawing object requires preservation review')
                    shapes.append((shape.Name,shape.Type,*[round(float(getattr(shape,p)),3) for p in ('Left','Top','Width','Height','Rotation')],shape.Placement,bool(shape.Visible)))
                    shape=None
            finally:
                shape=collection=None
            result[('shapes',sheet['name'])]=shapes
            from openpyxl.utils.cell import coordinate_from_string,column_index_from_string
            positions=[coordinate_from_string(a) for a in sheet['cells']]
            row_ids={r for _,r in positions}|set(sheet['rows'])
            col_ids={column_index_from_string(c) for c,_ in positions}|set(sheet['columns'])
            dimensions={}
            for r in row_ids:
                cell=ws.Rows.Item(r)
                dimensions[r]=(cell.RowHeight,bool(cell.Hidden),cell.OutlineLevel)
                cell=None
            result[('rows',sheet['name'])]=dimensions
            dimensions={}
            for col in col_ids:
                cell=ws.Columns.Item(col)
                dimensions[col]=(cell.ColumnWidth,bool(cell.Hidden),cell.OutlineLevel)
                cell=None
            result[('columns',sheet['name'])]=dimensions
            for a in sheet['cells']:
                cell=ws.Range(a)
                font=cell.Font
                interior=cell.Interior
                borders=cell.Borders
                edges=[]
                for i in (7,8,9,10,11,12):
                    border=borders.Item(i)
                    line=border.LineStyle
                    edges.append((line,border.Weight,border.Color) if line!=-4142 else (line,))
                    border=None
                result[(sheet['name'],a)]=(font.Name,font.Size,bool(font.Bold),bool(font.Italic),font.Underline,
                    bool(font.Strikethrough),font.Color,interior.Pattern,interior.Color,interior.PatternColor,
                    cell.NumberFormat,cell.HorizontalAlignment,cell.VerticalAlignment,bool(cell.WrapText),
                    cell.Orientation,cell.IndentLevel,bool(cell.ShrinkToFit),bool(cell.Locked),bool(cell.FormulaHidden),tuple(edges))
                borders=interior=font=cell=None
            ws=None
    finally:
        border=borders=interior=font=cell=ws=None
    return result


def apply_formats(native,formats):
    for sheet in native['sheets']:
        sheet['shapes']=formats[('shapes',sheet['name'])]
        sheet['rows']=formats[('rows',sheet['name'])]
        sheet['columns']=formats[('columns',sheet['name'])]
        for a,prior in list(sheet['cells'].items()):
            sheet['cells'][a]=(prior[0],formats[(sheet['name'],a)],*prior[2:])

def _text_value(v):
    return '' if v is None else str(v).strip()


def _preflight_cell(book, row):
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', row.cell)
    if not match or (len(match[1]) == 3 and match[1] > 'XFD') or int(match[2]) > 1048576:
        return '유효한 단일 셀 주소가 아님'
    sheets = book.Worksheets
    try:
        ws = sheets.Item(row.sheet)
    except Exception:
        # Only a successfully enumerated inventory can prove the sheet is missing.
        names = {str(sheets.Item(i).Name).casefold() for i in range(1, int(sheets.Count) + 1)}
        if not names or row.sheet.casefold() in names:
            raise
        return '수정 대상 시트가 없음'
    cell = ws.Range(row.cell)
    if bool(cell.HasFormula):
        return '수식 칸'
    if bool(cell.MergeCells) and str(cell.MergeArea.Cells.Item(1,1).Address).replace('$','') != row.cell:
        return '병합 영역 안쪽 칸'
    if row.status == 'CREATE' and not is_blank_range(cell.Value2):
        return 'CREATE 대상 셀이 더 이상 공란이 아님'
    if _text_value(cell.Value2) != _text_value(row.existing):
        return '값이 비교 때와 다름'
    return ''


class SaveOutcomes(dict):
    report_warning = ''


def _fixed_path(path):
    if path.stem.casefold().endswith('_fixed'):
        return path
    return path.with_name(path.stem + '_fixed' + path.suffix)


def _failure(diag, stage, exc):
    error = exc if isinstance(exc, ReadFailure) else ReadFailure(stage, exc)
    diag.emit('writer', stage, 'FAILED', error.code, api=error.api,
              exception_type=error.exception_type, inner_hresult=error.inner_hresult)
    return dict(stage=stage, api=error.api, code=error.code, exception_type=error.exception_type,
                inner_hresult=error.inner_hresult)


def _cleanup_warnings(diag, exc, phase):
    entries = getattr(exc, 'cleanup_errors', ())
    warnings = []
    for item in entries:
        diag.emit('writer', 'writer.cleanup', 'FAILED', item.get('code'),
                  api=item.get('api', item.get('stage', '')),
                  exception_type=item.get('exception_type', ''), inner_hresult=item.get('inner_hresult'))
        warnings.append(dict(status='CLEANUP_WARNING', phase=phase, api=item.get('stage'),
                             code=item.get('code'), exception_type=item.get('exception_type')))
    return warnings


def _apply_file(source,target,edits,diag,outcomes):
    """Write only preflight-approved cells, then verify the saved copy."""
    warnings = []
    saveas_returned = False
    stage = 'writer.open'
    diag.emit('writer',stage)
    try:
        with excel_book(source,diag) as pair:
            book,mark=pair
            diag.emit('writer',stage,'OK')
            ws=cell=None
            try:
                stage = 'writer.preflight'
                mark(stage)
                valid = []
                for r in edits:
                    reason = _preflight_cell(book, r)
                    if reason:
                        outcomes[r.id] = dict(file=source, output='', status='SKIPPED',
                                              stage=stage, reason=f'{r.sheet}!{r.cell} {reason}')
                    else:
                        valid.append(r)
                if valid:
                    edits = valid
                    stage = 'writer.output_check'
                    mark(stage)
                    if target.exists() or canonical(target)==canonical(source):
                        raise FileExistsError('출력 파일이 이미 있거나 원본과 같음 — 덮어쓰지 않음')
                    stage = 'writer.edit'
                    mark(stage)
                    for r in edits:
                        ws=book.Worksheets.Item(r.sheet)
                        cell=ws.Range(r.cell)
                        if str(r.expected)[:1] in '=+-@':
                            cell.NumberFormat='@'
                        cell.Value2=r.expected
                        cell=ws=None
                    stage = 'writer.output_check'
                    mark(stage)
                    if target.exists() or canonical(target)==canonical(source):
                        raise FileExistsError('출력 파일이 이미 있거나 원본과 같음 — 덮어쓰지 않음')
                    stage = 'writer.saveas'
                    diag.emit('writer',stage)
                    mark('com.SaveAs')
                    book.SaveAs(Filename=str(target),FileFormat=51,CreateBackup=False,AddToMru=False,Local=True)
                    saveas_returned = True
                    diag.emit('writer',stage,'OK')
                stage = 'writer.cleanup'
            finally:
                cell=ws=pair=book=None
        diag.emit('writer','writer.cleanup','OK')
    except Exception as exc:
        warnings.extend(_cleanup_warnings(diag, exc, 'source'))
        all_skipped = all(outcomes.get(r.id, {}).get('status') == 'SKIPPED' for r in edits)
        if getattr(exc, 'stage', '') != 'cleanup' or not warnings or not (saveas_returned or all_skipped):
            _record_save_failure(source,target,edits,diag,outcomes,stage,exc,saveas_returned,warnings)
            return
    if not saveas_returned:
        for r in edits:
            outcomes[r.id]['cleanup_warnings'] = warnings
        return

    stage = 'writer.output_check'
    diag.emit('writer',stage)
    try:
        if not target.is_file() or target.stat().st_size==0:
            raise RuntimeError('SaveAs 반환 후 출력 파일을 확인할 수 없음')
    except Exception as exc:
        _record_save_failure(source,target,edits,diag,outcomes,stage,exc,True,warnings)
        return
    diag.emit('writer',stage,'OK')

    stage = 'writer.reopen'
    diag.emit('writer',stage)
    pending = {}
    error_details = {}
    try:
        with excel_book(target,diag) as pair:
            saved,mark=pair
            ws=cell=None
            try:
                diag.emit('writer',stage,'OK')
                stage = 'writer.verify'
                mark(stage)
                for r in edits:
                    ws=saved.Worksheets.Item(r.sheet)
                    cell=ws.Range(r.cell)
                    formula = bool(cell.HasFormula)
                    actual = _text_value(cell.Value2)
                    if formula or actual != _text_value(r.expected):
                        reason = f'{r.sheet}!{r.cell} 기대값 {r.expected!r}, 확인값 {actual!r}'
                        if formula:
                            reason += ' (수식 셀)'
                        pending[r.id] = ('SAVED_VERIFY_FAILED', reason)
                    else:
                        pending[r.id] = ('SAVED_VERIFIED', '저장본 값 확인 완료')
                    cell=ws=None
                stage = 'writer.cleanup'
            finally:
                cell=ws=saved=pair=None
        diag.emit('writer','writer.cleanup','OK')
    except Exception as exc:
        verification_warnings = _cleanup_warnings(diag, exc, 'verification')
        warnings.extend(verification_warnings)
        if getattr(exc, 'stage', '') != 'cleanup' or not verification_warnings:
            error_details = _failure(diag,stage,exc)

    for r in edits:
        status, reason = pending.get(r.id, ('SAVED_UNVERIFIED', '저장됨 / 값 검증을 완료하지 못함'))
        details = (error_details if status == 'SAVED_UNVERIFIED' and error_details
                   else {'stage': 'writer.verify'})
        outcomes[r.id] = dict(file=source, output=str(target), status=status, reason=reason,
                              saveas_returned=True, cleanup_warnings=warnings, **details)
    statuses = {outcomes[r.id]['status'] for r in edits}
    if 'SAVED_VERIFY_FAILED' in statuses:
        diag.emit('writer','writer.verify','FAILED')
    elif statuses == {'SAVED_VERIFIED'}:
        diag.emit('writer','writer.verify','OK')


def _record_save_failure(source,target,edits,diag,outcomes,stage,exc,returned=False,warnings=()):
    details = _failure(diag,stage,exc)
    try:
        output_path = str(target) if target.is_file() and canonical(target) != canonical(source) else ''
    except OSError:
        output_path = ''
    for r in edits:
        if outcomes.get(r.id, {}).get('status') == 'SKIPPED':
            outcomes[r.id]['cleanup_warnings'] = warnings
        else:
            outcomes[r.id] = dict(file=source, output=output_path, status='SAVE_FAILED',
                                  reason=str(exc), saveas_returned=returned, cleanup_warnings=warnings, **details)


def _output_root(output_parent):
    from datetime import datetime
    parent=Path(output_parent).resolve()
    parent.mkdir(parents=True,exist_ok=True)
    output=parent/('RangeValidator_'+datetime.now().strftime('%Y%m%d_%H%M%S'))
    n=1
    while output.exists():
        n+=1
        output=parent/('RangeValidator_'+datetime.now().strftime('%Y%m%d_%H%M%S')+f'_{n}')
    output.mkdir()
    return output


def _write_groups(session,chosen,output,progress=None):
    try:
        common_root=Path(os.path.commonpath([str(Path(p).parent) for p in session.details]))
    except ValueError:
        common_root=None
    groups={}
    for r in chosen:
        groups.setdefault(r.file,[]).append(r)
    outcomes=SaveOutcomes()
    for i,(source,edits) in enumerate(groups.items()):
        if progress:
            progress(i,len(groups),Path(source).name)
        if common_root is not None:
            try:
                target=output/Path(source).relative_to(common_root)
            except ValueError:
                target=output/Path(source).name
        else:
            target=output/Path(source).name
        target = _fixed_path(target)
        stage = 'writer.output_check'
        try:
            target.parent.mkdir(parents=True,exist_ok=True)
            stage = 'writer.lock'
            with input_locks([source]):
                stage = 'writer.open'
                _apply_file(source,target,edits,session.diagnostics,outcomes)
                stage = 'writer.unlock'
        except Exception as exc:
            if stage == 'writer.unlock':
                warning = dict(status='CLEANUP_WARNING', phase='lock', **_failure(session.diagnostics,stage,exc))
                for r in edits:
                    outcome = outcomes[r.id]
                    outcome['cleanup_warnings'] = [*outcome.get('cleanup_warnings', ()), warning]
            else:
                _record_save_failure(source,target,edits,session.diagnostics,outcomes,stage,exc)
        statuses = {outcomes[r.id]['status'] for r in edits} - {'SKIPPED'}
        file_status = next((state for state in ('SAVE_FAILED', 'SAVED_VERIFY_FAILED', 'SAVED_UNVERIFIED',
                                                'SAVED_VERIFIED') if state in statuses), 'SKIPPED')
        for r in edits:
            outcomes[r.id]['file_status'] = file_status
            outcomes[r.id]['edit_kind'] = getattr(r, 'kind', AUTO_CREATE if r.status == 'CREATE' else AUTO_MISMATCH)
            if isinstance(r, ManualEdit):
                outcomes[r.id]['manual_value'] = r.expected
    if progress:
        progress(len(groups),len(groups),'완료')
    return outcomes


def _write_report(session, outcomes, output, *, manual_requests=()):
    try:
        create_report(session,outcomes,output/'RangeValidator_Report.xlsx',manual_requests=manual_requests)
    except Exception as exc:
        _failure(session.diagnostics, 'writer.report', exc)
        outcomes.report_warning = '보고서 생성 실패. 수정본 처리 결과는 유지됩니다.'


def _target(row):
    return canonical(row.file), row.sheet.casefold(), row.cell.upper()


def _validate_targets(session, chosen):
    sources = {canonical(path) for path in session.details}
    targets = [_target(r) for r in chosen]
    if any(source not in sources for source, _, _ in targets):
        raise ValueError('수정 대상이 비교한 상세설계서에 속하지 않음')
    if len(targets) != len(set(targets)):
        raise ValueError('같은 칸을 두 번 수정할 수 없음')


def _manual_edits(session, requests, rows, automatic):
    from collections import Counter

    selected = [request for request in requests if request.selected is True]
    request_counts = Counter(request.row_id for request in selected)
    files = {canonical(path) for path in session.details}
    records = {canonical(f.file): f for f in session.files}
    auto_targets = {_target(row) for row in automatic}
    pending, skipped = [], {}

    def skip(row, reason):
        skipped[row.id] = dict(file=row.file, output='', status='SKIPPED', stage='writer.preflight',
                               reason=reason, edit_kind=MANUAL_REVIEW)

    for request in selected:
        row = rows.get(request.row_id)
        if row is None or not is_review_target(row):
            continue
        file, _, _ = _target(row)
        record = records.get(file)
        if file not in files or (record and (record.status == 'ERROR' or record.read_status != 'SUCCESS')):
            skip(row, '읽기에 성공한 상세설계서에 속하지 않는 항목')
            continue
        if record and record.partial:
            errors = [e for e in session.errors if e.get('error_kind') == 'SHEET_READ_ERROR'
                      and canonical(e['file']) == file]
            if not errors or any(not e.get('sheet_name') or e['sheet_name'].casefold() == row.sheet.casefold()
                                 for e in errors):
                skip(row, '함수 시트의 읽기 근거가 불완전함')
                continue
        if request_counts[row.id] != 1:
            skip(row, '중복 수동 수정 요청')
            continue
        try:
            value = manual_value(request.value)
        except ValueError as exc:
            skip(row, str(exc))
            continue
        try:
            existing = manual_value(row.existing)
        except ValueError:
            existing = None
        if existing == value:
            skip(row, '기존 범위와 동일한 값')
            continue
        pending.append((row, value))

    target_counts = Counter(_target(row) for row, _ in pending)
    edits = []
    for row, value in pending:
        target = _target(row)
        if target in auto_targets or target_counts[target] != 1:
            skip(row, '중복 수정 대상 셀')
        else:
            edits.append(ManualEdit.from_result(row, value))
    return edits, skipped


def save_selected(session,selected_ids,output_parent,progress=None,*,manual_requests=()):
    """Save selected automatic edits and explicitly approved REVIEW values."""
    session.assert_unchanged()
    ids=list(selected_ids)
    if len(ids)!=len(set(ids)):
        raise ValueError('수정할 MISMATCH/CREATE 행을 선택하세요')
    byid={r.id:r for r in session.results}
    if any(i not in byid or byid[i].status not in ('MISMATCH','CREATE') or not byid[i].candidate for i in ids):
        raise ValueError('REVIEW/MATCH/SKIP 행은 자동 수정 대상이 아님')
    automatic=[byid[i] for i in ids]
    manual_requests = tuple(manual_requests)
    manual, skipped = _manual_edits(session, manual_requests, byid, automatic)
    chosen = automatic + manual
    if not chosen and not skipped:
        raise ValueError('수정할 MISMATCH/CREATE 또는 승인한 REVIEW 행을 선택하세요')
    _validate_targets(session,chosen)
    output=_output_root(output_parent)
    outcomes=_write_groups(session,chosen,output,progress)
    if skipped:
        states = {canonical(o['file']): o['file_status'] for o in outcomes.values()}
        for rid, outcome in skipped.items():
            outcome['file_status'] = states.get(canonical(outcome['file']), 'SKIPPED')
            outcomes[rid] = outcome
    _write_report(session,outcomes,output,manual_requests=manual_requests)
    return output,outcomes


def export_all(session,output_parent,progress=None):
    """Export: every document that has calculated-range differences (MISMATCH candidates) is saved to
    output_parent/RangeValidator_<time>/ with all its ranges updated. Documents without differences are not copied."""
    session.assert_unchanged()
    blocked={f.file for f in getattr(session,'files',[]) if f.read_status!='SUCCESS' or f.status=='ERROR'}
    chosen=[r for r in session.results if r.status in ('MISMATCH','CREATE') and r.candidate and r.file not in blocked]
    if not chosen:
        raise ValueError('선택 가능한 MISMATCH/CREATE가 없어 내보낼 파일이 없음')
    _validate_targets(session,chosen)
    output=_output_root(output_parent)
    outcomes=_write_groups(session,chosen,output,progress)
    _write_report(session,outcomes,output)
    return output,outcomes
