"""Fresh report workbook. Source design workbooks are never opened here."""
from collections import Counter
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .manual_edits import AUTO_MISMATCH, AUTO_CREATE, MANUAL_REVIEW
from .result_index import ResultIndex, STATES
from .save_display import CLEANUP_LABEL, SAVE_LABELS
from .workflow import canonical

MISMATCH_HEADERS = ['원본 파일명','원본 파일 경로','함수 시트명','셀 주소','변수 표현식','원본 구조체 타입',
                    'C/H 근거 파일','C/H 선언 줄','기존 Excel 값','C/H 계산 범위','수정 요청값',
                    '원래 판정','수정 유형','수정 사유','사용자 선택 여부','최종 처리 상태','출력 파일 경로','처리 상세']
REVIEW_HEADERS = ['원본 파일명','원본 파일 경로','함수 시트명','셀 주소','변수 표현식','기존 Excel 값',
                  '관련 C/H 구조체 또는 후보','C/H 근거 파일 및 위치','REVIEW 판정 사유',
                  '수동 수정 선택 여부','사용자 입력값','최종 처리 상태','출력 파일 경로','처리 상세']
HEADERS = ['File','Sheet','Cell','Variable','Struct','Member','Bit width','Signed','Calculated Range',
           'Existing Range','Decision','Write result','Reason','Source','Source line','Reader',
           'File write result','Output path',CLEANUP_LABEL,'셀 저장 상태','파일 저장 상태']
KINDS = {AUTO_MISMATCH: '불일치 수정', AUTO_CREATE: '생성', MANUAL_REVIEW: '수동 수정'}
HEADER_FILL = PatternFill('solid', fgColor='24496B')


def literal(ws, row, values):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row, col)
        cell.value = value[:32767] if isinstance(value, str) else value
        if isinstance(value, str):
            cell.data_type = 's'


def header(ws, row, values):
    literal(ws, row, values)
    for cell in ws[row][:len(values)]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = HEADER_FILL


def _warning(outcome):
    return ', '.join(f"{w.get('phase')}: {w.get('api') or w.get('stage')} ({w.get('code')})"
                     for w in outcome.get('cleanup_warnings', ()))


def _entry(row, outcomes, requests, selected_ids):
    outcome = outcomes.get(row.id, {})
    request = requests.get(row.id)
    kind = outcome.get('edit_kind') or (
        AUTO_CREATE if row.status == 'CREATE' else AUTO_MISMATCH if row.status == 'MISMATCH'
        else MANUAL_REVIEW if request and request.selected else '')
    selected = bool(outcome) or row.id in selected_ids or bool(request and request.selected)
    value = (outcome.get('manual_value', request.value if request else '')
             if row.status == 'REVIEW' else row.expected if row.status in ('MISMATCH', 'CREATE') else '')
    state = outcome.get('status') or ('PENDING' if selected else
                                    'NOT_APPLICABLE' if row.status in ('MATCH', 'SKIP') else 'NOT_SELECTED')
    return dict(outcome=outcome, kind=kind, selected=selected, value=value, state=state)


def _summary(ws, session, index, entries):
    counts = index.counts
    selected = Counter(e['kind'] for e in entries.values() if e['selected'])
    saved = Counter(e['state'] for e in entries.values())
    sheets = [s for f in index.files.values() for s in f.sheets.values()]
    incomplete = {'ERROR', 'PARTIAL', 'REVIEW'}
    metrics = [
        ('실행 ID', session.id), ('분석 대상 Excel 파일 수', len(index.files)),
        ('대상 함수 시트 수', len(sheets)),
        ('분석한 함수 시트 수', sum(s.status not in incomplete for s in sheets)),
        ('오류·불완전 함수 시트 수', sum(s.status in incomplete for s in sheets)),
        *[('일반 멤버 (SKIP)' if state == 'SKIP' else state, counts[state]) for state in STATES],
        ('자동 MISMATCH 수정 선택 건수', selected[AUTO_MISMATCH]),
        ('CREATE 수정 선택 건수', selected[AUTO_CREATE]),
        ('REVIEW 수동 수정 선택 건수', selected[MANUAL_REVIEW]),
        *[(label + ' 건수', saved[state]) for state, label in SAVE_LABELS.items()],
    ]
    if getattr(session, 'build_options', None):
        options = session.build_options
        metrics.extend([('빌드옵션 XML', options['file']), ('빌드옵션 SHA256', options['sha256']),
                        ('빌드옵션 매크로', '\n'.join(f'{name}={value}' for name, value in options['macros'].items())),
                        ('빌드옵션 include 경로', '\n'.join(options['include_paths'])),
                        ('전처리 실패', '\n'.join(options.get('errors', {}).values()))])
    header(ws, 1, ['전체 집계', '건수 / 값'])
    for number, pair in enumerate(metrics, 2):
        literal(ws, number, pair)
    row = ws.max_row + 3
    header(ws, row, ['파일별 집계'])
    row += 1
    header(ws, row, ['원본 파일명','전체 경로','함수 시트 수',
                    *('일반 멤버 (SKIP)' if s == 'SKIP' else s for s in STATES),
                    *SAVE_LABELS.values(),'파일 저장 상태','출력 파일 경로',CLEANUP_LABEL,'읽기·분석 상태'])
    by_file = {}
    file_keys = {row.file: canonical(row.file) for row in index.rows.values()}
    for rid, entry in entries.items():
        by_file.setdefault(file_keys[index.rows[rid].file], []).append(entry)
    for key, file in index.files.items():
        items = by_file.get(key, [])
        results = Counter(e['state'] for e in items)
        states = {e['outcome'].get('file_status') for e in items} - {None, ''}
        output = sorted({e['outcome']['output'] for e in items if e['outcome'].get('output')})
        warnings = sorted({_warning(e['outcome']) for e in items} - {''})
        row += 1
        literal(ws, row, [Path(file.path).name,file.path,len(file.sheets),
                         *[file.counts[s] for s in STATES],*[results[s] for s in SAVE_LABELS],
                         ', '.join(sorted(states)), '\n'.join(output), '\n'.join(warnings),
                         file.record.status if file.record else ''])
    row += 3
    header(ws, row, ['시트별 집계'])
    row += 1
    header(ws, row, ['원본 파일명','전체 경로','함수 시트명',
                    *('일반 멤버 (SKIP)' if s == 'SKIP' else s for s in STATES),'분석 상태','상태 상세'])
    for file in index.files.values():
        for name, sheet in file.sheets.items():
            row += 1
            literal(ws, row, [Path(file.path).name,file.path,name,*[sheet.counts[s] for s in STATES],
                             sheet.status or '결과 수집',sheet.reason])


def _format(ws, table=True):
    ws.freeze_panes = 'A2'
    if table:
        ws.auto_filter.ref = ws.dimensions
        ws.print_title_rows = '1:1'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    for column in ws.columns:
        length = max((max((len(line) for line in str(c.value).splitlines()), default=0)
                      for c in column if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(column[0].column)].width = min(52, max(13, length + 3))
        for cell in column:
            cell.alignment = Alignment(vertical='top', wrap_text=True,
                                       horizontal='right' if isinstance(cell.value, (int, float)) else 'left')
    ws.print_area = ws.dimensions


def create_report(session, outcomes, path, *, manual_requests=(), selected_ids=()):
    """Write analysis and request history without changing either source."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        raise FileExistsError('Report already exists')
    view = SimpleNamespace(details=getattr(session, 'details', ()), files=getattr(session, 'files', ()),
                           coverage=getattr(session, 'coverage', ()), errors=session.errors,
                           results=session.results)
    index = ResultIndex(view)
    for key, file in index.files.items():
        errors = [e for e in session.errors if e.get('error_kind') == 'SHEET_READ_ERROR'
                  and canonical(e['file']) == key]
        if file.record and file.record.partial and (not errors or any(not e.get('sheet_name') for e in errors)):
            for sheet in file.sheets.values():
                sheet.status = 'PARTIAL'
                sheet.reason = '읽기 오류 시트를 식별할 수 없어 분석 완전성을 확인할 수 없음'
    requests = {request.row_id: request for request in manual_requests}
    selected_ids = set(selected_ids)
    entries = {rid: _entry(row, outcomes, requests, selected_ids) for rid, row in index.rows.items()}
    workbook = Workbook()
    try:
        workbook.active.title = 'Summary'
        for name in ('Mismatch', 'Review', 'Errors', 'Results'):
            workbook.create_sheet(name)
        _summary(workbook['Summary'], session, index, entries)
        header(workbook['Mismatch'], 1, MISMATCH_HEADERS)
        header(workbook['Review'], 1, REVIEW_HEADERS)
        header(workbook['Results'], 1, HEADERS)
        header(workbook['Errors'], 1, ['File','Stage','Code','Status','Reason','Sheet','API','Exception'])
        for rid, row in index.rows.items():
            entry = entries[rid]
            o, kind, state = entry['outcome'], entry['kind'], entry['state']
            common = [Path(row.file).name,row.file,row.sheet,row.cell,row.expression]
            selected = '선택' if entry['selected'] else '미선택'
            reason = o.get('reason', '')
            if row.status in ('MISMATCH', 'CREATE') or (row.status == 'REVIEW' and kind == MANUAL_REVIEW):
                ws = workbook['Mismatch']
                literal(ws, ws.max_row+1, [*common,row.struct,row.source,row.source_line or '',
                    row.existing,row.expected,entry['value'],row.status,kind,
                    KINDS.get(kind, ''),selected,state,o.get('output', ''),reason])
            if row.status == 'REVIEW':
                ws = workbook['Review']
                location = row.source + (f':{row.source_line}' if row.source and row.source_line else '')
                literal(ws, ws.max_row+1, [*common,row.existing,row.struct,location,row.reason,
                    selected,entry['value'],state,o.get('output', ''),reason])
            ws = workbook['Results']
            literal(ws, ws.max_row+1, [row.file,row.sheet,row.cell,row.expression,row.struct,row.member,
                row.width,row.signed,row.expected,row.existing,
                '일반 멤버 (SKIP)' if row.status == 'SKIP' else row.status,state,reason or row.reason,
                row.source,row.source_line,row.backend,o.get('file_status', ''),o.get('output', ''),
                _warning(o),SAVE_LABELS.get(state, ''),SAVE_LABELS.get(o.get('file_status'), '')])
        errors = workbook['Errors']
        for error in session.errors:
            literal(errors, errors.max_row+1, [error.get('file',''),error.get('stage',''),error.get('code'),
                    error.get('status') or error.get('error_kind',''),error.get('reason',''),
                    error.get('sheet_name',''),error.get('api',''),error.get('exception_type','')])
        for rid, entry in entries.items():
            outcome = entry['outcome']
            if entry['state'] in ('SAVE_FAILED','SAVED_UNVERIFIED','SAVED_VERIFY_FAILED','SKIPPED'):
                literal(errors, errors.max_row+1, [index.rows[rid].file,outcome.get('stage',''),
                        outcome.get('code'),entry['state'],outcome.get('reason',''),index.rows[rid].sheet,
                        outcome.get('api',''),outcome.get('exception_type','')])
            for warning in outcome.get('cleanup_warnings', ()):
                literal(errors, errors.max_row+1, [index.rows[rid].file,warning.get('api') or warning.get('stage'),
                        warning.get('code'),'CLEANUP_WARNING',warning.get('phase',''),index.rows[rid].sheet,
                        warning.get('api',''),warning.get('exception_type','')])
        for ws in workbook.worksheets:
            _format(ws, table=ws.title != 'Summary')
        with path.open('xb') as stream:
            workbook.save(stream)
    finally:
        workbook.close()
    return path
