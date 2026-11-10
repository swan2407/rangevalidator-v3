"""Per-file comparison states and in-memory session integrity."""
from collections import Counter
from dataclasses import dataclass, field, replace, asdict
import hashlib
import json
import os
from pathlib import Path
import uuid
from time import perf_counter
from .diagnostics import Diagnostics
from .engine import catalog_headers, compare_book, is_function_sheet
from .readers import read_auto, read_workbook, read_com
from .models import ReadFailure
from .error_info import classify_error, read_error_summary


def source_path(path):
    try:return str(Path(path).resolve())
    except OSError:return os.path.abspath(os.fspath(path))


def canonical(path):
    return os.path.normcase(source_path(path))


@dataclass
class FileStatus:
    index: int
    file: str
    reader: str = 'unknown'
    read_status: str = 'NOT_ATTEMPTED'
    status: str = 'PENDING'
    error_kind: str = ''
    failed_api: str = ''
    failed_stage: str = ''
    code: int | None = None
    exception_type: str = ''
    error_class: str = ''
    sheet_index: int = 0
    search_phase: str = ''
    compared: bool = False
    extracted_variables: int = 0
    counts: dict = field(default_factory=lambda:{s:0 for s in ('MATCH','MISMATCH','REVIEW','CREATE','SKIP')})
    save_blocked: bool = True
    reason: str = ''
    partial: bool = False


@dataclass
class Session:
    headers: tuple[str,...]
    details: tuple[str,...]
    backend: str = 'auto'
    encoding: str = 'utf-8-sig'
    id: str = field(default_factory=lambda:uuid.uuid4().hex)
    results: list = field(default_factory=list)
    coverage: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    readers: dict = field(default_factory=dict)
    diagnostics: object = None
    books: dict = field(default_factory=dict,repr=False)
    files: list[FileStatus] = field(default_factory=list)
    catalog_summary: dict = field(default_factory=dict)
    excel1_timings: dict = field(default_factory=dict, repr=False)
    _signature: str = ''

    def signature(self):
        payload=dict(id=self.id,inputs=[self.headers,self.details,self.backend,self.encoding],
                     results=[asdict(r) for r in self.results],files=[asdict(f) for f in self.files],errors=self.errors)
        return hashlib.sha256(json.dumps(payload,sort_keys=True,ensure_ascii=False).encode()).hexdigest()

    def assert_unchanged(self):
        if self._signature!=self.signature():raise RuntimeError('REVIEW: analysis session was changed; compare again')


def _run_comparison(headers,details,backend='auto',encoding='utf-8-sig',diagnostics=None,progress=None):
    diag=diagnostics or Diagnostics()
    hdr=tuple(dict.fromkeys(source_path(p) for p in headers));det=tuple(dict.fromkeys(source_path(p) for p in details))
    s=Session(hdr,det,backend,encoding,diagnostics=diag)
    allpaths=[*hdr,*det]
    if len({canonical(p) for p in allpaths})!=len(allpaths):raise ValueError('Input roles contain duplicate files')
    s.files=[FileStatus(i,p) for i,p in enumerate(det,1)]
    # Catalog, read and compare per workbook, then successful finalization in the GUI.
    completed, total = 0, 2 + 2 * len(det)

    def advance(stage):
        if progress:
            progress(completed, total, stage)

    roles={p:('header',i) for i,p in enumerate(hdr,1)} | {p:('detail',i) for i,p in enumerate(det,1)}
    def context(path):
        role,index=roles[path];diag.set_context(input_role=role,file_index=index,sheet_index=0,row=0,column=0,sheet_name='')
    def read(path):
        context(path)
        if backend=='auto':book=read_auto(path,diagnostics=diag,function_sheets_only=True)
        elif backend=='com':book=read_com(path,_trace=diag,function_sheets_only=True)
        else:
            book=read_workbook(path,backend);diag.emit(backend,'read.complete','OK')
        s.readers[path]=book.backend;s.books[path]=book
        return book
    def error(path,exc,kind):
        code=getattr(exc,'code',None)
        if code is None:code=ReadFailure('detail.read',exc).code
        stage=getattr(exc,'stage','detail.read');api=getattr(exc,'api',stage)
        typename=getattr(exc,'exception_type',type(exc).__module__+'.'+type(exc).__qualname__)
        if stage=='cleanup' and getattr(exc,'cleanup_errors',None):
            first=exc.cleanup_errors[0];stage=first['stage'];api=first.get('api',stage);code=first.get('code');typename=first.get('exception_type',typename)
        ctx=getattr(exc,'context',{})
        item=dict(file=path,stage=stage,api=api,code=code,exception_type=typename,error_class=getattr(exc,'error_class',classify_error(code,typename)),error_kind=kind,
                  inner_hresult=getattr(exc,'inner_hresult',None),
                  file_index=roles[path][1],sheet_index=ctx.get('sheet_index',0),search_phase=ctx.get('search_phase',''),
                  limit_reason=ctx.get('limit_reason', getattr(exc, 'detail', '') if typename.endswith('.ReadBudgetExceeded') else ''),
                  native_read_failure=ctx.get('native_read_failure'))
        item['reason'] = read_error_summary(item) if kind == 'READ_ERROR' else kind + ': see file execution status'
        s.errors.append(item)
        return item
    diag.set_context(input_role='header',file_index=0,sheet_index=0,row=0,column=0,sheet_name='')
    advance('C/H 파일 분석 및 비트필드 카탈로그 구축 중')
    catalog=catalog_headers(hdr,encoding,diag)
    loaded={canonical(path) for path in catalog.files}
    missing=[path for path in hdr if canonical(path) not in loaded]
    if missing:
        raise ReadFailure('source.read', OSError('C/H input could not be read: ' + ', '.join(missing)))
    completed += 1
    s.catalog_summary=dict(files=len(catalog.files),structs=len(catalog.types),bitfields=sum(len(t.fields) for t in catalog.types),
                           certified=sum(1 for t in catalog.types for f in t.fields if f.minimum is not None),
                           variables=len(catalog.variables),typedefs=len(catalog.typedefs),macros=len(catalog.macros),reviews=list(catalog.reviews))
    try:common_root=Path(os.path.commonpath([str(Path(p).parent) for p in det])) if det else None
    except ValueError:common_root=None
    for path, record in zip(det, s.files):
        name = Path(path).name
        advance(f'Excel 열기·시트 읽기 중 — {name}')
        read_started = perf_counter()
        try:
            book=read(path);record.read_status='SUCCESS';record.reader=book.backend
        except Exception as exc:
            s.excel1_timings[path] = {'read_total_seconds': perf_counter() - read_started}
            record.reader=getattr(exc,'backend','') or ('com' if str(getattr(exc,'stage','')).startswith('com.') else backend)
            s.readers[path]=record.reader
            item=error(path,exc,'READ_ERROR');record.read_status='FAILED';record.status='ERROR';record.error_kind='READ_ERROR'
            for attr,k in (('failed_api','api'),('failed_stage','stage'),('code','code'),('exception_type','exception_type'),('error_class','error_class'),('sheet_index','sheet_index'),('search_phase','search_phase')):setattr(record,attr,item[k])
            record.reason=item['reason']
            continue
        s.excel1_timings[path] = {'read_total_seconds': perf_counter() - read_started,
                                  **book.read_timings}
        completed += 1
        advance(f'Excel 함수 시트 범위 비교 중 — {name}')
        comparison_started = perf_counter()
        try:
            relative=str(Path(path).relative_to(common_root)) if common_root else Path(path).name
            result,coverage=compare_book(None,catalog,book,relative)
            if book.partial:
                unknown = not book.read_errors or any(not e.get('sheet_name') for e in book.read_errors)
                affected = {e['sheet_name'] for e in book.read_errors
                            if e.get('sheet_name') and is_function_sheet(e['sheet_name'])}
                result = [replace(r, candidate=False, reason=r.reason + '; PARTIAL_READ: sheet evidence incomplete')
                          if r.candidate and (unknown or r.sheet in affected) else r for r in result]
            s.results.extend(result);s.coverage.extend((path,*r) for r in coverage)
            if diag:
                sheet_numbers={sheet.name:i for i,sheet in enumerate(book.sheets,1)}
                for r in result:
                    if r.status=='REVIEW':
                        diag.set_context(input_role='detail',file_index=record.index,
                                         sheet_index=sheet_numbers.get(r.sheet,0),sheet_name=r.sheet)
                        diag.emit('native','comparison.result','OK',decision='REVIEW',
                                  cell_address=r.cell,detail=r.reason)
            record.status='COMPLETED'
            record.partial=bool(book.partial)
            if book.partial:
                record.status='PARTIAL'
                record.error_kind='PARTIAL_READ'
                record.reason = ('PARTIAL_READ: 시트 식별 불가로 파일 전체 수정 차단' if unknown else
                                 'PARTIAL_READ: 읽기 오류가 있는 함수 시트만 수정 제외' if affected else
                                 'PARTIAL_READ: 비교 대상 밖 시트의 읽기 오류')
                if book.read_errors:
                    first_error = book.read_errors[0]
                    location = first_error.get('sheet_name') or '시트 확인 불가'
                    record.reason += f'\n{location} · {read_error_summary(first_error)}'
                for item in book.read_errors:
                    s.errors.append(dict(file=path,**item,error_kind='SHEET_READ_ERROR',file_index=record.index,
                                         reason=read_error_summary(item)))
            record.extracted_variables=sum(bool(r.cell and r.expression) for r in result)
            record.compared=bool(record.extracted_variables)
            record.counts={status:sum(r.status==status for r in result) for status in record.counts}
            record.save_blocked=not any(r.candidate for r in result)
            if not record.partial:
                record.reason='C/H 소스 기준 비교 완료'
            if not coverage and not record.partial:
                record.status='NO_VARIABLES';record.reason='비교 대상 함수 시트(이름에 숫자가 포함된 시트)가 없습니다.'
            elif not record.extracted_variables and not record.partial:
                record.status='NO_VARIABLES'
                layout_issues = [f'{sheet}: {reason}' for sheet, status, reason in coverage if status == 'REVIEW']
                record.reason = ('Excel 시트/헤더 확인 필요 — ' + '; '.join(layout_issues)
                                 if layout_issues else
                                 '함수 시트에서 구조체 멤버 표현식을 찾지 못했습니다. Name 열과 Range/Scope 헤더를 확인하세요.')
        except Exception as exc:
            item=error(path,exc,'COMPARISON_ERROR');record.status='ERROR';record.error_kind='COMPARISON_ERROR'
            record.failed_stage='comparison.run';record.failed_api='comparison.run';record.code=item['code'];record.exception_type=item['exception_type']
            record.reason='Comparison failed; no result rows committed'
            diag.set_context(input_role='detail',file_index=record.index,sheet_index=item.get('sheet_index',0))
            diag.emit('native','comparison.result','OK',decision='ERROR',detail=record.reason)
        else:
            completed += 1
        finally:
            s.excel1_timings[path]['comparison_seconds'] = perf_counter() - comparison_started
    advance('결과 집계 및 Excel 정리 중')
    s._signature=s.signature()
    return s


def run_comparison(common_path=None, headers=(), details=(), backend='auto', encoding='utf-8-sig', diagnostics=None, progress=None):
    """Compare C/H and detail files. The legacy common_path argument is ignored."""
    from .com_values import excel_session
    with excel_session(diagnostics):
        return _run_comparison(headers, details, backend, encoding, diagnostics, progress)
