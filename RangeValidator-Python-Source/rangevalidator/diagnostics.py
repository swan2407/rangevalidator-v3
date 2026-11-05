"""Bounded, masked diagnostics. No exception messages/paths/cell data are logged."""
from dataclasses import dataclass
from datetime import datetime, timezone
import re
import threading
import uuid
from .error_info import ERROR_CLASSES, classify_error

STAGES = frozenset('''path backend cleanup fallback dependency.pywin32 dependency.openpyxl
com.sheet.error com.cell.error com.cellwise com.bounds.find com.bounds.usedrange com.bounds.end com.bounds.final com.bounds.blocked com.merge.restored com.merge.skipped com.extract
com.STA.initialize com.STA.uninitialize com.process.snapshot com.DispatchEx com.instance.verify
com.security com.Workbooks.Open com.readonly.verify com.sheets com.Find com.cell com.Value2
com.formula com.merge com.Close(False) com.Quit com.process.exit com.complete
openpyxl.open openpyxl.extract openpyxl.complete values.close formulas.close
source.read read.complete detail.read writer.recheck writer.edit writer.saved
com.SaveAs verification.saved.cells saved.reopen saved.verify writer.verify writer.lock writer.unlock writer.output writer.open writer.preflight writer.saveas writer.output_check writer.reopen writer.cleanup writer.report gui.run
com.Find.empty comparison.run comparison.result com.bounds.limit com.sheet.limit com.text.limit'''.split())
READERS = frozenset(('COM','openpyxl','auto','C','native','writer','verification','GUI'))
MESSAGES = frozenset(('FALLBACK_SAMPLING_MAY_OMIT_DATA','CELL_LIMIT_NO_PARTIAL_SUCCESS','MERGE_LIMIT_EDITS_BLOCKED','EXTRACTION_FAILED_EDITS_BLOCKED','Native Reader failed; DRM status unknown','Fingerprint unavailable; edits blocked','STABLE','CONTENT_CHANGED','METADATA_CHANGED','STATE_UNAVAILABLE'))
MASKED = '[민감정보 생략]'
HOT = frozenset(('com.cell','com.Value2','com.formula','com.merge'))


@dataclass(frozen=True)
class Event:
    time: str
    run_id: str
    reader: str
    stage: str
    status: str
    code: int | None = None
    hresult: str | None = None
    open_entered: bool = False
    last_success: str = ''
    failed_stage: str = ''
    excel_exit: str = ''
    message: str = ''
    win32_code: int | None = None
    file_index: int = 0
    input_role: str = ''
    sheet_index: int = 0
    row: int = 0
    column: int = 0
    search_phase: str = ''
    api: str = ''
    exception_type: str = ''
    error_class: str = ''
    inner_hresult: int | None = None
    attempt: int = 0
    max_attempts: int = 0
    empty_sheet: bool = False
    hash_changed: bool | None = None
    size_changed: bool | None = None
    mtime_changed: bool | None = None
    sheet_name: str = ''
    cell_address: str = ''
    decision: str = ''
    detail: str = ''


class Diagnostics:
    def __init__(self, callback=None):
        self.run_id = uuid.uuid4().hex
        self.events = []
        self.callback = callback
        self.lock = threading.Lock()
        self.local = threading.local()

    def set_context(self,**values):
        context=getattr(self.local,'context',{}).copy()
        for k,v in values.items():
            if k in ('file_index','sheet_index','row','column') and isinstance(v,int) and 0<=v<=1048576:context[k]=v
            elif k=='input_role' and v in ('common','header','detail'):context[k]=v
            elif k=='sheet_name' and isinstance(v,str):context[k]=v[:31]
        self.local.context=context

    def context(self):
        return dict(getattr(self.local,'context',{}))

    def _record(self, reader, stage, status, code, message, **metadata):
        if status == 'OK':
            from .com_read import READ_APIS
            self.local.last_success = metadata.get('api') if metadata.get('api') in READ_APIS else stage
        if stage=='com.process.exit' and status in ('OK','FAILED'):
            self.local.excel_exit=status
        event = Event(datetime.now(timezone.utc).isoformat(),self.run_id,reader,stage,status,
                      code,f'0x{code & 0xffffffff:08X}' if isinstance(code,int) and (code<0 or code>0xffff) else None,
                      getattr(self.local,'open_entered',False),getattr(self.local,'last_success',''),
                      stage if status=='FAILED' else '',getattr(self.local,'excel_exit',''),message,code if isinstance(code,int) and 0<=code<=0xffff else None)
        from dataclasses import replace
        from .com_read import READ_APIS
        data=self.context();data.update(metadata)
        safe={}
        for k in ('file_index','sheet_index','row','column','attempt','max_attempts'):
            v=data.get(k,0);safe[k]=v if isinstance(v,int) and 0<=v<=1048576 else 0
        safe['input_role']=data.get('input_role','') if data.get('input_role','') in ('','common','header','detail') else MASKED
        name=data.get('sheet_name',self.context().get('sheet_name',''))
        safe['sheet_name']=re.sub(r'[\x00-\x1f\x7f]','?',name[:31]) if isinstance(name,str) else ''
        safe['api']=data.get('api','') if data.get('api','') in READ_APIS|STAGES|{''} else MASKED
        safe['search_phase']=data.get('search_phase','') if data.get('search_phase','') in ('','cells','limits','first_row','last_row','first_column','last_column','coordinates') else MASKED
        typename=data.get('exception_type','')
        safe['exception_type']=typename if not typename or re.fullmatch(r'[A-Za-z_][A-Za-z_0-9.]*(?:<locals>[A-Za-z_0-9.]*)?',typename) else MASKED
        classification=data.get('error_class','') or (classify_error(code,safe['exception_type']) if status in ('FAILED','RETRY') else '')
        safe['error_class']=classification if classification in ERROR_CLASSES|{''} else MASKED
        safe['inner_hresult']=data.get('inner_hresult') if isinstance(data.get('inner_hresult'),int) else None
        safe['empty_sheet']=data.get('empty_sheet') is True
        address=data.get('cell_address','')
        safe['cell_address']=address[:16] if isinstance(address,str) and re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]{0,6}',address) else ''
        safe['decision']=data.get('decision','') if data.get('decision','') in ('ERROR','REVIEW','MISMATCH','MATCH','CREATE','SKIP','READ_ERROR','SHEET_READ_ERROR') else ''
        detail=data.get('detail','')
        safe['detail']=re.sub(r'[\x00-\x1f\x7f]',' ',detail[:600]) if isinstance(detail,str) else ''
        for k in ('hash_changed','size_changed','mtime_changed'):
            safe[k]=data.get(k) if isinstance(data.get(k),bool) else None
        event=replace(event,**safe)
        # Keep the first and each 1000th occurrence; never suppress failures.
        if status in ('ENTER','OK') and stage in HOT:
            counts=getattr(self.local,'counts',{})
            k=(stage,status);counts[k]=counts.get(k,0)+1;self.local.counts=counts
            if counts[k]>1 and counts[k]%1000:
                return event
        with self.lock:
            self.events.append(event)
            if len(self.events)>50000:
                del self.events[:10000]
        if self.callback:
            try:self.callback(event)
            except Exception:pass  # Logging cannot change read/cleanup behavior.
        return event

    def emit(self,reader,stage,status='ENTER',code=None,message='',**metadata):
        # Closed metadata vocabulary: arbitrary filename/cell/exception text never survives.
        reader=reader if reader in READERS else MASKED
        stage=stage if stage in STAGES else MASKED
        status=status if status in ('ENTER','OK','FAILED','SKIP','RETRY') else MASKED
        code=code if isinstance(code,int) and not isinstance(code,bool) else None
        message=message if message in MESSAGES else (MASKED if message else '')
        active=getattr(self.local,'active',None)
        # Existing Reader mark() calls occur only after the previous step returned.
        # A FAILED step is closed explicitly, so cleanup never certifies it as OK.
        if active and ((status=='ENTER' and (active!=(reader,stage) or not metadata.get('api'))) or (status=='OK' and active!=(reader,stage) and stage in ('openpyxl.complete','com.complete','read.complete'))):
            self._record(*active,'OK',None,'')
            self.local.active=None
        if stage in ('path','openpyxl.open') and status=='ENTER':
            self.local.last_success=''
            self.local.open_entered=False
            self.local.excel_exit=''
            self.local.counts={}
        if stage=='com.Workbooks.Open' and status=='ENTER':
            self.local.open_entered=True
        if status in ('FAILED','RETRY'):
            if reader=='openpyxl' and stage=='openpyxl.extract' and active==('openpyxl','openpyxl.open'):
                self._record(*active,'OK',None,'')
            self.local.active=None
        if status=='ENTER':
            self.local.active=(reader,stage)
        elif active==(reader,stage):
            self.local.active=None
        event=self._record(reader,stage,status,code,message,**metadata)
        return event


def format_event(event):
    cell=event.cell_address
    if not cell and event.row and event.column:
        n=event.column;letters=''
        while n:
            n,rem=divmod(n-1,26);letters=chr(65+rem)+letters
        cell=f'{letters}{event.row}'
    location=(f' | Sheet: {event.sheet_name}' if event.sheet_name else '')+(f' | Cell: {cell}' if cell else '')
    decision=(f' | Decision: {event.decision} | Reason: {event.detail}' if event.decision or event.detail else '')
    status={'ENTER':'진입','OK':'성공','FAILED':'실패','SKIP':'미실행','RETRY':'제한 재시도'}.get(event.status,event.status)
    return (f'[{event.time}] 실행 ID: {event.run_id}{location}{decision}\n'
            f'  Diagnostic: {event.message}\n'
            f'  {event.reader} | API/단계: {event.stage} | {status} | 코드: {event.code} | Win32: {event.win32_code} | HRESULT: {event.hresult or "없음"}\n'
            f'  입력: {event.input_role} #{event.file_index} | 시트 #{event.sheet_index} | 검색: {event.search_phase or "해당 없음"} | 행/열: {event.row}/{event.column} | 시도: {event.attempt}/{event.max_attempts}\n'
            f'  호출: {event.api or event.stage} | 예외: {event.exception_type or "없음"} | 분류: {event.error_class or "해당 없음"} | 내부 HRESULT: {event.inner_hresult} | 빈 시트: {event.empty_sheet}\n'
            f'  입력 변경 근거: SHA-256={event.hash_changed}, 크기={event.size_changed}, 수정시간={event.mtime_changed} (None=조회 불가/해당 없음)\n'
            f'  마지막 성공(last_success): {event.last_success or "없음"} | 실패 단계(failed_stage): {event.failed_stage or "없음"} | '
            f'Open 진입(open_entered): {"예" if event.open_entered else "아니오"} | Excel 종료: {event.excel_exit or "확인 중/해당 없음"}\n')
