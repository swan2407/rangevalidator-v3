"""Retry ONLY named read-only COM operations; never Open/SaveAs/Close/Quit."""
import time
from .models import ReadFailure

class ReadBudgetExceeded(ValueError):
    pass

class FindBoundsError(ValueError):
    pass

TRANSIENT = frozenset((0x80010001,0x8001010A))
READ_APIS = frozenset('''com.Find.cells com.Find.limits com.Find.first_row com.Find.last_row
com.Find.first_column com.Find.last_column com.Find.coordinates com.Find.bounds com.sheets.count com.sheets.item
com.sheets.name com.sheets.visibility com.cell com.Value2 com.formula.flag com.formula.text
com.merge.flag com.merge.area com.merge.address'''.split())


def parent_stage(api):
    for family in ('com.Find','com.sheets','com.formula','com.merge'):
        if api.startswith(family+'.'):return family
    return api


def pump_sta():
    try:
        import pythoncom
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass  # Message pumping is assistance, not a reason to suppress an API error.


def read_call(api,operation,diagnostics=None,*,search_phase='',row=0,column=0,
              attempts=3,sleep=time.sleep,pump=pump_sta):
    if api not in READ_APIS:
        raise ValueError('Read-only retry API is not allowed')
    if not 1<=attempts<=3:
        raise ValueError('COM read attempts must be 1..3')
    context=dict(api=api,search_phase=search_phase,row=row,column=column)
    def emit(status,code=None,exc=None,attempt=1):
        if diagnostics:
            try:diagnostics.emit('COM',parent_stage(api),status,code,
                **context,attempt=attempt,max_attempts=attempts,
                error_class=ReadFailure(parent_stage(api),exc).error_class if exc else '',
                exception_type=(type(exc).__module__+'.'+type(exc).__qualname__) if exc else '',
                inner_hresult=getattr(exc,'excepinfo',None)[5] if exc and isinstance(getattr(exc,'excepinfo',None),tuple) and len(exc.excepinfo)>5 and isinstance(exc.excepinfo[5],int) else None)
            except Exception:pass
    for attempt in range(1,attempts+1):
        emit('ENTER',attempt=attempt)
        try:
            result=operation()
        except Exception as exc:
            failure=ReadFailure(parent_stage(api),exc)
            failure.api=api
            failure.context=dict(context,attempt=attempt,max_attempts=attempts)
            if diagnostics and hasattr(diagnostics,'context'):
                inherited=diagnostics.context();inherited.update(failure.context);failure.context=inherited
            retry=isinstance(failure.code,int) and (failure.code&0xffffffff) in TRANSIENT and attempt<attempts
            emit('RETRY' if retry else 'FAILED',failure.code,exc,attempt)
            if not retry:raise failure from None
            pump();sleep((.1,.25)[attempt-1])
        else:
            emit('OK',attempt=attempt)
            return result
