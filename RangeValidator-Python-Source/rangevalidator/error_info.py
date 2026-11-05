"""Observed error classification only; never infer DRM denial from a code."""
ERROR_CLASSES=frozenset(('COM_CALL_REJECTED','COM_DISCONNECTED','COM_ERROR','OS_ERROR',
    'PYTHON_TYPE_ERROR','PYTHON_ATTRIBUTE_ERROR','PYTHON_VALUE_ERROR','PYTHON_ERROR',
    'READ_BUDGET_EXCEEDED','FIND_BOUNDS_INCONSISTENT','INPUT_STATE_UNAVAILABLE'))

def classify_error(code,exception_type):
    name=exception_type.rsplit('.',1)[-1]
    if name=='ReadBudgetExceeded':return 'READ_BUDGET_EXCEEDED'
    if name=='FindBoundsError':return 'FIND_BOUNDS_INCONSISTENT'
    if name=='InputUnstableError':return 'INPUT_STATE_UNAVAILABLE'
    if isinstance(code,int):
        value=code&0xffffffff
        if value in (0x80010001,0x8001010A):return 'COM_CALL_REJECTED'
        if value in (0x80010108,0x800706BA,0x80010007,0x80010012):return 'COM_DISCONNECTED'
        return 'COM_ERROR' if code<0 or code>0xffff else 'OS_ERROR'
    return {'TypeError':'PYTHON_TYPE_ERROR','AttributeError':'PYTHON_ATTRIBUTE_ERROR',
            'ValueError':'PYTHON_VALUE_ERROR'}.get(name,'PYTHON_ERROR')
