"""Real Excel reads with injected pywintypes.com_error to verify bounded recovery."""
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import sys
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
import pywintypes
from rangevalidator import com_values
from rangevalidator.diagnostics import Diagnostics
from rangevalidator.readers import read_com

root = Path(__file__).resolve().parent / 'com-retry3-verification'
root.mkdir(exist_ok=True)
path = root / 'synthetic.xlsx'
book = Workbook()
book.active.title = '1.Recover'
book.create_sheet('2.Persistent')
book.create_sheet('3.Normal')
for sheet in book:
    sheet.append(['Name', 'Type', 'Range'])
    sheet.append(['can_chg.ready', 'unsigned int', '0~1'])
book.save(path)
book.close()
original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
calls = {'1.Recover': 0, '2.Persistent': 0}


class RangeProxy:
    def __init__(self, target, name):
        self.target, self.name = target, name

    def __getattr__(self, attribute):
        if attribute == 'Value2':
            calls[self.name] += 1
            if self.name == '2.Persistent' or calls[self.name] < 3:
                raise pywintypes.com_error(-2147352567, 'synthetic injected failure',
                    (0, 'Excel', 'synthetic test', None, 0, -2146827284), None)
        return getattr(self.target, attribute)


class SheetProxy:
    def __init__(self, target, name):
        self.target, self.name = target, name

    def __getattr__(self, attribute):
        return getattr(self.target, attribute)

    def Range(self, *args):
        return RangeProxy(self.target.Range(*args), self.name)


original_copy = com_values._copy_sheet


def copy_sheet(sheet, *args, **kwargs):
    name = str(sheet.Name)
    if name in calls:
        sheet = SheetProxy(sheet, name)
    return original_copy(sheet, *args, **kwargs)


diagnostics = Diagnostics()
with patch.object(com_values, '_copy_sheet', copy_sheet):
    result = read_com(path, _trace=diagnostics, function_sheets_only=True)
assert result.partial
assert calls == {'1.Recover': 3, '2.Persistent': 3}, calls
assert [s.name for s in result.sheets] == ['1.Recover', '3.Normal']
assert all(s.cells['C2'].value == '0~1' for s in result.sheets)
assert len(result.read_errors) == 1
error = result.read_errors[0]
assert (error['sheet_name'], error['api'], error['attempt'], error['max_attempts']) == (
    '2.Persistent', 'com.Value2', 3, 3), error
assert error['code'] & 0xffffffff == 0x80020009
assert error['inner_hresult'] == -2146827284
assert any(e.sheet_name == '1.Recover' and e.api == 'com.Value2' and e.status == 'OK' and e.attempt == 3
           for e in diagnostics.events)
assert original_hash == hashlib.sha256(path.read_bytes()).hexdigest()
(root / 'result.json').write_text(json.dumps(dict(calls=calls, partial=result.partial,
    sheets=[s.name for s in result.sheets], read_errors=result.read_errors, original_unchanged=True,
    events=[asdict(e) for e in diagnostics.events if e.status in ('RETRY', 'FAILED') or e.attempt == 3]),
    ensure_ascii=False, indent=2), encoding='utf-8')
print('PASS: real Excel with injected 0x80020009, success on attempt 3, persistent failure stops at 3, next sheet survives, original unchanged')
