"""Synthetic, read-only COM/OpenPyXL verification for flexible catalog matching."""
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from rangevalidator.workflow import run_comparison

root = Path(__file__).resolve().parent / 'flexible-catalog-verification'
root.mkdir(exist_ok=True)
types = root / 'platform.h'
types.write_text('#define UINT unsigned int\ntypedef unsigned short u2;\n', encoding='utf-8')
header = root / 'device.h'
header.write_text('TYPEDEF STRUCT { uint Ready:3; U2 Current; } can_chg_t;\n'
    'typedef Struct { unsigned int ready:5; } device_t;\n'
    'typedef struct { unsigned int value:2; } sensor_t;\n'
    'typedef STRUCT { unsigned int missing:3; } unrelated_t;\n', encoding='utf-8')
source = root / 'device.c'
paths = []
rows = [('CAN_CHG . READY', '0~7'), ('can_chg_p - \n> ready', '0~1'),
        ('CAN_CHG . CURRENT', '0~4'), ('DEVICE . READY', '0~1'),
        ('SENSOR_P - > VALUE', '0~3'), ('can_chg.missing', '0~7'),
        ('unknown.ready', '0~7'), ('CAN_CHG_P - > ready', None)]
for form in ('A', 'B'):
    book = Workbook()
    sheet = book.active
    sheet.title = '1.Main'
    sheet.append(['INPUT / OUTPUT', 'PROPERTY', ' NAME ', 'DATA TYPE', 'SCOPE']
                 if form == 'A' else [' NAME ', ' TYPE ', ' RANGE '])
    for expression, bounds in rows:
        row = [expression, '', bounds]
        sheet.append([None, None, *row] if form == 'A' else row)
    path = root / f'design_{form}.xlsx'
    book.save(path)
    book.close()
    paths.append(path)
hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
receipts = {}
for phase in ('declarations_without_includes', 'types_without_variables'):
    source.write_text('EXTERN CAN_CHG_T CAN_CHG;\nEXTERN CONST can_chg_t * CONST can_chg_p;\n'
                      if phase == 'declarations_without_includes' else '', encoding='utf-8')
    for backend in ('openpyxl', 'com'):
        session = run_comparison(None, [types, header, source], paths, backend=backend)
        assert not session.errors, session.errors
        expected = ['MATCH', 'MISMATCH', 'SKIP', 'MISMATCH', 'MATCH', 'REVIEW', 'REVIEW', 'CREATE'] * 2
        assert [r.status for r in session.results] == expected, session.results
        assert [r.expected for r in session.results] == ['0~7', '0~7', '', '0~31', '0~3', '', '', '0~7'] * 2
        assert [r.cell for r in session.results] == [f'{col}{row}' for col in ('E', 'C') for row in range(2, 10)]
        if phase == 'types_without_variables':
            assert all('변수 선언 연결 없이' in r.reason for r in session.results if r.status != 'REVIEW')
        assert all(not r.candidate for r in session.results if r.status in ('REVIEW', 'SKIP'))
        receipts[f'{phase}_{backend}'] = [dict(expression=r.expression, status=r.status,
            expected=r.expected, cell=r.cell, reason=r.reason) for r in session.results]
assert hashes == {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
(root / 'result.json').write_text(json.dumps(receipts, ensure_ascii=False, indent=2), encoding='utf-8')
print('PASS: COM/OpenPyXL A/B, STRUCT without macro, no includes, no variables, case/spacing, type-first matching, original hashes unchanged')
