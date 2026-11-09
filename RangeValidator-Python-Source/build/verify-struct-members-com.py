"""Read-only integration verification with synthetic C/H and A/B workbooks."""
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from rangevalidator.workflow import run_comparison

root = Path(__file__).resolve().parent / 'struct-members-verification'
root.mkdir(exist_ok=True)
types = root / 'platform.h'
types.write_text('#ifndef PLATFORM_H\n#define PLATFORM_H 1\n#define STRUCT struct\n'
                 '#define uint unsigned int\ntypedef unsigned short u2;\n#endif\n', encoding='utf-8')
header = root / 'device.h'
header.write_text('#pragma once\n#if !defined(DEVICE_H)\n#define DEVICE_H 1\n'
    '#include "platform.h"\ntypedef STRUCT { uint obc_ready:1; u2 qcs_maxcurr; } can_chg_t;\n'
    'extern can_chg_t can_chg;\nextern const can_chg_t * const can_chg_p;\n'
    '#ifdef OPT_SDV\nextern int optional;\n#endif\nvoid Main(void);\n#endif\n', encoding='utf-8')
source = root / 'device.c'
paths = []
expressions = ['can_chg.obc_ready', 'can_chg.qcs_maxcurr',
               'can_chg_p->obc_ready', 'can_chg_p->qcs_maxcurr']
for form in ('A', 'B'):
    book = Workbook()
    sheet = book.active
    sheet.title = '1.Main'
    sheet.append(['Input/Output', 'Property', 'Name', 'Data Type', 'Scope']
                 if form == 'A' else ['Name', 'Type', 'Range'])
    for expr in expressions:
        row = [expr, '', '0~1']
        sheet.append(['Input', 'Global', *row] if form == 'A' else row)
    path = root / f'design_{form}.xlsx'
    book.save(path)
    book.close()
    paths.append(path)
hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
receipts = {}
for conditional in (False, True):
    definition = 'void Main(void) {}\n'
    if conditional:
        definition = '#if FEATURE\n' + definition + '#endif\n'
    source.write_text('#include "device.h"\n' + definition, encoding='utf-8')
    expected = ['REVIEW' if conditional else 'MATCH', 'SKIP'] * 4
    for backend in ('openpyxl', 'com'):
        session = run_comparison(None, [source, header, types], paths, backend=backend)
        assert not session.errors, session.errors
        assert [r.status for r in session.results] == expected, session.results
        assert not any(r.candidate for r in session.results)
        assert [r.cell for r in session.results] == ['E2', 'E3', 'E4', 'E5', 'C2', 'C3', 'C4', 'C5']
        receipts[f'{backend}_conditional_{conditional}'] = [
            dict(expression=r.expression, status=r.status, cell=r.cell, reason=r.reason)
            for r in session.results]
assert hashes == {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
(root / 'result.json').write_text(json.dumps(receipts, ensure_ascii=False, indent=2), encoding='utf-8')
print('PASS: COM and OpenPyXL, A/B formats, ordinary labels, conditional bitfield REVIEW, original hashes unchanged')
