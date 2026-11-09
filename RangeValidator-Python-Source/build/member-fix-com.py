import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from rangevalidator.workflow import run_comparison

root = Path(__file__).resolve().parent / 'member-fix-com'
root.mkdir(exist_ok=True)
source = root / 'types.h'
source.write_text('', encoding='utf-8')
paths = []
expressions = ['obj.state', 'ptr->state', 'arr[0].state', 'obj.status.mode']
for form in ('A', 'B'):
    book = Workbook()
    sheet = book.active
    sheet.title = '1. update'
    sheet.append(['Input/Output', 'Property', 'Name', 'Data Type', 'Scope']
                 if form == 'A' else ['Name', 'Type', 'Range'])
    for expr in expressions:
        row = [expr, 'unsigned int', '0~7']
        sheet.append(['Input', 'Global', *row] if form == 'A' else row)
    path = root / f'{form}.xlsx'
    book.save(path)
    book.close()
    paths.append(path)
receipts = {}
for phase in ('missing_types', 'known_types'):
    if phase == 'known_types':
        source.write_text('typedef struct { unsigned int state:3; } T;'
                          'extern T obj; extern T *ptr; extern T arr[4];', encoding='utf-8')
    session = run_comparison(None, [source], paths, backend='com')
    assert not session.errors, session.errors
    expected = ['REVIEW'] * 4 if phase == 'missing_types' else ['MATCH'] * 3 + ['REVIEW']
    assert [r.status for r in session.results] == expected * 2
    assert not any(r.candidate for r in session.results)
    assert [r.cell for r in session.results] == ['E2', 'E3', 'E4', 'E5', 'C2', 'C3', 'C4', 'C5']
    receipts[phase] = [dict(expression=r.expression, status=r.status, cell=r.cell)
                       for r in session.results]
(root / 'result.json').write_text(json.dumps(receipts, indent=2), encoding='utf-8')
print('COM A/B missing types 8 REVIEW; known types 6 MATCH + 2 REVIEW; cell addresses verified')
