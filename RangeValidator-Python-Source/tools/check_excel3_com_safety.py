"""Actual COM incomplete/merged/formula A tables; synthetic inputs only."""
import sys,json,hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
from openpyxl import Workbook,load_workbook
from test_a_extraction import fixture
from rangevalidator.workflow import run_comparison
from rangevalidator.extraction_diagnostics import diagnose_session

def main():
    out=Path(sys.argv[1]).resolve();wb=Workbook();wb.remove(wb.active)
    variants=['missing_scope','missing_name','duplicate_scope','duplicate_name','description_name','formula','nonanchor']
    temporary=out/'synthetic/safety_fixture.xlsx'
    for i,variant in enumerate(variants,1):
        fixture(temporary,variant=variant)
        original=load_workbook(temporary);ws=wb.create_sheet(f'{i}.{variant}')
        for row in original.active:
            for cell in row:
                if cell.value is not None:ws[cell.coordinate]=cell.value
        for merged in original.active.merged_cells.ranges:ws.merge_cells(str(merged))
        original.close()
    wb.create_sheet('Common Parameter')['A1']='excluded'
    path=out/'synthetic/a_safety.xlsx';wb.save(path);wb.close()
    before=hashlib.sha256(path.read_bytes()).hexdigest()
    session=run_comparison(None,[out/'synthetic/excel1_types.h'],[path],backend='com')
    report=diagnose_session(session)[0]
    for i,variant in enumerate(variants,1):
        sheet=next(s for s in report['sheets'] if s['sheet']==f'{i}.{variant}')
        if variant in ('formula','nonanchor'):
            assert sheet['variables'][0]['comparison_status']=='REVIEW'
        else:
            assert not any(v['extracted'] for v in sheet['variables'])
            assert all(v['range_cell'] is None for v in sheet['variables'])
            assert sheet['tables'][0]['state']!='HEADER_FOUND'
    assert before==hashlib.sha256(path.read_bytes()).hexdigest()
    assert report['counts']['body_read_success']==7 and report['counts']['excluded_sheets']==1
    (out/'com_safety.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Actual COM: 7 safety cases passed, excluded body reads 0, original unchanged')

if __name__=='__main__':main()
