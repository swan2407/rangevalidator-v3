"""Actual COM extraction and one synthetic Writer verification; preserves earlier artifacts."""
import sys,json,hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
from openpyxl import Workbook,load_workbook
from excel1_synthetic import build_suite,assess
from test_a_extraction import fixture
from rangevalidator.workflow import run_comparison
from rangevalidator.extraction_diagnostics import diagnose_session,diagnose_workbook
from rangevalidator.readers import read_com
from rangevalidator.writer import save_selected

def main():
    out=Path(sys.argv[1]).resolve();out.mkdir(parents=True,exist_ok=True)
    source,paths,truth=build_suite(out/'synthetic')
    reports=[];sessions={}
    for path in paths:
        digest=hashlib.sha256(path.read_bytes()).hexdigest()
        session=run_comparison(None,[source],[path],backend='com')
        report=diagnose_session(session)[0];reports.append(report);sessions[path.name]=session
        assert digest==hashlib.sha256(path.read_bytes()).hexdigest()
        print(path.name,len(session.results),flush=True)
    assessment=assess(json.loads(truth.read_text(encoding='utf-8')),reports)
    old=json.loads((ROOT.parent/'excel2-diagnostic/diagnostics.json').read_text(encoding='utf-8'))
    def entries(data):
        return {(Path(v['file']).name,v['sheet'],v['name'],v['name_cell'],v['range_cell']):v['comparison_status']
                for r in data for s in r['sheets'] for v in s['variables'] if v['extracted']}
    before,after=entries(old),entries(reports)
    assert all(after.get(k)==v for k,v in before.items())
    added=sorted(set(after)-set(before));assert len(added)==6
    assert all(k[0]=='labels_spaced.xlsx' and k[1]=='1. spaced_a' for k in added)
    matrix=Workbook();matrix.remove(matrix.active);expected={};variants=['spaced','alias','merged_header',None]
    temp=out/'synthetic/fixture.xlsx'
    for i in range(105):
        size=(1,10,50,200,500)[i%5]
        coords=fixture(temp,count=size,row=(5,20,100)[i%3],name_col=(4,6,8)[i%3],variant=variants[i%4])
        src=load_workbook(temp);ws=matrix.create_sheet(f'{i+1}.A_function')
        for row in src.active:
            for cell in row:
                if cell.value is not None:ws[cell.coordinate]=cell.value
        for merge in src.active.merged_cells.ranges:ws.merge_cells(str(merge))
        src.close();expected[ws.title]=coords
    matrix.create_sheet('commonparameter')['A1']='excluded'
    matrix_path=out/'synthetic/a_105_sheets.xlsx';matrix.save(matrix_path);matrix.close()
    digest=hashlib.sha256(matrix_path.read_bytes()).hexdigest()
    session=run_comparison(None,[source],[matrix_path],backend='com')
    matrix_report=diagnose_session(session)[0]
    for sheet in matrix_report['sheets']:
        if sheet['function_sheet']:
            assert [(v['name_cell'],v['range_cell']) for v in sheet['variables']]==expected[sheet['sheet']]
            assert all(v['extracted'] for v in sheet['variables'])
    assert digest==hashlib.sha256(matrix_path.read_bytes()).hexdigest()
    save_session=sessions['a_10.xlsx'];candidate=next(r for r in save_session.results if r.candidate)
    original=hashlib.sha256(Path(candidate.file).read_bytes()).hexdigest()
    folder,outcomes=save_selected(save_session,[candidate.id],out/'writer')
    assert outcomes[candidate.id]['status']=='SAVED_VERIFIED'
    assert original==hashlib.sha256(Path(candidate.file).read_bytes()).hexdigest()
    sample=ROOT.parent/'A형.xlsx';sample_hash=hashlib.sha256(sample.read_bytes()).hexdigest()
    sample_report=diagnose_workbook(read_com(sample,function_sheets_only=True))
    assert sample_hash==hashlib.sha256(sample.read_bytes()).hexdigest()
    receipt=dict(previous_extracted=len(before),current_extracted=len(after),added_rows=added,
        existing_results_unchanged=True,b_results_unchanged=True,matrix_sheets=105,
        matrix_variables=sum(map(len,expected.values())),matrix_original_unchanged=True,
        writer=outcomes[candidate.id],sample_original_unchanged=True)
    for name,value in [('assessment',assessment),('diagnostics',reports),('matrix',matrix_report),
                       ('matrix_truth',expected),('sample',sample_report),('validation',receipt)]:
        (out/(name+'.json')).write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(receipt,ensure_ascii=True),flush=True)

if __name__=='__main__':main()
