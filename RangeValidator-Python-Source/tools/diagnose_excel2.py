"""Synthetic actual-COM before/after check. Excel-1 artifacts are never overwritten."""
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import hashlib
import json
from dataclasses import asdict
from time import perf_counter
from unittest.mock import patch
from openpyxl import load_workbook
from excel1_synthetic import build_suite, assess
from rangevalidator.workflow import run_comparison
from rangevalidator.readers import read_com
from rangevalidator.extraction_diagnostics import diagnose_session
from rangevalidator import com_values


def main():
    output=Path(sys.argv[1]).resolve(); output.mkdir(parents=True,exist_ok=True)
    source, paths, truth_path=build_suite(output/'synthetic')
    bulk=next(p for p in paths if p.name=='large_120.xlsx')
    wb=load_workbook(bulk)
    for name in ['제정 및 개정이력','목차','개요','구현파일관계']:
        ws=wb.create_sheet(name)
        for i in range(1,101): ws.cell(i,1,'Synthetic excluded content')
    wb.save(bulk);wb.close()
    reports=[]; receipts=[]
    real_copy=com_values._copy_sheet
    for path in paths:
        digest=hashlib.sha256(path.read_bytes()).hexdigest()
        def legacy(p, **kwargs):
            kwargs['function_sheets_only']=False
            return read_com(p,**kwargs)
        tick=perf_counter()
        with patch('rangevalidator.workflow.read_com',side_effect=legacy):
            before=run_comparison(None,[source],[path],backend='com')
        before_seconds=perf_counter()-tick
        bodies=[]
        def observe(ws,*args,**kwargs):
            bodies.append(str(ws.Name))
            return real_copy(ws,*args,**kwargs)
        tick=perf_counter()
        with patch('rangevalidator.com_values._copy_sheet',new=observe):
            after=run_comparison(None,[source],[path],backend='com')
        elapsed=perf_counter()-tick
        same=[asdict(r) for r in before.results]==[asdict(r) for r in after.results]
        unchanged=digest==hashlib.sha256(path.read_bytes()).hexdigest()
        report=diagnose_session(after)[0];reports.append(report)
        excluded=[s['sheet'] for s in report['sheets'] if not s['function_sheet']]
        receipt=dict(file=path.name,before_seconds=before_seconds,after_seconds=elapsed,
            result_rows_identical=same,original_unchanged=unchanged,
            body_read_sheets=len(bodies),excluded_body_calls=sum(n in excluded for n in bodies),
            counts=report['counts'],total_sheets=report['total_sheets'],
            timings=after.excel1_timings[str(path)])
        receipts.append(receipt)
        assert same and unchanged and receipt['excluded_body_calls']==0, receipt
        print(json.dumps(receipt,ensure_ascii=True),flush=True)
    assessment=assess(json.loads(truth_path.read_text(encoding='utf-8')),reports)
    for name,value in [('diagnostics',reports),('validation',receipts),('assessment',assessment)]:
        (output/(name+'.json')).write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
    print({k:v for k,v in assessment.items() if not isinstance(v,list)})

if __name__=='__main__':main()
