"""Explicit distribution self-check: creates its own sanitized files only."""
from pathlib import Path
from openpyxl import Workbook,load_workbook
from .workflow import run_comparison
from .writer import _fixed_path, save_selected


def run(folder):
    root=Path(folder).resolve();root.mkdir()
    header=root/'types.h';header.write_text('typedef struct { unsigned int active:1; } Device_t; extern Device_t *p;',encoding='utf-8')
    detail=root/'device_Detail.xlsx';w=Workbook();s=w.active;s.title='1.Main';s['A1']='Global(Input)';s.append(['Name','Type','Range']);s.append(['p->active','uint','0~99']);w.save(detail);w.close()
    session=run_comparison(None,[header],[detail])
    selected=[r.id for r in session.results if r.candidate]
    if len(selected)!=1:raise RuntimeError('Self-check candidate mismatch')
    before=detail.read_bytes()
    output,outcomes=save_selected(session,selected,root/'outputs')
    if not all(o['status']=='SAVED_VERIFIED' for o in outcomes.values()):raise RuntimeError(str(outcomes))
    if detail.read_bytes()!=before:raise RuntimeError('Self-check original changed')
    w=load_workbook(_fixed_path(output/detail.name))
    try:
        if w.active['C3'].value!='0~1':raise RuntimeError('Saved value incorrect')
    finally:w.close()
    return dict(candidate_count=1,verified=1,original_preserved=True,report=(output/'RangeValidator_Report.xlsx').is_file(),output=str(output))
