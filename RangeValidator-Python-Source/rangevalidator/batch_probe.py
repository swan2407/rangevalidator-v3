"""Explicit read-only frozen batch acceptance probe; synthetic data only."""
import hashlib,json,time
from pathlib import Path
import tkinter as tk
from openpyxl import Workbook
from .gui import App


def run(receipt):
    receipt=Path(receipt).resolve();folder=receipt.with_suffix('');folder.mkdir()
    header=folder/'types.h';header.write_text('typedef struct { unsigned int active:1; } Device_t; extern Device_t *p;',encoding='utf-8')
    good=folder/'motor_Detail.xlsx';w=Workbook();w.active.title='1.Main';w.active['A1']='Global(Input)';w.active.append(['Name','Type','Range']);w.active.append(['p->active','uint','0~99']);w.save(good);w.close()
    bad=folder/'broken.xlsx';bad.write_bytes(b'Generated corrupt synthetic XLSX, no DRM')
    inputs=(header,good,bad)
    stamps=lambda:{p.name:(hashlib.sha256(p.read_bytes()).hexdigest(),p.stat().st_mtime_ns) for p in inputs}
    before=stamps();root=tk.Tk();root.withdraw();app=App(root)
    app.headers=[str(header)];app.details=[str(good),str(bad)]
    try:
        app.compare();deadline=time.monotonic()+150
        while app.busy and time.monotonic()<deadline:root.update();time.sleep(.02)
        if app.busy:raise TimeoutError('Batch probe did not finish')
        s=app.session
        assert s is not None and len(s.files)==2
        first,second=s.files
        assert first.read_status=='SUCCESS' and first.compared and first.counts['MISMATCH']==1
        assert second.status=='ERROR' and second.reader=='com' and not second.compared
        assert second.extracted_variables==0 and not sum(second.counts.values())
        assert len(s.results)==1 and s.results[0].member=='active' and s.results[0].expected=='0~1'
        assert s.results[0].candidate
        assert len(app.file_tree.get_children())==2
        failed = app.file_tree.get_children()[1]
        app.file_tree.selection_set(failed);app.show_file_status()
        assert 'ERROR' in app.file_tree.item(failed,'values') and not app.tree.get_children()
        assert stamps()==before
        assert app.diagnostic_window is not None and app.diagnostic_window.winfo_exists()
        result=dict(file_statuses=[dict(index=f.index,reader=f.reader,read_status=f.read_status,status=f.status,
            compared=f.compared,extracted=f.extracted_variables,counts=f.counts,failed_api=f.failed_api,code=f.code) for f in s.files],
            normal_bitfield_comparison=True,file_error_separate_from_review=True,partial_failure_preserved=True,
            gui_file_status=True,originals_unchanged=True,writer_executed=False,report_executed=False,actual_drm_tested=False)
        receipt.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        return result
    finally:root.destroy()
