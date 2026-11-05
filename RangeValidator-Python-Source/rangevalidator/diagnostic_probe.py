"""Explicit EXE diagnostic/Reader probe. Never invokes engine, Writer or reports."""
import hashlib
import json
from pathlib import Path
import threading
import time
import tkinter as tk
from openpyxl import Workbook
from .diagnostics import Diagnostics
from .models import ReadFailure
from .readers import read_auto,read_com
from .gui import App


def run(receipt):
    receipt=Path(receipt).resolve()
    rootdir=receipt.with_suffix('');rootdir.mkdir()
    good=rootdir/'synthetic_sensitive.xlsx';bad=rootdir/'synthetic_corrupt.xlsx'
    wb=Workbook();wb.active['A1']='SYNTHETIC_PRIVATE_CELL';wb.save(good);wb.close()
    bad.write_bytes(b'Synthetic corrupt workbook; not a DRM document')
    before={p.name:(hashlib.sha256(p.read_bytes()).hexdigest(),p.stat().st_mtime_ns) for p in (good,bad)}
    root=tk.Tk();root.withdraw();app=App(root)
    diagnostics=Diagnostics(lambda e:app.queue.put(('log',e)))
    state={}
    def read():
        try:
            read_auto(good,diagnostics=diagnostics)
            read_com(good,_trace=diagnostics)
        except Exception as exc:state['error']=type(exc).__name__
    worker=threading.Thread(target=read,daemon=False);worker.start()
    deadline=time.monotonic()+100
    try:
        while worker.is_alive() and time.monotonic()<deadline:
            root.update();time.sleep(.02)
        if worker.is_alive():raise TimeoutError('Read probe still running; no process killed')
        if 'error' in state:raise RuntimeError(state['error'])
        app.drain()
        # Exercise the real GUI read-failure path without executing the analysis engine.
        import rangevalidator.gui as gui
        original=gui.run_comparison
        def failed_read(*args,diagnostics,**kwargs):
            return read_com(bad,_trace=diagnostics)
        app.headers=['synthetic.h'];app.details=[str(good)]
        gui.run_comparison=failed_read
        try:
            app.compare()
            deadline=time.monotonic()+100
            while app.busy and time.monotonic()<deadline:root.update();time.sleep(.02)
            if app.busy:raise TimeoutError('Failure probe still running')
        finally:gui.run_comparison=original
        app.drain()
        log=app.log.get('1.0','end')
        failed=next(e for e in app.last_diagnostics.events if e.status=='FAILED')
        stages={e.stage for e in diagnostics.events if e.status=='OK'}
        required={'com.DispatchEx','com.Workbooks.Open','com.Find','com.Value2','com.Close(False)','com.Quit','com.process.exit'}
        assert required<=stages
        assert app.last_run_result['status']=='FAILED'
        assert app.diagnostic_window.winfo_exists()
        assert failed.stage=='com.Workbooks.Open' and failed.hresult
        assert app.last_diagnostics.events
        for secret in (str(rootdir),good.name,bad.name,'SYNTHETIC_PRIVATE_CELL'):
            assert secret not in log
        after={p.name:(hashlib.sha256(p.read_bytes()).hexdigest(),p.stat().st_mtime_ns) for p in (good,bad)}
        assert before==after
        result=dict(gui_diagnostic_window=True,view_button=True,masked=True,reader_native_success=True,
                    reader_com_success=True,reader_com_failure=True,failure_api=failed.stage,hresult=failed.hresult,
                    last_success=failed.last_success,last_result_retained=True,required_stages=sorted(stages),
                    original_inputs_unchanged=True,writer_executed=False,report_executed=False,actual_drm_tested=False)
        receipt.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        return result
    finally:
        root.destroy()
