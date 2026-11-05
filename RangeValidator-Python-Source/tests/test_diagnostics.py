import hashlib,json,tempfile,time,unittest
from dataclasses import asdict
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
from rangevalidator.diagnostics import Diagnostics,MASKED,format_event
from rangevalidator.readers import read_com,read_auto
from rangevalidator.models import ReadFailure
from test_readers import FakeRuntime,FakeCell,ComError


class DiagnosticLogTests(unittest.TestCase):
    def test_review_log_has_sheet_cell_reason_but_not_cell_data_or_path(self):
        d=Diagnostics();d.set_context(input_role='detail',file_index=2,sheet_index=3,sheet_name='ControlMap')
        e=d.emit('native','comparison.result','OK',decision='REVIEW',cell_address='G28',detail='Unknown member or uncertified width/signedness')
        line=format_event(e)
        self.assertIn('ControlMap',line);self.assertIn('G28',line);self.assertIn('REVIEW',line)
        self.assertIn('Unknown member or uncertified width/signedness',line)
        self.assertNotIn('C:\\Company',line);self.assertNotIn('PRIVATE_CELL_VALUE',line)

    def test_com_cell_failure_log_formats_coordinate_and_sheet_name(self):
        d=Diagnostics();d.set_context(input_role='detail',file_index=1,sheet_index=4,sheet_name='BitFields')
        e=d.emit('COM','com.cell.error','FAILED',-2146827284,api='com.Value2',row=12,column=8)
        line=format_event(e)
        self.assertIn('BitFields',line);self.assertIn('H12',line)

    def test_success_transitions_and_completion(self):
        d=Diagnostics();read_com('synthetic.xlsx',_runtime=FakeRuntime(),_trace=d)
        for stage in ('com.DispatchEx','com.Workbooks.Open','com.Find','com.Value2','com.Close(False)','com.Quit','com.STA.uninitialize','com.process.exit','com.complete'):
            self.assertTrue(any(e.stage==stage and e.status=='OK' for e in d.events),stage)
        self.assertEqual(d.events[-1].excel_exit,'OK')
    def test_repeated_empty_find_retains_previous_success(self):
        d=Diagnostics();d.emit('COM','com.Find');d.emit('COM','com.Find')
        failed=d.emit('COM','com.Find','FAILED',-2146827284)
        self.assertEqual(failed.last_success,'com.Find')
        self.assertEqual(sum(e.stage=='com.Find' and e.status=='OK' for e in d.events),1)

    def test_native_extract_failure_last_success(self):
        d=Diagnostics()
        with patch('rangevalidator.readers.read_openpyxl',side_effect=ReadFailure('openpyxl.extract',ValueError('sensitive cell'))),self.assertRaises(ReadFailure):
            read_auto('synthetic.xlsx',diagnostics=d)
        failure=next(e for e in d.events if e.status=='FAILED')
        self.assertEqual(failure.last_success,'openpyxl.open')

    def test_open_failure_primary_and_cleanup_stages(self):
        d=Diagnostics()
        with self.assertRaises(ReadFailure):read_com('synthetic.xlsx',_runtime=FakeRuntime('open'),_trace=d)
        e=next(e for e in d.events if e.stage=='com.Workbooks.Open' and e.status=='FAILED')
        self.assertEqual((e.hresult,e.last_success,e.open_entered),('0x800A03EC','com.security',True))
        self.assertFalse(any(e.stage=='com.Workbooks.Open' and e.status=='OK' for e in d.events))
        self.assertFalse(any(e.stage=='com.Close(False)' for e in d.events))
        self.assertTrue(any(e.stage=='com.Quit' and e.status=='OK' for e in d.events))
    def test_find_and_value2_failure_api(self):
        for mode in ('find','value2'):
            rt=FakeRuntime('find' if mode=='find' else None)
            if mode=='value2':
                class Block:
                    @property
                    def Value2(self):raise ComError('PRIVATE_CELL_VALUE')
                rt.book.Worksheets.ws.Range = lambda *_: Block()
            d=Diagnostics()
            book = read_com('synthetic.xlsx',_runtime=rt,_trace=d)
            self.assertEqual(book.partial, mode == 'value2')
            if mode == 'value2':
                self.assertEqual(book.read_errors[0]['code'], -2146827284)
            stage='com.Find' if mode=='find' else 'com.Value2'
            self.assertTrue(any(e.stage==stage and e.status=='FAILED' and e.hresult=='0x800A03EC' for e in d.events))
            self.assertNotIn('PRIVATE_CELL_VALUE',str([asdict(e) for e in d.events]))
    def test_cleanup_failure_hresults_retained(self):
        d=Diagnostics()
        with self.assertRaises(ReadFailure):read_com('synthetic.xlsx',_runtime=FakeRuntime('cleanup'),_trace=d)
        failed=[e for e in d.events if e.status=='FAILED']
        self.assertEqual([e.stage for e in failed],['com.Close(False)','com.Quit'])
        self.assertTrue(all(e.hresult=='0x800A03EC' for e in failed))
    def test_masking_is_closed_by_default(self):
        d=Diagnostics()
        for secret in (r'C:\Company\secret.xlsx','secret.xlsx','PRIVATE_CELL_VALUE',r'\\server\share\workbook.xlsx'):
            d.emit(secret,secret,'FAILED',-2146827284,message=secret)
        serialized=json.dumps([asdict(e) for e in d.events],ensure_ascii=False)
        for secret in ('Company','secret.xlsx','PRIVATE_CELL_VALUE','server'):
            self.assertNotIn(secret,serialized)
        self.assertIn(MASKED,serialized)
    def test_native_success_missing_and_no_modification(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'sensitive.xlsx';w=Workbook();w.active['A1']='PRIVATE_CELL_VALUE';w.save(p);w.close()
            before=(hashlib.sha256(p.read_bytes()).hexdigest(),p.stat().st_mtime_ns)
            d=Diagnostics();read_auto(p,diagnostics=d)
            self.assertTrue(any(e.stage=='openpyxl.open' and e.status=='OK' for e in d.events))
            self.assertTrue(any(e.stage=='openpyxl.complete' and e.status=='OK' for e in d.events))
            self.assertEqual(before,(hashlib.sha256(p.read_bytes()).hexdigest(),p.stat().st_mtime_ns))
            with self.assertRaises(ReadFailure):read_auto(Path(tmp)/'missing.xlsx',diagnostics=d)
            self.assertTrue(any(e.status=='FAILED' and e.win32_code in (2,3) for e in d.events))
            self.assertNotIn('sensitive.xlsx',str(d.events));self.assertNotIn('PRIVATE_CELL_VALUE',str(d.events))
    def test_fallback_records_both_outcomes_without_drm_claim(self):
        d=Diagnostics()
        original=read_com
        def fallback(path,limits,_trace):return original(path,limits,_runtime=FakeRuntime(),_trace=_trace)
        with patch('rangevalidator.readers.read_openpyxl',side_effect=ReadFailure('openpyxl.open',ValueError('sensitive.xlsx'))),patch('rangevalidator.readers.read_com',side_effect=fallback):read_auto('synthetic.xlsx',diagnostics=d)
        self.assertTrue(any(e.reader=='openpyxl' and e.status=='FAILED' for e in d.events))
        self.assertTrue(any(e.stage=='com.complete' and e.status=='OK' for e in d.events))
        self.assertNotIn('sensitive.xlsx',str(d.events))
    def test_trace_never_changes_reader_calls_or_results(self):
        a,b=FakeRuntime(),FakeRuntime()
        x=read_com('synthetic.xlsx',_runtime=a)
        y=read_com('synthetic.xlsx',_runtime=b,_trace=Diagnostics())
        self.assertEqual(x,y);self.assertEqual(a.events,b.events)
        def broken_sink(_):
            raise RuntimeError('logging sink failed')
        self.assertEqual(x,read_com('synthetic.xlsx',_runtime=FakeRuntime(),_trace=Diagnostics(broken_sink)))
    def test_success_cannot_be_fabricated_after_failure(self):
        d=Diagnostics();d.emit('COM','com.security');d.emit('COM','com.Workbooks.Open');d.emit('COM','com.Workbooks.Open','FAILED',-2146827284)
        d.emit('COM','com.Quit');d.emit('COM','com.Quit','OK')
        self.assertFalse(any(e.stage=='com.Workbooks.Open' and e.status=='OK' for e in d.events))


class DiagnosticGuiTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        from rangevalidator.gui import App
        self.root=tk.Tk();self.root.withdraw();self.app=App(self.root)
        def close():
            try:self.root.destroy()
            except tk.TclError:pass
        self.addCleanup(close)
    def test_view_button_scroll_and_zoom(self):
        def walk(w):
            yield w
            for c in w.winfo_children():yield from walk(c)
        buttons=[w for w in walk(self.root) if w.winfo_class()=='TButton' and w.cget('text')=='진단 로그 보기']
        self.assertEqual(len(buttons),1);buttons[0].invoke()
        self.assertTrue(self.app.diagnostic_window.winfo_exists())
        self.assertTrue(self.app.diagnostic_log.cget('xscrollcommand'));self.assertTrue(self.app.diagnostic_log.cget('yscrollcommand'))
        self.app.zoom_logs(2);self.assertEqual(self.app.log_font_size,13)
    def test_failed_read_retains_object_and_last_result(self):
        def fail(*args,diagnostics,**kwargs):
            diagnostics.emit('COM','com.security');diagnostics.emit('COM','com.Workbooks.Open')
            diagnostics.emit('COM','com.Workbooks.Open','FAILED',-2146827284)
            raise ReadFailure('com.Workbooks.Open',ComError(r'C:\Company\sensitive.xlsx PRIVATE_CELL_VALUE'))
        self.app.common='synthetic.xlsx';self.app.headers=['synthetic.h'];self.app.details=['synthetic.xlsx']
        with patch('rangevalidator.gui.run_comparison',side_effect=fail),patch('rangevalidator.gui.messagebox.showerror') as modal:
            self.app.compare()
            deadline=time.monotonic()+10
            while self.app.busy and time.monotonic()<deadline:self.root.update();time.sleep(.01)
            self.assertFalse(self.app.busy);modal.assert_not_called()
        self.assertTrue(self.app.last_diagnostics.events)
        self.assertEqual(self.app.last_run_result['status'],'FAILED')
        self.assertTrue(self.app.diagnostic_window.winfo_exists())
        content=self.app.log.get('1.0','end')
        self.assertIn('0x800A03EC',content);self.assertIn('com.security',content)
        self.assertNotIn('sensitive.xlsx',content);self.assertNotIn('PRIVATE_CELL_VALUE',content)
        self.app.invalidate();self.assertIn('0x800A03EC',self.app.log.get('1.0','end'))
        self.assertEqual(self.app.last_run_result['status'],'FAILED')
    def test_failure_window_reopen_and_copy(self):
        d=Diagnostics(lambda e:self.app.queue.put(('log',e)));d.emit('COM','com.Workbooks.Open','FAILED',-2146827284)
        self.app.drain();self.app.diagnostic_window.destroy();self.app.show_diagnostics()
        self.app.copy_logs();self.assertIn('0x800A03EC',self.root.clipboard_get())
    def test_log_view_remains_available_while_busy(self):
        self.app.busy=True;self.app.show_diagnostics()
        self.assertTrue(self.app.diagnostic_window.winfo_exists());self.app.busy=False
