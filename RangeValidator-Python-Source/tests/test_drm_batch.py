"""Sanitized batch/state/COM tests only, no Writer/SaveAs/report execution."""
import os,tempfile,time,unittest
from pathlib import Path
from dataclasses import replace,asdict
from unittest.mock import patch
from rangevalidator.workflow import run_comparison
from rangevalidator.diagnostics import Diagnostics
from rangevalidator.models import ReadFailure
from rangevalidator.readers import read_com,_com_bounds
from rangevalidator.com_read import read_call
from v3_fixtures import fixtures
from test_readers import FakeRuntime,FakeSheet,FakeCells


class Rejected(Exception):
    hresult=-2147418111  # 0x80010001
class Disconnected(Exception):
    hresult=-2147417848  # 0x80010108


class BatchWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name).resolve();self.common,self.headers,self.details=fixtures(self.root)
    def session(self,details=None):return run_comparison(self.common,self.headers,details or self.details,'openpyxl')
    def test_normal_comparison_contains_real_bitfield_evidence(self):
        s=self.session();r=next(r for r in s.results if r.status=='MISMATCH')
        self.assertEqual((r.struct,r.member,r.width,r.expected),('Device_t','active',1,'0~1'))
        self.assertTrue(r.cell);self.assertTrue(r.source_line)
        self.assertEqual([f.read_status for f in s.files],['SUCCESS','SUCCESS'])
        self.assertTrue(all(f.compared and f.extracted_variables==13 for f in s.files))
        self.assertEqual(sum(f.counts['SKIP'] for f in s.files), 2)
        self.assertEqual(sum(f.counts['MISMATCH'] for f in s.files),8);s.assert_unchanged()
    def test_empty_document_not_reported_as_variable_comparison(self):
        from openpyxl import Workbook
        p=self.root/'empty.xlsx';w=Workbook();w.save(p);w.close()
        s=self.session([p]);f=s.files[0]
        self.assertEqual((f.read_status,f.status,f.compared,f.extracted_variables),('SUCCESS','NO_VARIABLES',False,0))
        self.assertFalse(s.results)
    def test_partial_com_workbook_keeps_results_but_blocks_all_saves(self):
        from rangevalidator.readers import read_openpyxl
        from unittest.mock import patch
        def partial(path, *args, **kwargs):
            book=read_openpyxl(path)
            if str(path) in [str(p) for p in self.details]:
                book.partial=True
                book.read_errors=[dict(stage='com.Value2',api='com.Value2',code=-1,exception_type='builtins.OSError',error_class='OS_ERROR',sheet_index=1,row=9,column=2,search_phase='')]
            return book
        with patch('rangevalidator.workflow.read_com',side_effect=partial):
            session=run_comparison(self.common,self.headers,self.details,'com')
        self.assertTrue(all(f.status=='PARTIAL' and f.partial and f.compared and f.save_blocked for f in session.files))
        self.assertTrue(any(r.status=='MISMATCH' for r in session.results))
        self.assertFalse(any(r.candidate for r in session.results))
        self.assertEqual(sum(e['error_kind']=='SHEET_READ_ERROR' for e in session.errors),2)

    def test_headers_only_unique_type_name_mismatch_is_editable(self):
        session=run_comparison('',self.headers,self.details,'openpyxl')
        self.assertTrue(all(f.status=='COMPLETED' and not f.save_blocked for f in session.files))
        self.assertTrue(any(r.status=='MISMATCH' and r.member=='active' for r in session.results))
        self.assertTrue(any(r.status=='MATCH' and r.member=='mode' for r in session.results))
        self.assertTrue(any(r.status=='MISMATCH' and r.member=='active' and r.candidate for r in session.results))
        self.assertFalse(any(r.status=='MATCH' and r.candidate for r in session.results))
        reviews=[e for e in session.diagnostics.events if e.stage=='comparison.result' and e.decision=='REVIEW']
        self.assertTrue(reviews)
        self.assertTrue(all(e.sheet_name and e.cell_address and e.detail for e in reviews))

    def test_partial_reader_failure_is_file_error_not_review(self):
        bad=self.root/'broken.xlsx';bad.write_bytes(b'corrupt synthetic')
        s=self.session([self.details[0],bad])
        self.assertEqual(s.files[1].status,'ERROR');self.assertEqual(s.files[1].read_status,'FAILED')
        self.assertFalse(s.files[1].compared);self.assertEqual(s.files[1].extracted_variables,0)
        self.assertFalse(any(r.file==str(bad) for r in s.results));self.assertTrue(any(r.candidate for r in s.results))
        self.assertEqual(sum(s.files[1].counts.values()),0)
    def test_twenty_read_failures_never_become_twenty_bitfield_reviews(self):
        many=[]
        for i in range(20):
            p=self.root/f'bad-{i}.xlsx';p.write_bytes(b'synthetic nonstandard');many.append(p)
        def read(path,backend):
            e=ReadFailure('com.Find',Rejected('synthetic'));e.api='com.Find.last_column';e.context={'sheet_index':2,'search_phase':'last_column'};raise e
        with patch('rangevalidator.workflow.read_workbook',side_effect=read):
            s=self.session(many)
        self.assertEqual(len(s.files),20);self.assertFalse(s.results)
        self.assertTrue(all(f.status=='ERROR' and not f.compared for f in s.files))
        self.assertEqual([f.index for f in s.files],list(range(1,21)))
        self.assertTrue(all(f.failed_api=='com.Find.last_column' and f.sheet_index==2 for f in s.files))


class ComReadTests(unittest.TestCase):
    def test_rejected_then_success_bounded_same_thread(self):
        import threading
        ids=[];calls=[];sleeps=[];d=Diagnostics();d.set_context(file_index=7,input_role='detail',sheet_index=2)
        def op():
            ids.append(threading.get_ident());calls.append(1)
            if len(calls)<3:raise Rejected('PRIVATE_CELL_VALUE')
            return 'ok'
        self.assertEqual(read_call('com.Value2',op,d,row=4,column=3,sleep=sleeps.append,pump=lambda:None),'ok')
        self.assertEqual(sleeps,[.1,.25]);self.assertEqual(len(set(ids)),1)
        retries=[e for e in d.events if e.status=='RETRY'];self.assertEqual(len(retries),2)
        self.assertEqual((retries[0].file_index,retries[0].sheet_index,retries[0].row,retries[0].column),(7,2,4,3))
        self.assertNotIn('PRIVATE_CELL_VALUE',str(d.events))
    def test_rejected_exhausts_three_attempts(self):
        calls=[]
        def op():calls.append(1);raise Rejected('rejected')
        with self.assertRaises(ReadFailure) as ctx:read_call('com.Find.last_row',op,sleep=lambda _:None,pump=lambda:None)
        self.assertEqual(len(calls),3);self.assertEqual(ctx.exception.code&0xffffffff,0x80010001)
        self.assertEqual(ctx.exception.api,'com.Find.last_row')
    def test_disconnected_and_generic_not_retried(self):
        for error in (Disconnected('gone'),ValueError('bad')):
            calls=[]
            def op():calls.append(1);raise error
            with self.assertRaises(ReadFailure):read_call('com.Value2',op,sleep=lambda _:self.fail('must not sleep'))
            self.assertEqual(len(calls),1)
    def test_mutating_apis_cannot_retry(self):
        for name in ('com.SaveAs','com.Quit','com.Close(False)','com.Workbooks.Open','com.DispatchEx'):
            with self.assertRaises(ValueError):read_call(name,lambda:self.fail('must not call'))
    def test_empty_find_is_success_not_error(self):
        d=Diagnostics();rt=FakeRuntime(empty=True);b=read_com('synthetic.xlsx',_runtime=rt,_trace=d)
        self.assertIsNone(b.sheets[0].bounds);self.assertFalse(b.partial)
        self.assertFalse(any(e.status=='RETRY' for e in d.events))
    def test_find_failure_falls_back_to_usedrange_cap(self):
        rt=FakeRuntime();cells=rt.book.Worksheets.ws.Cells;original=cells.Find
        def find(**kwargs):
            if kwargs['SearchOrder']==2:raise Disconnected('private.xlsx')
            return original(**kwargs)
        cells.Find=find;d=Diagnostics();d.set_context(file_index=3,input_role='detail')
        b=read_com('synthetic.xlsx',_runtime=rt,_trace=d)
        self.assertEqual(b.sheets[0].cells['C3'].value,2)           # 열 Find 실패 → UsedRange 상한으로 읽음
        self.assertNotIn('private.xlsx',str(d.events))



    def test_find_rejection_recovers_in_full_reader(self):
        rt=FakeRuntime();cells=rt.book.Worksheets.ws.Cells;original=cells.Find;attempts=[]
        def find(**kw):
            if kw['SearchOrder']==2 and kw['SearchDirection']==2:
                attempts.append(1)
                if len(attempts)==1:raise Rejected('synthetic transient Find rejection')
            return original(**kw)
        cells.Find=find;d=Diagnostics();book=read_com('synthetic.xlsx',_runtime=rt,_trace=d)
        self.assertEqual(len(attempts),2);self.assertEqual(book.sheets[0].bounds,(1,1,3,3))
        self.assertTrue(any(e.status=='RETRY' and e.search_phase=='last_column' for e in d.events))
    def test_none_hresult_python_type_error_is_classified(self):
        def op():raise TypeError('synthetic bad argument')
        with self.assertRaises(ReadFailure) as ctx:read_call('com.Find.first_row',op)
        self.assertIsNone(ctx.exception.code);self.assertEqual(ctx.exception.exception_type,'builtins.TypeError')
        self.assertEqual(ctx.exception.error_class,'PYTHON_TYPE_ERROR')


    def test_com_error_args_fallback_preserves_hresult(self):
        com_error=type('com_error',(Exception,),{})
        e=ReadFailure('com.Find',com_error(-2147418111,'synthetic'))
        self.assertEqual(e.code&0xffffffff,0x80010001);self.assertEqual(e.error_class,'COM_CALL_REJECTED')
    def test_cleanup_failure_is_not_a_successful_read(self):
        rt=FakeRuntime('cleanup')
        with self.assertRaises(ReadFailure) as ctx:read_com('synthetic.xlsx',_runtime=rt)
        self.assertEqual(len(ctx.exception.cleanup_errors),2)


class BatchGuiTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        from rangevalidator.gui import App
        self.root=tk.Tk();self.root.withdraw();self.app=App(self.root)
        self.addCleanup(self.root.destroy)
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        c,h,ds=fixtures(self.tmp.name);bad=Path(self.tmp.name).resolve()/'bad.xlsx';bad.write_bytes(b'bad')
        self.session=run_comparison(c,h,[ds[0],bad],'openpyxl')
        self.app.session=self.session;self.app.refresh()
    def test_file_error_table_and_counts(self):
        self.assertEqual(len(self.app.file_tree.get_children()),2)
        self.assertEqual(len(self.app.tree.get_children()),len(self.session.results))
        failed = self.app.file_tree.get_children()[1]
        self.assertIn('ERROR', self.app.file_tree.item(failed, 'values'))
        self.app.file_tree.selection_set(failed); self.app.show_file_status()
        self.assertIn('FAILED', self.app.file_detail.get())
        self.assertFalse(self.app.tree.get_children())
    def test_file_read_error_is_not_a_bitfield_review(self):
        failed = self.app.file_tree.get_children()[1]
        self.app.file_tree.selection_set(failed); self.app.show_file_status()
        self.app.filter.set('REVIEW'); self.app.change_page(0, reset=True)
        self.assertFalse(self.app.tree.get_children())
    def test_failed_document_selection_does_not_start_writer(self):
        self.app.selected={str(Path(self.tmp.name)/'bad.xlsx')+'|error'};self.app.output.set(str(Path(self.tmp.name)/'out'))
        with patch('rangevalidator.gui.messagebox.showerror') as error,patch.object(self.app,'start') as start:
            self.app.save();error.assert_called_once();start.assert_not_called()
    def test_diagnostic_header_displays_file_sheet_phase(self):
        d=Diagnostics(lambda e:self.app.queue.put(('log',e)));d.set_context(file_index=8,input_role='detail',sheet_index=3)
        d.emit('COM','com.Find','FAILED',-2147417848,api='com.Find.last_column',search_phase='last_column',exception_type='pywintypes.com_error')
        self.app.drain();summary=self.app.diagnostic_summary.get()
        self.assertIn('#8',summary);self.assertIn('#3',summary);self.assertIn('last_column',summary)
        self.assertTrue(self.app.diagnostic_window.winfo_exists())
