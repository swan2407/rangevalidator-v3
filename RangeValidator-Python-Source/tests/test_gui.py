import os,tempfile,time,unittest
from pathlib import Path
from unittest.mock import patch
from v3_fixtures import fixtures


@unittest.skipUnless(os.environ.get('RV_TEST_GUI')=='1','Tk GUI opt-in required')
class GuiTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        from rangevalidator.gui import App
        self.root=tk.Tk();self.root.withdraw();self.app=App(self.root)
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        def close():
            try:self.root.destroy()
            except tk.TclError:pass
        self.addCleanup(close)
        c,h,d=fixtures(self.tmp.name)
        self.app.common=str(c);self.app.headers=list(map(str,h));self.app.details=list(map(str,d));self.app.sync_inputs()
    def pump(self,predicate,timeout=30):
        deadline=time.monotonic()+timeout
        while time.monotonic()<deadline and not predicate():
            self.root.update();time.sleep(.02)
        self.assertTrue(predicate(),'UI worker timed out')
    def test_async_compare_selection_and_invalidation(self):
        self.assertEqual(len(self.app.notebook.tabs()),4)
        self.app.compare();self.assertTrue(self.app.busy)
        self.pump(lambda:not self.app.busy)
        self.assertIsNotNone(self.app.session)
        self.app.select_visible();self.assertEqual(len(self.app.selected),5)
        self.app.filter.set('REVIEW');self.app.refresh()
        previous=set(self.app.selected)
        self.app.select_visible();self.assertEqual(previous,self.app.selected)
        self.app.clear('headers');self.assertIsNone(self.app.session);self.assertFalse(self.app.selected)
    def test_com_failure_auto_log_and_masking(self):
        from rangevalidator.diagnostics import Diagnostics
        d=Diagnostics(lambda e:self.app.queue.put(('log',e)))
        d.emit('COM','com.Workbooks.Open')
        d.emit('COM','com.Workbooks.Open','FAILED',-2146827284)
        self.root.update();self.app.drain()
        self.assertEqual(self.app.notebook.select(),str(self.app.tabs[3]))
        content=self.app.log.get('1.0','end')
        self.assertIn('800A03EC',content)
        self.assertNotIn(str(self.tmp.name),content)
        self.assertIn('open_entered',content)
    def test_close_waits_for_worker_cleanup(self):
        self.app.busy=True;self.app.close();self.assertTrue(self.app.closing);self.assertTrue(self.root.winfo_exists())
        self.app.busy=False
    def test_report_only_worker(self):
        self.app.compare();self.pump(lambda:not self.app.busy)
        out=Path(self.tmp.name)/'report.xlsx'
        with patch('rangevalidator.gui.filedialog.asksaveasfilename',return_value=str(out)):
            self.app.report_only()
        self.pump(lambda:not self.app.busy)
        self.assertTrue(out.is_file())

    def test_pagination_preserves_access_to_all_results(self):
        from dataclasses import replace
        self.app.compare();self.pump(lambda:not self.app.busy)
        base=self.app.session.results[0]
        self.app.session.results=[replace(base,id=f'synthetic-{i}',cell=f'C{i+1}') for i in range(2005)]
        self.app.refresh();self.assertEqual(len(self.app.tree.get_children()),1000)
        self.app.change_page(1);self.assertEqual(len(self.app.tree.get_children()),1000)
        self.app.change_page(1);self.assertEqual(len(self.app.tree.get_children()),5)
        self.assertEqual(self.app.encoding.get(),'utf-8-sig')

    def test_ordinary_member_filter_and_non_editable_label(self):
        self.app.compare();self.pump(lambda:not self.app.busy)
        self.app.filter.set('일반 멤버');self.app.refresh()
        rows = self.app.tree.get_children()
        self.assertEqual(len(rows), 1)
        values = self.app.tree.item(rows[0], 'values')
        self.assertEqual(values[0], '')
        self.assertEqual(values[1], '일반 멤버 (SKIP)')
        self.assertIn('일반 멤버 2', self.app.count.get())
        self.app.select_visible()
        self.assertEqual(self.app.selected, set())
