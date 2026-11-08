from dataclasses import replace
import hashlib
import os
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch
from rangevalidator.models import ReadFailure, ReaderLimits
from rangevalidator.readers import read_workbook, read_com, input_path

try:
    import openpyxl
except ImportError:
    openpyxl = None


def snapshot(path):
    return hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns


def range_ends(cells, first, last):
    """Resolve both Range overloads exposed by the real Excel worksheet."""
    if not isinstance(first, str):
        return first, last
    endpoints = first.split(':')
    resolved = []
    for address in endpoints:
        letters, row = re.fullmatch(r'([A-Z]+)([0-9]+)', address).groups()
        column = 0
        for letter in letters:
            column = column * 26 + ord(letter) - 64
        resolved.append(cells.Item(int(row), column))
    return resolved[0], resolved[-1]


def fixture(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '설계 A'
    ws.append(['title'])
    ws.append(['Name', 'Type', 'Range'])
    ws.append(['p->active', 'unsigned int', '0~1'])
    ws['C4'] = '=1+1'
    ws['A5'] = 'merged'
    ws.merge_cells('A5:B5')
    ws['A6'] = '#N/A'
    ws.cell(1048576, 16384).number_format = '0.00'  # style-only far cell
    hidden = wb.create_sheet('Hidden')
    hidden.sheet_state = 'hidden'
    hidden['A1'] = 'hidden data'
    wb.create_sheet('VeryHidden').sheet_state = 'veryHidden'
    wb.save(path)
    wb.close()


class PathTests(unittest.TestCase):
    def test_invalid_paths(self):
        for path in ('', 'x\x00.xlsx', 'x.xlsm', 'x.txt'):
            with self.subTest(path=path), self.assertRaises(ReadFailure) as ctx:
                read_workbook(path)
            self.assertEqual(ctx.exception.stage, 'path')

    def test_explicit_backend(self):
        with self.assertRaises(ReadFailure) as ctx:
            read_workbook('x.xlsx', 'automatic')
        self.assertEqual(ctx.exception.stage, 'backend')

    def test_path_normalization(self):
        self.assertTrue(input_path('./합성 공백.xlsx').is_absolute())
        self.assertEqual(input_path('./합성 공백.xlsx').name, '합성 공백.xlsx')

    def test_code(self):
        e = OSError(2, 'missing')
        self.assertEqual(ReadFailure('open', e).as_dict()['code'], 2)
        self.assertNotIn('DRM', str(ReadFailure('open', e)))


@unittest.skipUnless(openpyxl, 'openpyxl not installed')
class OpenpyxlTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / '합성 공백.xlsx'
        fixture(self.path)

    def test_read_metadata_preserve(self):
        before = snapshot(self.path)
        book = read_workbook(self.path)
        ws = book.sheets[0]
        self.assertEqual(ws.bounds, (1, 1, 6, 3))
        self.assertEqual([s.state for s in book.sheets], ['visible', 'hidden', 'veryHidden'])
        self.assertTrue(ws.cells['C4'].is_formula)
        self.assertEqual(ws.cells['C4'].formula, '=1+1')
        self.assertIsNone(ws.cells['C4'].value)  # no invented cached result
        self.assertEqual(ws.cells['B5'].merged_range, 'A5:B5')
        self.assertEqual(ws.cells['B5'].merge_anchor, 'A5')
        self.assertEqual(ws.cells['A6'].kind, 'error')
        self.assertIsNone(book.sheets[2].bounds)
        self.assertEqual(before, snapshot(self.path))
        self.assertEqual(list(self.path.parent.iterdir()), [self.path])

    def test_budgets(self):
        for limits in (ReaderLimits(max_sheets=1), ReaderLimits(max_cells=1), ReaderLimits(max_text=1), ReaderLimits(max_cells=0)):
            with self.subTest(limits=limits), self.assertRaises(ReadFailure):
                read_workbook(self.path, limits=limits)

    def test_missing_corrupt_no_com_fallback(self):
        for name in ('missing.xlsx', 'corrupt.xlsx'):
            p = self.path.parent / name
            if name == 'corrupt.xlsx':
                p.write_bytes(b'not an xlsx')
            with patch('rangevalidator.readers.read_com') as com, self.assertRaises(ReadFailure) as ctx:
                read_workbook(p)
            com.assert_not_called()
            self.assertEqual(ctx.exception.stage, 'openpyxl.open')

    def test_os_denied(self):
        with patch('openpyxl.load_workbook', side_effect=PermissionError(13, 'denied')), self.assertRaises(ReadFailure) as ctx:
            read_workbook(self.path)
        self.assertEqual(ctx.exception.code, 13)
        self.assertEqual(ctx.exception.stage, 'openpyxl.open')

    def test_actual_far_data_budget(self):
        wb = openpyxl.load_workbook(self.path)
        wb.active.cell(1048576, 16384).value = 'far actual data'
        wb.save(self.path)
        wb.close()
        with self.assertRaises(ReadFailure) as ctx:
            read_workbook(self.path)
        self.assertIn('budget', ctx.exception.detail)


class ComError(Exception):
    hresult = -2146827284


class FakeCell:
    def __init__(self, row, col, value=None, formula=None, merged=False):
        self.Row, self.Column = row, col
        self.Value2 = value
        self.Formula = formula
        self.HasFormula = formula is not None
        self.MergeCells = merged

    @property
    def MergeArea(self):
        return FakeArea()


class FakeArea:
    Row, Column = 2, 1
    @property
    def Address(self):
        return 'A2:B2'


class FakeCells:
    def __init__(self, runtime):
        self.runtime = runtime
        self.finds = []

    def Item(self, row, col):
        if self.runtime.empty: return FakeCell(row,col)
        if self.runtime.fail == 'value' and row == 1 and col == 1:
            raise ComError('synthetic extraction failure')
        value = {(1, 1): 'Name', (1, 2): 'Type', (1, 3): 'Range',
                 (2, 1): 'merged', (3, 3): 2}.get((row, col))
        return FakeCell(row, col, value, '=1+1' if (row, col) == (3, 3) else None,
                        row == 2 and col in (1, 2))

    def Find(self, **kw):
        self.finds.append(kw)
        if self.runtime.fail == 'find':
            raise ComError('synthetic Find failure')
        if self.runtime.empty:
            return None
        last = kw['SearchDirection'] == 2
        return FakeCell(3 if last else 1, 3 if last else 1)


class FakeSheet:
    Name, Visible = 'Synthetic', -1
    Rows = type('Rows', (), {'Count': 1048576})()
    Columns = type('Cols', (), {'Count': 16384})()
    def __init__(self, runtime):
        self.Cells = FakeCells(runtime)
        from types import SimpleNamespace as NS
        self.UsedRange = NS(Row=1, Column=1, Rows=NS(Count=1), Columns=NS(Count=1))
    def Range(self, first, last=None):
        from types import SimpleNamespace as NS
        if isinstance(first, str) and ':' not in first:
            return range_ends(self.Cells, first, last)[0]
        first, last = range_ends(self.Cells, first, last)
        values=tuple(tuple(self.Cells.Item(r,c).Value2 for c in range(first.Column,last.Column+1)) for r in range(first.Row,last.Row+1))
        formulas=tuple(tuple(self.Cells.Item(r,c).Formula or self.Cells.Item(r,c).Value2 for c in range(first.Column,last.Column+1)) for r in range(first.Row,last.Row+1))
        return NS(Value2=values,Formula=formulas,MergeCells=None)


class FakeSheets:
    Count = 1
    def __init__(self, runtime):
        self.ws = FakeSheet(runtime)
    def Item(self, i):
        return self.ws


class FakeBook:
    ReadOnly = True
    def __init__(self, runtime):
        self.runtime = runtime
        self.Worksheets = FakeSheets(runtime)
    def Close(self, **kw):
        self.runtime.events.append(('close', kw))
        if self.runtime.fail in ('close', 'cleanup'):
            raise ComError('synthetic close failure')


class FakeBooks:
    def __init__(self, runtime):
        self.runtime = runtime
    def Open(self, **kw):
        self.runtime.events.append(('open', kw))
        if self.runtime.fail == 'open':
            raise ComError('synthetic open failure')
        return self.runtime.book


class FakeApp:
    def __init__(self, runtime):
        object.__setattr__(self, 'runtime', runtime)
        object.__setattr__(self, 'Workbooks', FakeBooks(runtime))
    def __setattr__(self, name, value):
        self.runtime.events.append((name, value))
        if self.runtime.fail == 'security' and name == 'AutomationSecurity':
            raise ComError('synthetic settings failure')
        object.__setattr__(self, name, value)
    def Quit(self):
        self.runtime.events.append(('quit',))
        if self.runtime.fail in ('quit', 'cleanup'):
            raise ComError('synthetic quit failure')


class FakeRuntime:
    def __init__(self, fail=None, empty=False):
        self.fail, self.empty, self.events = fail, empty, []
        self.book = FakeBook(self)
        self.app = FakeApp(self)
    def initialize(self):
        self.events.append(('init',))
        if self.fail == 'sta':
            raise ComError('MTA already initialized')
    def uninitialize(self):
        self.events.append(('uninit',))
    def process_ids(self):
        return {10}
    def dispatch(self):
        self.events.append(('dispatch',))
        if self.fail == 'dispatch':
            raise ComError('no Excel')
        return self.app
    def wait_exit(self, pid):
        if self.fail == 'exit':
            raise TimeoutError('synthetic process timeout')

    def pid(self, app):
        return 10 if self.fail == 'ownership' else 20


class ComContractTests(unittest.TestCase):
    def test_success_contract(self):
        rt = FakeRuntime()
        book = read_com('synthetic.xlsx', _runtime=rt)
        names = [e[0] for e in rt.events]
        self.assertLess(names.index('AutomationSecurity'), names.index('open'))
        kw = next(e[1] for e in rt.events if e[0] == 'open')
        self.assertEqual({k: kw[k] for k in ('ReadOnly', 'UpdateLinks', 'Notify', 'AddToMru')}, dict(ReadOnly=True, UpdateLinks=0, Notify=False, AddToMru=False))
        self.assertTrue(Path(kw['Filename']).is_absolute())
        self.assertEqual(rt.app.AutomationSecurity, 3)
        self.assertFalse(rt.app.ScreenUpdating)
        self.assertFalse(rt.app.EnableEvents)
        self.assertEqual(rt.events[-3:], [('close', {'SaveChanges': False}), ('quit',), ('uninit',)])
        ws = book.sheets[0]
        self.assertEqual(ws.bounds, (1, 1, 3, 3))
        self.assertEqual(ws.cells['C3'].value, 2)
        self.assertEqual(ws.cells['C3'].formula, '=1+1')
        self.assertEqual(ws.cells['B2'].merge_anchor, 'A2')
        self.assertEqual(ws.merged_ranges, ['A2:B2'])
        finds = rt.book.Worksheets.ws.Cells.finds
        self.assertEqual(len(finds), 2)
        self.assertEqual([(f['SearchOrder'], f['SearchDirection']) for f in finds], [(1,2), (2,2)])
        self.assertTrue(all(f['LookIn'] == -4123 and f['LookAt'] == 2 and 'SearchFormat' not in f for f in finds))

    def test_error_stages_and_cleanup(self):
        cases = {'sta': 'com.STA.initialize', 'dispatch': 'com.DispatchEx',
                 'ownership': 'com.instance.verify', 'security': 'com.security',
                 'open': 'com.Workbooks.Open', 'find': 'com.Find'}
        for mode, stage in cases.items():
            with self.subTest(mode=mode):
                rt = FakeRuntime(mode)
                if mode in ('find','ownership','security'):
                    partial=read_com('synthetic.xlsx', _runtime=rt)
                    self.assertFalse(partial.partial)
                    continue
                with self.assertRaises(ReadFailure) as ctx:
                    read_com('synthetic.xlsx', _runtime=rt)
                self.assertEqual(ctx.exception.stage, stage)
                names = [e[0] for e in rt.events]
                self.assertEqual('uninit' in names, mode != 'sta')
                self.assertEqual('quit' in names, mode in ('security', 'open', 'find'))
                if mode not in ('ownership','find'):
                    self.assertEqual(ctx.exception.as_dict()['code_hex'], '0x800A03EC')
                if mode == 'ownership':
                    self.assertNotIn('open', names)
                    self.assertNotIn('AutomationSecurity', names)

    def test_cleanup_failures_independent(self):
        rt = FakeRuntime('cleanup')
        with self.assertRaises(ReadFailure) as ctx:
            read_com('synthetic.xlsx', _runtime=rt)
        self.assertEqual([e['stage'] for e in ctx.exception.cleanup_errors], ['com.Close(False)', 'com.Quit'])
        self.assertEqual(rt.events[-1], ('uninit',))

    def test_empty_and_budget(self):
        self.assertIsNone(read_com('empty.xlsx', _runtime=FakeRuntime(empty=True)).sheets[0].bounds)
        rt = FakeRuntime()
        book=read_com('x.xlsx', ReaderLimits(max_cells=1), _runtime=rt)
        self.assertTrue(book.sheets)
        self.assertIn(('quit',), rt.events)

    def test_extraction_failure_marks_partial(self):
        rt = FakeRuntime('value')
        partial=read_com('x.xlsx', _runtime=rt)
        self.assertTrue(partial.partial)
        self.assertTrue(partial.read_errors)
        self.assertIn(('close', {'SaveChanges': False}), rt.events)


@unittest.skipUnless(openpyxl and os.environ.get('RV_TEST_EXCEL') == '1', 'Explicit RV_TEST_EXCEL=1 and Excel/pywin32 required')
class ExcelIntegrationTests(unittest.TestCase):
    def test_real_excel_synthetic_readonly(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / '합성 공백.xlsx'
            fixture(p)
            before = snapshot(p)
            book = read_com(p)
            self.assertEqual([s.state for s in book.sheets], ['visible', 'hidden', 'veryHidden'])
            self.assertEqual(book.sheets[0].bounds, (1, 1, 6, 3))
            self.assertTrue(book.sheets[0].cells['C4'].is_formula)
            self.assertEqual(book.sheets[0].cells['B5'].merge_anchor, 'A5')
            self.assertEqual(before, snapshot(p))
            self.assertEqual(list(Path(folder).iterdir()), [p])



@unittest.skipUnless(openpyxl and os.environ.get('RV_TEST_EXCEL') == '1', 'Explicit Excel integration opt-in required')
class ExcelLifecycleTests(unittest.TestCase):
    def test_existing_instance_preserved_and_owned_exits(self):
        import gc
        import pythoncom
        import win32com.client
        from rangevalidator.readers import Win32Runtime
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        existing = None
        try:
            existing = win32com.client.DispatchEx('Excel.Application')
            existing.Visible = False
            existing.DisplayAlerts = True
            existing.ScreenUpdating = True
            old_hwnd = existing.Hwnd
            with tempfile.TemporaryDirectory() as folder:
                p = Path(folder) / 'synthetic.xlsx'
                fixture(p)
                read_com(p)
            # Access still works and the Reader did not alter these settings.
            self.assertEqual(existing.Hwnd, old_hwnd)
            self.assertTrue(existing.DisplayAlerts)
            self.assertTrue(existing.ScreenUpdating)
        finally:
            if existing is not None:
                pid = Win32Runtime().pid(existing)
                existing.Quit()
                existing = None
                gc.collect()
                Win32Runtime().wait_exit(pid)
            pythoncom.CoUninitialize()

    def test_missing_corrupt_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            corrupt = Path(folder) / '손상 합성.xlsx'
            corrupt.write_bytes(b'Synthetic corrupt workbook')
            for p in (corrupt, Path(folder) / '없는 폴더' / 'missing.xlsx'):
                before = snapshot(corrupt)
                with self.subTest(path=p), self.assertRaises(ReadFailure) as ctx:
                    read_com(p)
                self.assertEqual(ctx.exception.stage, 'com.Workbooks.Open')
                self.assertIsInstance(ctx.exception.code, int)
                self.assertFalse(ctx.exception.cleanup_errors)
                self.assertEqual(before, snapshot(corrupt))

    def test_mta_refused(self):
        import threading
        import pythoncom
        results = []
        def worker():
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            try:
                try:
                    read_com('synthetic.xlsx')
                except ReadFailure as exc:
                    results.append(exc)
            finally:
                pythoncom.CoUninitialize()
        thread = threading.Thread(target=worker)
        thread.start()
        thread.join(20)
        self.assertFalse(thread.is_alive())
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].stage, 'com.Workbooks.Open')
        self.assertIsInstance(results[0].code,int)

    def test_shared_value_semantics(self):
        from datetime import datetime, time
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / 'values.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active
            for i, value in enumerate([True, 3, 1.5, 'text', '#N/A', datetime(2026, 1, 2), time(12, 0)], 1):
                ws.cell(i, 1).value = value
            wb.save(p)
            wb.close()
            a, b = read_workbook(p).sheets[0], read_com(p).sheets[0]
            self.assertEqual([(c.value, c.kind) for c in a.cells.values()], [(c.value, c.kind) for c in b.cells.values()])

class ComCleanupExtraTests(unittest.TestCase):
    def test_owned_process_timeout_reported(self):
        rt = FakeRuntime('exit')
        with self.assertRaises(ReadFailure) as ctx:
            read_com('x.xlsx', _runtime=rt)
        self.assertEqual([e['stage'] for e in ctx.exception.cleanup_errors], ['com.process.exit'])
        self.assertIn(('quit',), rt.events)
        self.assertIn(('uninit',), rt.events)

    def test_primary_failure_retains_cleanup_error(self):
        rt = FakeRuntime('find')
        def broken_quit():
            raise ComError('synthetic cleanup')
        object.__setattr__(rt.app, 'Quit', broken_quit)
        with self.assertRaises(ReadFailure) as ctx:
            read_com('x.xlsx', _runtime=rt)
        self.assertEqual(ctx.exception.stage, 'cleanup')
        self.assertEqual([e['stage'] for e in ctx.exception.cleanup_errors], ['com.Quit'])

    def test_path_validation_before_dispatch(self):
        rt = FakeRuntime()
        with self.assertRaises(ReadFailure) as ctx:
            read_com('bad.xls', _runtime=rt)
        self.assertEqual(ctx.exception.stage, 'path')
        self.assertFalse(rt.events)
