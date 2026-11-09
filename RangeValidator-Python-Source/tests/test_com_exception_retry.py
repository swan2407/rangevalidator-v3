"""Bounded 0x80020009 recovery and final PARTIAL behavior in the actual reader path."""
from types import SimpleNamespace
import unittest

from rangevalidator.com_read import read_call
from rangevalidator.diagnostics import Diagnostics
from rangevalidator.models import ReadFailure
from rangevalidator.readers import read_com
from test_readers import FakeRuntime, FakeSheet


class DispatchException(Exception):
    hresult = -2147352567  # DISP_E_EXCEPTION / 0x80020009
    excepinfo = (0, 'Excel', 'synthetic', None, 0, -2146827284)


class Fault:
    def __init__(self, failures=2):
        self.failures = failures
        self.calls = 0

    def check(self):
        self.calls += 1
        if self.failures is None or self.calls <= self.failures:
            raise DispatchException('synthetic read failure')


class FaultProxy:
    def __init__(self, target, attribute, fault):
        self.target, self.attribute, self.fault = target, attribute, fault

    def __getattr__(self, name):
        if name == self.attribute:
            self.fault.check()
        value = getattr(self.target, name)
        if name == 'MergeArea' and self.attribute == 'Address':
            return FaultProxy(value, 'Address', self.fault)
        return value


def fault_ranges(sheet, attribute, fault):
    original = sheet.Range
    sheet.Range = lambda *args: FaultProxy(original(*args), attribute, fault)
    return sheet


class ComExceptionRetryTests(unittest.TestCase):
    def test_success_on_third_attempt_is_logged_even_after_prior_success(self):
        fault, diagnostics, sleeps = Fault(), Diagnostics(), []
        read_call('com.Value2', lambda: 'previous block', diagnostics)

        def operation():
            fault.check()
            return 'third read'

        result = read_call('com.Value2', operation, diagnostics, row=4, column=3,
                           sleep=sleeps.append, pump=lambda: None)
        self.assertEqual((result, fault.calls, sleeps), ('third read', 3, [.1, .25]))
        retries = [e for e in diagnostics.events if e.status == 'RETRY']
        self.assertEqual([e.attempt for e in retries], [1, 2])
        self.assertTrue(all(e.hresult == '0x80020009' and e.max_attempts == 3 for e in retries))
        self.assertTrue(all(e.inner_hresult == -2146827284 for e in retries))
        self.assertTrue(any(e.status == 'OK' and e.attempt == 3 for e in diagnostics.events))
        self.assertFalse(any(e.status == 'FAILED' for e in diagnostics.events))

    def test_third_failure_preserves_code_api_and_attempt_limit(self):
        fault = Fault(None)
        with self.assertRaises(ReadFailure) as raised:
            read_call('com.formula.text', fault.check, sleep=lambda _: None, pump=lambda: None)
        error = raised.exception
        self.assertEqual(fault.calls, 3)
        self.assertEqual(error.code & 0xffffffff, 0x80020009)
        self.assertEqual(error.inner_hresult, -2146827284)
        self.assertEqual(error.api, 'com.formula.text')
        self.assertEqual((error.context['attempt'], error.context['max_attempts']), (3, 3))

    def test_values_formulas_and_merge_information_recover_in_full_reader(self):
        for attribute, api in [('Value2', 'com.Value2'), ('Formula', 'com.formula.text'),
                               ('MergeCells', 'com.merge.flag'), ('Address', 'com.merge.address')]:
            with self.subTest(attribute=attribute):
                runtime, fault, diagnostics = FakeRuntime(), Fault(), Diagnostics()
                fault_ranges(runtime.book.Worksheets.ws, attribute, fault)
                book = read_com('synthetic.xlsx', _runtime=runtime, _trace=diagnostics)
                self.assertFalse(book.partial)
                self.assertEqual(book.read_errors, [])
                self.assertEqual(book.sheets[0].cells['C3'].value, 2)
                self.assertEqual(book.sheets[0].cells['B2'].merge_anchor, 'A2')
                retries = [e for e in diagnostics.events if e.status == 'RETRY']
                self.assertEqual([(e.api, e.attempt) for e in retries], [(api, 1), (api, 2)])
                self.assertTrue(any(e.api == api and e.status == 'OK' and e.attempt == 3
                                    for e in diagnostics.events))

    def test_sheet_metadata_reads_recover(self):
        for attribute in ('Name', 'Visible', 'Count'):
            with self.subTest(attribute=attribute):
                runtime, fault, diagnostics = FakeRuntime(), Fault(), Diagnostics()
                sheets = runtime.book.Worksheets
                if attribute == 'Count':
                    runtime.book.Worksheets = FaultProxy(sheets, attribute, fault)
                else:
                    sheets.ws = FaultProxy(sheets.ws, attribute, fault)
                book = read_com('synthetic.xlsx', _runtime=runtime, _trace=diagnostics)
                self.assertFalse(book.partial)
                self.assertEqual(len(book.sheets), 1)
                self.assertEqual(len([e for e in diagnostics.events if e.status == 'RETRY']), 2)

    def test_persistent_failure_marks_only_failed_sheet_and_continues(self):
        runtime, fault, diagnostics = FakeRuntime(), Fault(None), Diagnostics()
        bad = fault_ranges(FakeSheet(runtime), 'Value2', fault)
        bad.Name = '1.Bad'
        good = FakeSheet(runtime)
        good.Name = '2.Good'
        runtime.book.Worksheets = SimpleNamespace(Count=2, Item=lambda n: (bad, good)[n - 1])
        book = read_com('synthetic.xlsx', _runtime=runtime, _trace=diagnostics, function_sheets_only=True)
        self.assertTrue(book.partial)
        self.assertEqual(fault.calls, 3)
        self.assertEqual([s.name for s in book.sheets], ['2.Good'])
        self.assertEqual(len(book.read_errors), 1)
        error = book.read_errors[0]
        self.assertEqual((error['sheet_name'], error['api'], error['attempt'], error['max_attempts']),
                         ('1.Bad', 'com.Value2', 3, 3))
        self.assertEqual(error['code'] & 0xffffffff, 0x80020009)
        self.assertEqual(error['inner_hresult'], -2146827284)
        events = [event[0] for event in runtime.events]
        self.assertEqual((events.count('open'), events.count('close'), events.count('quit')), (1, 1, 1))
