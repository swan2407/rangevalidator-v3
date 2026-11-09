from types import SimpleNamespace as NS
import unittest
from unittest.mock import patch
from rangevalidator.readers import read_com
from rangevalidator.models import Sheet
from rangevalidator.extraction_diagnostics import diagnose_workbook
from rangevalidator.engine import is_function_sheet
from test_readers import FakeRuntime


class NameOnly:
    def __init__(self, name, events, fail=False):
        self.name, self.events, self.fail = name, events, fail
    @property
    def Name(self):
        self.events.append(('name', self.name))
        if self.fail: raise RuntimeError('name denied')
        return self.name
    def __getattr__(self, attr):
        raise AssertionError('Unexpected body access: ' + attr)


class SelectionTests(unittest.TestCase):
    def run_reader(self, names, failing_body='', failing_name=''):
        events = []
        sheets = [NameOnly(n, events, n == failing_name) for n in names]
        rt = FakeRuntime()
        rt.book.Worksheets = NS(Count=len(sheets), Item=lambda i:sheets[i-1])
        def copy(ws, *args):
            events.append(('body', ws.name))
            self.assertEqual(sum(e[0]=='name' for e in events), len(names))
            self.assertTrue(is_function_sheet(ws.name))
            if ws.name == failing_body: raise RuntimeError('body denied')
            return Sheet(ws.name, 'visible'), 0, 0, []
        with patch('rangevalidator.com_values._copy_sheet', side_effect=copy):
            book = read_com('synthetic.xlsx', _runtime=rt, function_sheets_only=True)
        return book, events

    def test_inventory_before_body_and_126_sheets(self):
        excluded = ['표지','목차','개요','구현파일관계','Common Parameter','commonparameter']
        included = [f'{i}.Function' for i in range(1,121)]
        book, events = self.run_reader(excluded + included)
        self.assertEqual([n for kind,n in events if kind=='body'], included)
        self.assertFalse(book.partial)
        report = diagnose_workbook(book)
        self.assertEqual(report['total_sheets'],126)
        self.assertEqual(report['counts']['excluded_sheets'],6)
        self.assertEqual(report['counts']['body_read_success'],120)
        self.assertEqual(book.read_errors,[])

    def test_body_failure_is_named_and_other_function_survives(self):
        book, _ = self.run_reader(['commonparameter','1.Bad','2.Good'], failing_body='1.Bad')
        self.assertTrue(book.partial)
        self.assertEqual([s.name for s in book.sheets],['2.Good'])
        self.assertEqual([e['sheet_name'] for e in book.read_errors],['1.Bad'])

    def test_name_failure_has_unknown_scope(self):
        book, _ = self.run_reader(['unknown','1.Good'], failing_name='unknown')
        self.assertTrue(book.partial)
        self.assertEqual(book.read_errors[0]['sheet_name'],'')
        self.assertEqual([s.name for s in book.sheets],['1.Good'])

    def test_only_excluded_is_not_partial(self):
        book, events = self.run_reader(['표지','commonparameter'])
        self.assertEqual(book.sheets,[])
        self.assertFalse(book.partial)
        self.assertFalse(any(k=='body' for k,_ in events))

    def test_only_functions(self):
        included = ['1.Main', 'update1', '1_update', '1', '1.', 'CAN2.Update']
        book, events = self.run_reader(['Common Parameter', *included, '목차'])
        self.assertEqual([s.name for s in book.sheets], included)
        self.assertEqual([name for kind, name in events if kind == 'body'], included)

