"""COM boundary regressions with deterministic worksheet doubles."""
from types import SimpleNamespace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from rangevalidator.com_values import _copy_sheet, _merge_map, _read_blocks
from rangevalidator.models import ReadFailure, ReaderLimits
from rangevalidator.readers import read_com
from rangevalidator.workflow import run_comparison
from test_com_reference import ComError, Sheet, runtime
from test_excel6_reader import MergeSheet


class Excel5LimitTests(unittest.TestCase):
    def test_sparse_release_budget_counts_content_and_preserves_far_merge(self):
        source = Sheet({(1, 1): 'Name', (1, 2): 'Type', (1, 3): 'Range',
                        (2, 1): 'obj.state', (2, 3): '0~31', (5001, 2): 'note'})
        merges = MergeSheet()
        merges.areas = [(5001, 2, 5001, 3, '$B$5001:$C$5001')]
        value_range = source.Range
        source.Range = lambda first, last=None: (
            merges.Range(first) if isinstance(first, str) else value_range(first, last))
        with patch('rangevalidator.com_values._count_a',
                   side_effect=lambda ws, first, last, col: int(first == 1 or last == 5001)):
            sheet, cells, _, _ = _copy_sheet(source, ReaderLimits(max_cells=32), 32,
                                              merge_limit=64, sparse=True)
        self.assertLess(cells, 32)
        self.assertEqual(sheet.cells['C2'].value, '0~31')
        self.assertEqual(sheet.cells['C5001'].merge_anchor, 'B5001')
        self.assertEqual(sheet.merged_ranges, ['B5001:C5001'])

    def test_real_content_after_two_empty_blocks_is_not_dropped(self):
        tail = 393217
        sheet = Sheet({(1, 1): 'first', (tail, 1): 'last'})
        with patch('rangevalidator.com_values._count_a',
                   side_effect=lambda ws, first, last, col: int(first == 1 or last == tail)):
            values, _, last_row = _read_blocks(sheet, tail, 1)
        self.assertEqual(last_row, tail)
        self.assertEqual(values[tail], ['last'])

    def test_wide_sheet_is_partial_and_other_function_sheet_survives(self):
        wide = Sheet({(3, 257): 'outside'})
        wide.Name = '1.Wide'
        good = Sheet()
        good.Name = '2.Good'
        rt = runtime()
        rt.book.Worksheets = SimpleNamespace(Count=2, Item=lambda n: (wide, good)[n-1])
        book = read_com('x.xlsx', _runtime=rt, function_sheets_only=True)
        self.assertTrue(book.partial)
        self.assertEqual([sheet.name for sheet in book.sheets if sheet.cells], ['2.Good'])
        error, = book.read_errors
        self.assertEqual((error['stage'], error['sheet_name']), ('com.bounds.limit', '1.Wide'))
        self.assertEqual(error['attempted_bounds'], (1, 1, 3, 257))

    def test_cell_budget_omits_evidence_and_reports_failed_sheet(self):
        book = read_com('x.xlsx', ReaderLimits(max_cells=8), _runtime=runtime())
        self.assertTrue(book.partial)
        self.assertFalse(any(sheet.cells for sheet in book.sheets))
        error, = book.read_errors
        self.assertEqual((error['stage'], error['attempted_bounds']), ('com.bounds.limit', (1, 1, 3, 3)))
        self.assertNotIn('SHEET_BODY_READ_COMPLETED', book.sheet_observations[0]['states'])

    def test_short_matrix_and_value2_hresult_are_sheet_local(self):
        class FaultyRange:
            @property
            def Value2(self):
                raise ComError('synthetic Value2 failure')

        broken = Sheet()
        broken.Name = '1.Value2_denied'
        broken.Range = lambda *_: FaultyRange()
        short = Sheet()
        short.Name = '2.Short_matrix'
        short.Range = lambda *_: SimpleNamespace(Value2=(('Name',),))
        valid = Sheet()
        valid.Name = '3.Normal'
        rt = runtime()
        rt.book.Worksheets = SimpleNamespace(Count=3, Item=lambda n: (broken, short, valid)[n-1])
        book = read_com('x.xlsx', _runtime=rt, function_sheets_only=True)
        self.assertTrue(book.partial)
        self.assertEqual([s.name for s in book.sheets], ['3.Normal'])
        denied, incomplete = book.read_errors
        self.assertEqual((denied['api'], denied['sheet_name'], denied['code'] & 0xffffffff),
                         ('com.Value2', '1.Value2_denied', 0x800A03EC))
        self.assertEqual((incomplete['api'], incomplete['sheet_name'], incomplete['attempted_bounds']),
                         ('com.Value2', '2.Short_matrix', (1, 1, 3, 3)))

    def test_fallback_clipping_reports_affected_bounds(self):
        sheet = Sheet({(3, 300): 'outside'}, fail_find=(2,))
        sheet.Name = '1.Fallback_wide'
        book = read_com('x.xlsx', _runtime=runtime(sheet), function_sheets_only=True)
        self.assertTrue(book.partial)
        error, = book.read_errors
        self.assertEqual(error['attempted_bounds'], (1, 1, 3, 300))
        self.assertIn('fallback', error['limit_reason'])

    def test_merge_budget_blocks_uncertain_sheet_only_when_merges_may_exist(self):
        sheet = Sheet()
        with patch('rangevalidator.com_values._merge_state', return_value=False):
            self.assertEqual(_merge_map(sheet, 200001, 1), ({}, set()))
        with patch('rangevalidator.com_values._merge_state', return_value=None):
            with self.assertRaises(ReadFailure) as result:
                _merge_map(sheet, 200001, 1)
        self.assertEqual(result.exception.context['attempted_bounds'], (1, 1, 200001, 1))
        self.assertIn('Merge scan budget', result.exception.context['limit_reason'])

    def test_open_access_denial_and_hresult_are_isolated_by_file(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'types.h'
            source.write_text('typedef struct { unsigned int v001:3; } T; extern T obj;', encoding='utf-8')
            paths = [root / name for name in ('denied.xlsx', 'refused.xlsx', 'valid.xlsx')]
            sheet = Sheet({(1, 1): 'Name', (1, 2): 'Type', (1, 3): 'Range',
                           (2, 1): 'obj.v001', (2, 2): 'unsigned int', (2, 3): '0~3'})
            sheet.Name = '1.Good'

            def fake_reader(path, **kwargs):
                rt = runtime(sheet)
                failure = {'denied.xlsx': PermissionError(5, 'synthetic permission denial'),
                           'refused.xlsx': ComError('synthetic Open failure')}.get(Path(path).name)
                if failure:
                    def open_failed(**_):
                        raise failure
                    rt.app.Workbooks.Open = open_failed
                return read_com(path, _runtime=rt, **kwargs)

            with patch('rangevalidator.workflow.read_com', side_effect=fake_reader):
                session = run_comparison(None, [source], paths, backend='com')
            self.assertEqual([f.status for f in session.files], ['ERROR', 'ERROR', 'COMPLETED'])
            self.assertEqual([(Path(e['file']).name, e['stage'], e['code'] & 0xffffffff) for e in session.errors],
                             [('denied.xlsx', 'com.Workbooks.Open', 5),
                              ('refused.xlsx', 'com.Workbooks.Open', 0x800A03EC)])
            result, = session.results
            self.assertEqual((Path(result.file).name, result.sheet, result.cell, result.candidate),
                             ('valid.xlsx', '1.Good', 'C2', True))
