from contextlib import contextmanager
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rangevalidator import writer
from rangevalidator.engine import Result
from rangevalidator.manual_edits import ReviewRequest
from rangevalidator.workflow import FileStatus, Session


class ManualWriterTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name).resolve()
        self.source = self.root / 'design.xlsx'
        self.source.write_bytes(b'original')
        self.auto = Result('auto', str(self.source), '1.Main', 'B2', 'obj.mode',
                           'MISMATCH', '', existing='0~0', expected='0~3', candidate=True)
        self.review = replace(self.auto, id='review', cell='B3', status='REVIEW',
                              expected='', reason='타입 불명', candidate=False)
        self.session = Session((), (str(self.source),), diagnostics=Mock(),
                               files=[FileStatus(1, str(self.source), status='COMPLETED', read_status='SUCCESS')])

    def save(self, requests, rows, auto=True):
        self.session.results = rows
        self.session._signature = self.session.signature()
        cells = {row.cell: SimpleNamespace(HasFormula=False, MergeCells=False, Value2=row.existing)
                 for row in rows}
        self.cells = cells
        self.opens = []
        self.save_as = Mock(side_effect=lambda **kwargs: Path(kwargs['Filename']).write_bytes(b'saved'))

        @contextmanager
        def lock(_):
            yield

        @contextmanager
        def open_book(path, _):
            self.opens.append(Path(path))
            sheet = SimpleNamespace(Range=cells.__getitem__)
            yield SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda _: sheet),
                                  SaveAs=self.save_as), Mock()

        with patch.object(writer, 'input_locks', lock), patch.object(writer, 'excel_book', open_book), \
                patch.object(writer, 'create_report'):
            return writer.save_selected(self.session, ['auto'] if auto else [], self.root / 'out',
                                        manual_requests=requests)[1]

    def test_automatic_and_manual_share_save_without_changing_analysis(self):
        unselected = replace(self.review, id='unselected', cell='B4')
        rows = [self.auto, self.review, unselected]
        result = self.save([ReviewRequest('review', '[-1, 7]', True),
                            ReviewRequest('unselected', '0~7')], rows)
        self.assertEqual({key: item['edit_kind'] for key, item in result.items()},
                         {'auto': 'AUTO_MISMATCH', 'review': 'MANUAL_REVIEW'})
        self.assertEqual({item['status'] for item in result.values()}, {'SAVED_VERIFIED'})
        self.assertEqual(self.cells['B3'].Value2, '-1~7')
        self.assertEqual(self.cells['B4'].Value2, unselected.existing)
        self.assertEqual(self.opens[-1].name, 'design_fixed.xlsx')
        self.save_as.assert_called_once()
        self.assertEqual(self.session.results, rows)
        self.assertEqual(self.review.status, 'REVIEW')
        self.assertEqual(self.source.read_bytes(), b'original')

    def test_manual_only_save_does_not_require_automatic_candidates(self):
        self.session.files[0].save_blocked = True
        outcomes = self.save([ReviewRequest('review', '0~7', True)], [self.review], auto=False)
        self.assertEqual(outcomes['review']['status'], 'SAVED_VERIFIED')
        self.save_as.assert_called_once()

    def test_invalid_empty_noop_and_duplicate_requests_do_not_block_automatic(self):
        variations = [('invalid', 'B3', '=3'), ('empty', 'B4', ''),
                      ('noop', 'B5', '[0, 0]'), ('duplicate-auto', 'B2', '0~7'),
                      ('duplicate-a', 'B6', '0~7'), ('duplicate-b', 'B6', '0~8')]
        rows = [self.auto, *(replace(self.review, id=rid, cell=cell) for rid, cell, _ in variations)]
        result = self.save([ReviewRequest(rid, value, True) for rid, _, value in variations], rows)
        self.assertEqual(result['auto']['status'], 'SAVED_VERIFIED')
        self.assertTrue(all(result[rid]['status'] == 'SKIPPED' for rid, _, _ in variations))
        self.save_as.assert_called_once()

    def test_nonfunction_foreign_and_partial_requests_cannot_write(self):
        self.session.files[0].partial = True
        self.session.errors = [dict(file=str(self.source), sheet_name='1.Main', error_kind='SHEET_READ_ERROR')]
        nonfunction = replace(self.review, id='cover', sheet='cover')
        result = self.save([ReviewRequest(i, '0~7', True) for i in ('review', 'foreign', 'cover')],
                           [self.review, nonfunction], auto=False)
        self.assertEqual(set(result), {'review'})
        self.assertEqual(result['review']['status'], 'SKIPPED')
        self.assertEqual(self.opens, [])
        self.save_as.assert_not_called()

    def test_manual_preflight_still_checks_formula_merge_and_old_value(self):
        for variant in ('formula', 'merge', 'old_value'):
            with self.subTest(variant=variant):
                cell = SimpleNamespace(HasFormula=variant == 'formula', MergeCells=variant == 'merge',
                                       Value2='changed' if variant == 'old_value' else self.review.existing,
                                       MergeArea=SimpleNamespace(Cells=SimpleNamespace(
                                           Item=lambda *_: SimpleNamespace(Address='$A$3'))))
                book = SimpleNamespace(Worksheets=SimpleNamespace(
                    Item=lambda _: SimpleNamespace(Range=lambda _: cell)), SaveAs=Mock())

                @contextmanager
                def excel_book(*_):
                    yield book, Mock()

                edit = writer.ManualEdit.from_result(self.review, '0~7')
                outcomes = {}
                with patch.object(writer, 'excel_book', excel_book):
                    writer._apply_file(self.source, self.root / 'design_fixed.xlsx', [edit], Mock(), outcomes)
                self.assertEqual(outcomes[edit.id]['status'], 'SKIPPED')
                book.SaveAs.assert_not_called()
