from contextlib import contextmanager
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.manual_edits import ReviewRequest
from rangevalidator.models import Workbook
from rangevalidator.workflow import FileStatus, Session
from rangevalidator import writer
from test_function_sheets import sheet


class CreateTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name).resolve()
        source = self.root / 'types.h'
        source.write_text('typedef struct { unsigned int state:4; int normal; char sign:2; } T; T obj;')
        self.catalog = catalog_headers([source])
        self.path = self.root / 'design.xlsx'
        self.path.write_bytes(b'original')

    def book(self, value, expression='obj.state'):
        return Workbook(str(self.path), 'com', [
            sheet('1.Main', [('Name', 'Type', 'Range'), (expression, '', value)])])

    def test_blank_forms_create_and_nonblank_decisions_are_preserved(self):
        for value in (None, '', ' \t '):
            results, _ = compare_book(None, self.catalog, self.book(value))
            row = results[0]
            self.assertEqual((row.status, row.expected, row.width, row.signed, row.candidate),
                             ('CREATE', '0~15', 4, False, True))
            self.assertEqual(row.existing, value)
        for value, status in [('0~15', 'MATCH'), ('0~7', 'MISMATCH')]:
            results, _ = compare_book(None, self.catalog, self.book(value))
            self.assertEqual(results[0].status, status)

    def test_unresolved_types_and_non_bitfields_do_not_create(self):
        for expression, status in [('unknown.state', 'REVIEW'), ('obj.sign', 'REVIEW'),
                                   ('count', None), ('obj.normal', 'SKIP')]:
            results, _ = compare_book(None, self.catalog, self.book(None, expression))
            self.assertEqual([r.status for r in results], [status] if status else [])

    def test_formula_merge_and_partial_evidence_cannot_create(self):
        for condition in ('formula', 'merge', 'error', 'partial', 'unknown-sheet'):
            book = self.book(None)
            target = book.sheets[0]
            # Include an explicit cell to simulate metadata on an otherwise empty Range.
            from rangevalidator.models import Cell
            cell = Cell(2, 3, 'C2', None)
            if condition == 'formula':
                cell = replace(cell, is_formula=True, formula='=""')
            elif condition == 'merge':
                cell = replace(cell, merged_range='B2:C2', merge_anchor='B2')
            elif condition == 'error':
                cell = replace(cell, kind='error')
            else:
                book.partial = True
                book.read_errors = [{'sheet_name': '1.Main' if condition == 'partial' else ''}]
            target.cells['C2'] = cell
            results, _ = compare_book(None, self.catalog, book)
            self.assertEqual(results[0].status, 'REVIEW', condition)
            self.assertFalse(results[0].candidate)

    def test_create_mismatch_and_manual_review_save_together(self):
        create = compare_book(None, self.catalog, self.book(None))[0][0]
        mismatch = replace(create, id='mismatch', cell='C3', status='MISMATCH', existing='0~0')
        review = replace(create, id='review', cell='C4', status='REVIEW',
                         expected='', existing='0~0', candidate=False)
        rows = [create, mismatch, review]
        session = Session((), (str(self.path),), results=rows, diagnostics=Mock(),
                          files=[FileStatus(1, str(self.path), status='COMPLETED', read_status='SUCCESS')])
        session._signature = session.signature()
        cells = {r.cell: SimpleNamespace(HasFormula=False, MergeCells=False, Value2=r.existing) for r in rows}
        opens = []
        save = Mock(side_effect=lambda **kw: Path(kw['Filename']).write_bytes(b'saved'))
        @contextmanager
        def excel_book(path, diag):
            opens.append(Path(path))
            yield SimpleNamespace(Worksheets=SimpleNamespace(
                Item=lambda _: SimpleNamespace(Range=cells.__getitem__)), SaveAs=save), Mock()
        @contextmanager
        def lock(_):
            yield

        for filled in (False, True):
            cells['C2'].Value2 = 'new value' if filled else None
            cells['C3'].Value2 = mismatch.existing
            cells['C4'].Value2 = review.existing
            with patch.object(writer, 'excel_book', excel_book), patch.object(writer, 'input_locks', lock), \
                    patch.object(writer, 'create_report'):
                folder, outcomes = writer.save_selected(
                    session, [create.id, mismatch.id], self.root / 'out',
                    manual_requests=[ReviewRequest(review.id, '-2~7', True)])
            self.assertEqual(outcomes[create.id]['status'], 'SKIPPED' if filled else 'SAVED_VERIFIED')
            self.assertEqual([outcomes[r.id]['edit_kind'] for r in rows],
                             ['AUTO_CREATE', 'AUTO_MISMATCH', 'MANUAL_REVIEW'])
            self.assertEqual([outcomes[r.id]['status'] for r in rows[1:]], ['SAVED_VERIFIED'] * 2)
            self.assertEqual(opens[-1], folder / 'design_fixed.xlsx')
            self.assertEqual(cells['C2'].Value2, 'new value' if filled else '0~15')
            self.assertEqual(self.path.read_bytes(), b'original')
        self.assertEqual(save.call_count, 2)
        self.assertEqual(session.results, rows)
