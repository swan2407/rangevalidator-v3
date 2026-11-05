from contextlib import contextmanager
from dataclasses import replace
import importlib
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rangevalidator.models import ReadFailure, Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet


class WriterIsolationTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name).resolve()
        self.source = self.root / 'types.h'
        self.source.write_text('typedef struct { unsigned int active:1; } T; T obj;')
        self.docs = [self.root / name for name in ('one.xlsx', 'two.xlsx', 'unused.xlsx')]
        for doc in self.docs:
            doc.write_bytes(b'original')

    def writer(self):
        with patch.dict(sys.modules, {
            'rangevalidator.preservation': SimpleNamespace(manifest=Mock(), verify=Mock()),
            'rangevalidator.report': SimpleNamespace(create_report=Mock()),
        }):
            return importlib.import_module('rangevalidator.writer')

    def session(self, errors=None):
        def read(path, **_):
            rows = [('Name', 'Type', 'Range'), ('obj.active', '', '0~0')]
            return Workbook(path, 'com', [sheet('1.Main', rows), sheet('2.Other', rows)],
                            partial=errors is not None, read_errors=errors or [])
        with patch('rangevalidator.workflow.read_com', side_effect=read):
            return run_comparison(headers=[self.source], details=self.docs, backend='com')

    def test_partial_read_scopes(self):
        for errors, expected in (
            ([{'sheet_name': 'commonparameter'}], {'1.Main', '2.Other'}),
            ([{'sheet_name': '1.Main'}], {'2.Other'}),
            ([{'sheet_index': 1, 'sheet_name': ''}], set()),
            ([], set()),
        ):
            with self.subTest(errors=errors):
                session = self.session(errors)
                self.assertEqual({r.sheet for r in session.results if r.candidate}, expected)
                self.assertTrue(all(f.partial and f.save_blocked == (not expected) for f in session.files))

    def test_selected_file_locks_and_saveas_failure_are_isolated(self):
        for failing_step in ('lock', 'Open', 'SaveAs', 'verification'):
            with self.subTest(step=failing_step):
                writer = self.writer()
                session = self.session()
                selected = [r.id for r in session.results
                            if r.file in {str(d) for d in self.docs[:2]} and r.sheet == '1.Main']
                lock_calls = []
                @contextmanager
                def lock(paths):
                    lock_calls.append(paths)
                    if failing_step == 'lock' and paths == [str(self.docs[0])]:
                        raise PermissionError('locked')
                    yield
                books = {}
                @contextmanager
                def excel_book(path, diag):
                    path = Path(path)
                    if path in self.docs:
                        if path == self.docs[0] and failing_step == 'Open':
                            raise OSError('Open failed')
                        cell = SimpleNamespace(HasFormula=False, MergeCells=False, Value2='0~0')
                        def save(**kwargs):
                            if path == self.docs[0] and failing_step == 'SaveAs':
                                raise OSError('SaveAs failed')
                            Path(kwargs['Filename']).write_bytes(b'saved')
                        book = SimpleNamespace(Worksheets=SimpleNamespace(
                            Item=lambda _: SimpleNamespace(Range=lambda _: cell)), SaveAs=Mock(side_effect=save))
                        books[path] = book
                    else:
                        value = '0~0' if path.name == 'one_fixed.xlsx' and failing_step == 'verification' else '0~1'
                        cell = SimpleNamespace(HasFormula=False, Value2=value)
                        book = SimpleNamespace(Worksheets=SimpleNamespace(
                            Item=lambda _: SimpleNamespace(Range=lambda _: cell)))
                    yield book, Mock()
                with patch.object(writer, 'input_locks', lock), patch.object(writer, 'excel_book', excel_book),                         patch.object(writer, 'create_report'):
                    _, outcomes = writer.save_selected(session, selected, self.root / 'out')
                self.assertEqual(lock_calls, [[str(doc)] for doc in self.docs[:2]])
                first = 'SAVED_VERIFY_FAILED' if failing_step == 'verification' else 'SAVE_FAILED'
                self.assertEqual([outcomes[i]['status'] for i in selected], [first, 'SAVED_VERIFIED'])
                self.assertEqual(Path(outcomes[selected[1]]['output']).name, 'two_fixed.xlsx')
                books[self.docs[1]].SaveAs.assert_called_once()
                self.assertTrue(all(doc.read_bytes() == b'original' for doc in self.docs))

    def test_preflight_excludes_individual_cells_and_verifies_only_written_cells(self):
        writer = self.writer()
        row = self.session().results[0]
        edits = [replace(row, id=str(i), cell=f'C{i}') for i in range(2, 6)]
        cells = {r.cell: SimpleNamespace(HasFormula=False, MergeCells=False, Value2=r.existing)
                 for r in edits}
        cells['C2'].HasFormula = True
        cells['C3'].MergeCells = True
        cells['C3'].MergeArea = SimpleNamespace(Cells=SimpleNamespace(
            Item=lambda *_: SimpleNamespace(Address='$B$3')))
        cells['C4'].Value2 = 'changed'
        target = self.root / 'saved.xlsx'
        source = self.docs[0]
        reopened = Mock(return_value=SimpleNamespace(HasFormula=False, Value2=row.expected))
        @contextmanager
        def excel_book(path, diag):
            if path == source:
                book = SimpleNamespace(Worksheets=SimpleNamespace(
                    Item=lambda _: SimpleNamespace(Range=cells.__getitem__)),
                    SaveAs=lambda **_: target.write_bytes(b'saved'))
            else:
                book = SimpleNamespace(Worksheets=SimpleNamespace(
                    Item=lambda _: SimpleNamespace(Range=reopened)))
            yield book, Mock()
        outcomes = {}
        with patch.object(writer, 'excel_book', excel_book):
            writer._apply_file(source, target, edits, Mock(), outcomes)
        self.assertEqual([outcomes[r.id]['status'] for r in edits], ['SKIPPED'] * 3 + ['SAVED_VERIFIED'])
        self.assertEqual(cells['C5'].Value2, row.expected)
        self.assertEqual([cells[f'C{i}'].Value2 for i in (2, 3, 4)],
                         [row.existing, row.existing, 'changed'])
        reopened.assert_called_once_with('C5')
        self.assertEqual(source.read_bytes(), b'original')

    def test_missing_sheet_and_invalid_address_are_local_but_com_failure_is_not(self):
        writer = self.writer()
        row = self.session().results[0]
        def item(name):
            if name == 1:
                return SimpleNamespace(Name=row.sheet)
            raise OSError('lookup failed')
        book = SimpleNamespace(Worksheets=SimpleNamespace(Count=1, Item=item))
        self.assertIn('주소', writer._preflight_cell(book, replace(row, cell='XFE1')))
        self.assertIn('시트', writer._preflight_cell(book, replace(row, sheet='3.Missing')))
        with self.assertRaisesRegex(OSError, 'lookup failed'):
            writer._preflight_cell(book, row)

    def test_com_write_failure_never_calls_saveas(self):
        writer = self.writer()
        row = self.session().results[0]
        rows = [row, replace(row, id='second', cell='C3')]
        class ReadOnlyCell:
            HasFormula = MergeCells = False
            @property
            def Value2(self):
                return row.existing
            @Value2.setter
            def Value2(self, value):
                raise OSError('COM write rejected')
        valid = SimpleNamespace(HasFormula=False, MergeCells=False, Value2=row.existing)
        cells = {row.cell: valid, 'C3': ReadOnlyCell()}
        book = SimpleNamespace(Worksheets=SimpleNamespace(
            Item=lambda _: SimpleNamespace(Range=cells.__getitem__)), SaveAs=Mock())
        @contextmanager
        def excel_book(path, diag):
            yield book, Mock()
        outcomes = {}
        with patch.object(writer, 'excel_book', excel_book):
            writer._apply_file(self.docs[0], self.root / 'saved.xlsx', rows, Mock(), outcomes)
        self.assertTrue(all(o['status'] == 'SAVE_FAILED' and o['stage'] == 'writer.edit'
                            for o in outcomes.values()))
        self.assertEqual(valid.Value2, row.expected)
        book.SaveAs.assert_not_called()
        self.assertEqual(self.docs[0].read_bytes(), b'original')

    def test_fixed_path_preserves_suffix_and_existing_fixed_name(self):
        writer = self.writer()
        for source, expected in [('Design.xlsx', 'Design_fixed.xlsx'),
                                 ('BMS_Detailed_Design.xlsx', 'BMS_Detailed_Design_fixed.xlsx'),
                                 ('Design_FIXED.XLSX', 'Design_FIXED.XLSX')]:
            self.assertEqual(writer._fixed_path(Path(source)), Path(expected))

    def test_existing_output_and_partial_saveas_are_save_failures(self):
        writer = self.writer()
        row = self.session().results[0]
        for mode in ('collision', 'partial', 'missing'):
            with self.subTest(mode=mode):
                output = self.root / f'{mode}_fixed.xlsx'
                if mode == 'collision':
                    output.write_bytes(b'existing')
                cell = SimpleNamespace(HasFormula=False, MergeCells=False, Value2=row.existing)
                def save(**_):
                    if mode == 'partial':
                        output.write_bytes(b'partial')
                        raise OSError(5, 'SaveAs interrupted')
                book = SimpleNamespace(Worksheets=SimpleNamespace(
                    Item=lambda _: SimpleNamespace(Range=lambda _: cell)), SaveAs=Mock(side_effect=save))
                @contextmanager
                def excel_book(path, diag):
                    self.assertEqual(path, self.docs[0])
                    yield book, Mock()
                outcomes = {}
                with patch.object(writer, 'excel_book', excel_book):
                    writer._apply_file(self.docs[0], output, [row], Mock(), outcomes)
                result = outcomes[row.id]
                self.assertEqual(result['status'], 'SAVE_FAILED')
                self.assertEqual(result['saveas_returned'], mode == 'missing')
                self.assertEqual(result['stage'], 'writer.saveas' if mode == 'partial' else 'writer.output_check')
                if mode == 'collision':
                    book.SaveAs.assert_not_called()
                    self.assertEqual(output.read_bytes(), b'existing')
                if mode == 'partial':
                    self.assertEqual(result['code'], 5)
                    self.assertEqual(output.read_bytes(), b'partial')

    def test_reopen_failure_preserves_path_and_does_not_retry(self):
        writer = self.writer()
        row = self.session().results[0]
        output = self.root / 'one_fixed.xlsx'
        cell = SimpleNamespace(HasFormula=False, MergeCells=False, Value2=row.existing)
        book = SimpleNamespace(Worksheets=SimpleNamespace(
            Item=lambda _: SimpleNamespace(Range=lambda _: cell)),
            SaveAs=Mock(side_effect=lambda **_: output.write_bytes(b'saved')))
        opens = []
        @contextmanager
        def excel_book(path, diag):
            opens.append(path)
            if path == output:
                raise ReadFailure('com.Workbooks.Open', PermissionError(13, 'temporarily unavailable'))
            yield book, Mock()
        outcomes = {}
        with patch.object(writer, 'excel_book', excel_book):
            writer._apply_file(self.docs[0], output, [row], Mock(), outcomes)
        result = outcomes[row.id]
        self.assertEqual((result['status'], result['stage'], result['output'], result['code']),
                         ('SAVED_UNVERIFIED', 'writer.reopen', str(output), 13))
        self.assertEqual(opens, [self.docs[0], output])
        book.SaveAs.assert_called_once()

    def test_cell_evidence_and_cleanup_warning_are_preserved(self):
        writer = self.writer()
        session = self.session()
        rows = [r for r in session.results if r.file == str(self.docs[0])]
        formula_in_second_sheet = True
        @contextmanager
        def excel_book(path, diag):
            if Path(path) == self.docs[0]:
                cell = SimpleNamespace(HasFormula=False, MergeCells=False, Value2=rows[0].existing)
                # Each worksheet has its own cell, so all rows pass the preflight.
                def sheet_by_name(_):
                    return SimpleNamespace(Range=lambda _: cell)
                yield SimpleNamespace(Worksheets=SimpleNamespace(Item=sheet_by_name),
                                      SaveAs=lambda **kw: Path(kw['Filename']).write_bytes(b'saved')), Mock()
            else:
                def sheet_by_name(name):
                    cell = SimpleNamespace(HasFormula=formula_in_second_sheet and name == '2.Other', Value2=rows[0].expected)
                    return SimpleNamespace(Range=lambda _: cell)
                yield SimpleNamespace(Worksheets=SimpleNamespace(Item=sheet_by_name)), Mock()
                issue = ReadFailure('com.Quit', RuntimeError('Excel exit issue'))
                raise ReadFailure('cleanup', RuntimeError('Excel cleanup failed'), [issue.as_dict()])
        @contextmanager
        def lock(_):
            yield
        with patch.object(writer, 'excel_book', excel_book), patch.object(writer, 'input_locks', lock), \
                patch.object(writer, 'create_report', side_effect=OSError('report failed')):
            _, outcomes = writer.save_selected(session, [r.id for r in rows], self.root / 'out')
        self.assertEqual([outcomes[r.id]['status'] for r in rows], ['SAVED_VERIFIED', 'SAVED_VERIFY_FAILED'])
        self.assertEqual({o['file_status'] for o in outcomes.values()}, {'SAVED_VERIFY_FAILED'})
        self.assertTrue(all(o['cleanup_warnings'][0]['status'] == 'CLEANUP_WARNING' for o in outcomes.values()))
        self.assertIn('수식', outcomes[rows[1].id]['reason'])
        self.assertTrue(outcomes.report_warning)
        self.assertEqual(self.docs[0].read_bytes(), b'original')

        formula_in_second_sheet = False
        with patch.object(writer, 'excel_book', excel_book), patch.object(writer, 'input_locks', lock), \
                patch.object(writer, 'create_report'):
            _, outcomes = writer.save_selected(session, [r.id for r in rows], self.root / 'out')
        self.assertEqual({o['file_status'] for o in outcomes.values()}, {'SAVED_VERIFIED'})
        self.assertTrue(all(o['cleanup_warnings'] for o in outcomes.values()))
