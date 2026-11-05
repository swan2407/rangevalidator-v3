from contextlib import contextmanager
import importlib
from pathlib import Path
import os
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from dataclasses import replace
from unittest.mock import Mock, patch

from rangevalidator.models import ReadFailure, Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet


class FileStateRemovalTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name).resolve()
        self.source = self.root / 'device.c'
        self.source.write_text('typedef struct { unsigned int active:1; } T; T obj;')
        self.detail = self.root / 'device.xlsx'
        self.detail.write_bytes(b'original')

    def compare(self, reader=None, legacy_common=None):
        table = [('Name', 'Type', 'Range'), ('obj.active', '', '0~0'),
                 ('unknown.active', '', '0~0'), ('count', '', '0~0')]
        book = Workbook(str(self.detail), 'com', [
            sheet('1.Main', table), sheet('commonparameter', table)])
        with patch('rangevalidator.workflow.read_com', side_effect=reader or (lambda *a, **k: book)) as read:
            session = run_comparison(legacy_common, [self.source], [self.detail], 'com')
        read.assert_called_once()
        self.assertEqual(read.call_args.args[0], str(self.detail))
        return session

    def writer(self):
        # These tests exercise the Writer with a fake COM book and no XLSX report.
        with patch.dict(sys.modules, {
            'rangevalidator.preservation': SimpleNamespace(manifest=Mock(), verify=Mock()),
            'rangevalidator.report': SimpleNamespace(create_report=Mock()),
        }):
            return importlib.import_module('rangevalidator.writer')

    def test_file_metadata_and_content_changes_do_not_block_save(self):
        session = self.compare(legacy_common=self.root / 'missing-common.xlsx')
        self.assertEqual([(r.expression, r.status, r.candidate) for r in session.results],
                         [('obj.active', 'MISMATCH', True), ('unknown.active', 'REVIEW', False)])
        candidate = session.results[0]
        writer = self.writer()
        @contextmanager
        def locks(paths):
            self.assertEqual(paths, [str(self.detail)])
            yield
        def applied(source, target, edits, diag, outcomes):
            outcomes[candidate.id] = {'status': 'SAVED_VERIFIED'}

        for data in (None, b'different size', b'DIFFERENT SIZE'):
            stamp = self.detail.stat()
            if data is None:
                os.utime(self.detail, ns=(stamp.st_atime_ns, stamp.st_mtime_ns + 1_000_000_000))
            else:
                self.detail.write_bytes(data)
                os.utime(self.detail, ns=(stamp.st_atime_ns, stamp.st_mtime_ns))
            with patch.object(writer, 'input_locks', locks), patch.object(
                    writer, '_apply_file', side_effect=applied) as apply, patch.object(writer, 'create_report'):
                _, outcomes = writer.save_selected(session, [candidate.id], self.root / 'out')
            apply.assert_called_once()
            self.assertEqual(outcomes[candidate.id]['status'], 'SAVED_VERIFIED')

    def test_missing_inputs_still_fail(self):
        self.source.unlink()
        with self.assertRaisesRegex(ReadFailure, 'source.read'):
            run_comparison(headers=[self.source], details=[self.detail], backend='com')
        self.source.touch()
        self.detail.unlink()
        def missing(path, **_):
            Path(path).read_bytes()
        session = self.compare(missing)
        self.assertEqual(session.files[0].error_kind, 'READ_ERROR')
        self.assertEqual(session.results, [])

    def test_session_selection_and_source_integrity(self):
        session = self.compare()
        candidate, review = session.results
        writer = self.writer()
        for ids in ([review.id], [candidate.id, candidate.id], ['foreign']):
            with self.assertRaises(ValueError):
                writer.save_selected(session, ids, self.root / 'out')
        session.id = 'tampered'
        with self.assertRaises(RuntimeError):
            writer.save_selected(session, [candidate.id], self.root / 'out')
        session = self.compare()
        session.results[0] = replace(candidate, file=str(self.root / 'other.xlsx'))
        session._signature = session.signature()
        with self.assertRaises(ValueError):
            writer.save_selected(session, [candidate.id], self.root / 'out')

    def test_writer_checks_original_cell_and_refuses_overwrite(self):
        row = self.compare().results[0]
        writer = self.writer()
        cell = SimpleNamespace(HasFormula=False, MergeCells=False, Value2='changed')
        worksheet = SimpleNamespace(Range=Mock(return_value=cell))
        book = SimpleNamespace(Worksheets=SimpleNamespace(Item=Mock(return_value=worksheet)), SaveAs=Mock())
        @contextmanager
        def excel_book(path, diag):
            self.assertEqual(Path(path), self.detail)
            yield book, Mock()
        with patch.object(writer, 'excel_book', excel_book):
            outcomes = {}
            writer._apply_file(self.detail, self.root / 'new.xlsx', [row], Mock(), outcomes)
            self.assertEqual(outcomes[row.id]['status'], 'SKIPPED')
            self.assertIn('값이 비교 때와 다름', outcomes[row.id]['reason'])
            cell.Value2 = row.existing
            outcomes = {}
            writer._apply_file(self.detail, self.detail, [row], Mock(), outcomes)
            self.assertEqual(outcomes[row.id]['status'], 'SAVE_FAILED')
        book.SaveAs.assert_not_called()
        self.assertEqual(self.detail.read_bytes(), b'original')
