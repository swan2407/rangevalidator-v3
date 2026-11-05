from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Cell, Sheet, Workbook
from rangevalidator.workflow import run_comparison


def sheet(name, values):
    cells = {}
    for row, values_in_row in enumerate(values, 1):
        for column, value in enumerate(values_in_row, 1):
            address = f'{chr(64 + column)}{row}'
            cells[address] = Cell(row, column, address, value)
    return Sheet(name, 'visible', cells)


class FunctionSheetTests(unittest.TestCase):
    def test_selects_names_and_preserves_decisions(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / 'test.h'
            source.write_text('typedef struct { unsigned int bit:3; } T; extern T obj;')
            catalog = catalog_headers([source])
            table = [('Name', 'Type', 'Range'),
                     ('obj.bit', 'unsigned int', '0~7'),
                     ('obj.bit', 'unsigned int', '0~3')]
            names = ['7.Last', '1.Main', '12.Process_Data', '3. 함수이름']
            excluded = ['\uD45C\uC9C0', '\uBAA9\uCC28', 'commonparameter', 'CommonParameter',
                        'Design', 'Main', '1', '1.', '1. ', 'ABC.Main']
            sheets = [sheet(excluded[0], [('unknown layout',)])]
            sheets += [sheet(name, table) for name in excluded[1:] + names]
            results, coverage = compare_book(None, catalog, Workbook('test.xlsx', 'com', sheets))
            self.assertEqual([item[0] for item in coverage], names)
            self.assertEqual([(r.sheet, r.status, r.expected, r.candidate) for r in results],
                             [(name, status, '0~7', status == 'MISMATCH')
                              for name in names for status in ('MATCH', 'MISMATCH')])

    def test_no_function_sheets(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source, detail = root / 'test.h', root / 'test.xlsx'
            source.touch()
            detail.touch()
            book = Workbook(str(detail), 'com', [sheet('commonparameter', [('unknown',)])])
            with patch('rangevalidator.workflow.read_auto', return_value=book):
                session = run_comparison(None, [source], [detail], 'auto')
            self.assertEqual(session.errors, [])
            self.assertEqual(session.results, [])
            self.assertEqual(session.coverage, [])
            self.assertEqual(session.files[0].status, 'NO_VARIABLES')
            self.assertIn('비교 대상 함수 시트', session.files[0].reason)
