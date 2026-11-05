import unittest
from dataclasses import replace
from rangevalidator.analysis import analyze
from rangevalidator.design import discover, compare
from rangevalidator.models import Cell, Sheet, Workbook


def book(fmt='A'):
    names = ('Name', 'Type', 'Range') if fmt == 'A' else ('Name', 'Data type', 'Scope')
    cells = {}
    for col, v in enumerate(names, 1):
        a = f'{chr(64 + col)}8'
        cells[a] = Cell(8, col, a, v)
    for col, v in enumerate(('p->active', 'unsigned int', '0~3'), 1):
        a = f'{chr(64 + col)}9'
        cells[a] = Cell(9, col, a, v)
    return Workbook('synthetic.xlsx', 'test', [Sheet('Design', 'visible', cells)])


class DesignTests(unittest.TestCase):
    def setUp(self):
        self.model = analyze('typedef struct { unsigned int active:2; int normal; } A; A *p;')

    def test_headers_both_formats_coordinates(self):
        for fmt in ('A', 'B'):
            row, = discover(book(fmt))
            self.assertEqual((row.format, row.sheet, row.row, row.range_cell.address), (fmt, 'Design', 9, 'C9'))
            self.assertEqual(compare(self.model, row).status, 'MATCH')

    def test_numeric_policy(self):
        row, = discover(book())
        for text, status in [('0..3', 'MATCH'), ('0.0~3e0', 'MATCH'), ('0~255', 'AUTO_FIX_CANDIDATE'), ('3~0', 'AUTO_FIX_CANDIDATE'), ('', 'AUTO_FIX_CANDIDATE'), (None, 'AUTO_FIX_CANDIDATE'), ('3', 'AUTO_FIX_CANDIDATE'), ('3e0', 'AUTO_FIX_CANDIDATE'), ('unknown', 'REVIEW')]:
            with self.subTest(text=text):
                result = compare(self.model, replace(row, range_cell=replace(row.range_cell, value=text)))
                self.assertEqual(result.status, status)
                self.assertEqual(result.expected, (0, 3))
                self.assertEqual(result.evidence.field.file, '<memory>')

    def test_unsafe_data_headers(self):
        for key in ('A8', 'A9', 'B9', 'C9'):
            for flags in (dict(is_formula=True), dict(merged_range='A9:B9'), dict(kind='error')):
                wb = book()
                ws = wb.sheets[0]
                ws.cells[key] = replace(ws.cells[key], **flags)
                row, = discover(wb)
                self.assertEqual(compare(self.model, row).status, 'REVIEW')

    def test_type_contradiction(self):
        row, = discover(book())
        row = replace(row, type_cell=replace(row.type_cell, value='signed int'))
        self.assertEqual(compare(self.model, row).status, 'REVIEW')

    def test_duplicate_headers(self):
        wb = book()
        wb.sheets[0].cells['D8'] = Cell(8, 4, 'D8', 'Name')
        rows = discover(wb)
        self.assertEqual(len(rows), 1)
        self.assertEqual(compare(self.model, rows[0]).status, 'REVIEW')

    def test_ordinary_skip(self):
        row, = discover(book())
        row = replace(row, name_cell=replace(row.name_cell, value='p->normal'))
        self.assertEqual(compare(self.model, row).status, 'SKIP')

    def test_partial_table_terminates(self):
        wb = book()
        ws = wb.sheets[0]
        ws.cells.update(A10=Cell(10,1,'A10','Name'), B10=Cell(10,2,'B10','Type'), C10=Cell(10,3,'C10','Value'), A11=Cell(11,1,'A11','calibration'))
        self.assertEqual(len(discover(wb)), 1)
