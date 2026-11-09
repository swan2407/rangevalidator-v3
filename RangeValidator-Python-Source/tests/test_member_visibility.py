from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook
from rangevalidator.workflow import run_comparison


class MemberVisibilityTests(unittest.TestCase):
    def test_real_workbooks_retain_members_with_empty_or_partial_catalog(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'types.h'
            for fmt in ('A', 'B'):
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = 'update1' if fmt == 'A' else '1_update'
                sheet.append(['Input/Output', 'Property', 'Name', 'Data Type', 'Scope']
                             if fmt == 'A' else ['Name', 'Type', 'Range'])
                expressions = ['obj.state', 'ptr->state', 'obj.status.mode', 'arr[0].state',
                               'ptr -> state', 'obj.normal', 'obj.missing', 'obj.status.fault',
                               'obj.state']
                for i, expr in enumerate(expressions):
                    row = [expr, 'unsigned int', '0~7']
                    # A name-only nested row must not end the table.
                    if i == 7:
                        row = [expr, None, None]
                    sheet.append(([None, None] + row) if fmt == 'A' else row)
                path = root / f'{fmt}.xlsx'
                workbook.save(path)
                workbook.close()
                source.write_text('', encoding='utf-8')
                session = run_comparison(None, [source], [path], 'openpyxl')
                self.assertEqual(session.errors, [])
                self.assertEqual([r.expression for r in session.results], expressions)
                self.assertTrue(all(r.status == 'REVIEW' and not r.candidate for r in session.results))
                source.write_text('typedef struct { unsigned int state:3; unsigned int normal; } T;\n'
                                  'extern T obj; extern T *ptr; extern T arr[2];', encoding='utf-8')
                session = run_comparison(None, [source], [path], 'openpyxl')
                self.assertEqual(session.errors, [])
                self.assertEqual([(r.expression, r.status) for r in session.results], [
                    ('obj.state', 'MATCH'), ('ptr->state', 'MATCH'), ('obj.status.mode', 'REVIEW'),
                    ('arr[0].state', 'MATCH'), ('ptr -> state', 'MATCH'), ('obj.normal', 'SKIP'),
                    ('obj.missing', 'REVIEW'),
                    ('obj.status.fault', 'REVIEW'), ('obj.state', 'MATCH')])
                self.assertTrue(all(not r.candidate for r in session.results))
                self.assertEqual(session.results[-1].cell, 'E10' if fmt == 'A' else 'C10')

    def test_missing_range_column_keeps_variable_review_without_write_coordinate(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'empty.h'
            source.write_text('', encoding='utf-8')
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '1. update'
            sheet.append(['Name', 'Type'])
            sheet.append(['obj.state', 'unsigned int'])
            sheet.append(['ptr->state', 'unsigned int'])
            path = root / 'missing.xlsx'
            workbook.save(path)
            workbook.close()
            session = run_comparison(None, [source], [path], 'openpyxl')
            self.assertEqual([r.expression for r in session.results], ['obj.state', 'ptr->state'])
            for result in session.results:
                self.assertEqual((result.status, result.cell, result.candidate), ('REVIEW', '', False))
                self.assertIn('RANGE_CELL_UNRESOLVED', result.reason)
