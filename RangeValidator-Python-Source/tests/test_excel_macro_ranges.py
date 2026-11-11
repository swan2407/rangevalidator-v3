"""Excel symbols must use C/H definitions and the selected build context."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook as ExcelWorkbook
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from rangevalidator.build_options import load_build_options
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet
from test_build_options import xml, macro


class ExcelMacroRangeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def write(self, name, text):
        path = self.root / name
        path.write_text(text, encoding='utf-8')
        return path

    def compare(self, source, value, *, definitions=None, options=None):
        paths = [self.write('relay.c', source)]
        if definitions is not None:
            paths.append(self.write('com_def.h', definitions))
        catalog = catalog_headers(paths, build_options=options)
        book = Workbook(str(self.root / 'relay_Detail.xlsx'), 'openpyxl', [
            sheet('1.Main', [('Name', 'Type', 'Range'), ('relay_diag.welded', '', value)])])
        rows, _ = compare_book(None, catalog, book)
        self.assertEqual(len(rows), 1)
        return rows[0]

    def test_photo_no_definition_and_three_excel_cells(self):
        # The photograph shows com_def.h:56 and these three result rows.
        # The field declarations reconstruct the widths visible in that result table.
        header = self.write('com_def.h', '\n' * 55 + '#define NO 0U\n')
        source = self.write('relay.c', '#include "com_def.h"\n'
            'typedef struct { unsigned int welded:2; unsigned int ioffset_fail:2; '
            'unsigned int cell_sensing_fail:1; } relay_t; relay_t relay_diag; void Main(void) {}')
        workbook = self.root / 'relay_Detail.xlsx'
        w = ExcelWorkbook()
        ws = w.active
        ws.title = '1.Main'
        for col, value in enumerate(('Input/Output', 'Property', 'Name', 'Data Type', 'Scope'), 1):
            ws.cell(37, col, value)
        fields = [(38, 'welded'), (39, 'ioffset_fail'), (42, 'cell_sensing_fail')]
        for row, member in fields:
            for col, value in enumerate(('Input', 'Global', 'relay_diag.' + member, 'unsigned int', 'NO'), 1):
                ws.cell(row, col, value)
        w.save(workbook)
        w.close()
        before = workbook.read_bytes()
        session = run_comparison(None, [source, header], [workbook], backend='openpyxl')
        self.assertEqual(session.errors, [])
        self.assertEqual([(r.cell, r.existing, r.status, r.expected) for r in session.results], [
            ('E38', 'NO', 'MISMATCH', '0~3'), ('E39', 'NO', 'MISMATCH', '0~3'),
            ('E42', 'NO', 'MISMATCH', '0~1')])
        self.assertTrue(all("'NO' → 0" in r.reason for r in session.results))
        self.assertEqual(before, workbook.read_bytes())

    def test_no_is_looked_up_instead_of_assumed_zero(self):
        for number in (0, 2, 9):
            with self.subTest(number=number):
                source = f'#define NO {number}U\nstruct T {{ unsigned int welded:2; }} relay_diag;'
                row = self.compare(source, 'NO')
                self.assertEqual((row.status, row.existing), ('MISMATCH', 'NO'))
                self.assertIn(f"'NO' → {number}", row.reason)

    def test_symbolic_interval_aliases_and_unsigned_suffixes(self):
        source = ('#define NO 0U\n#define END 3U\n#define LIMIT END\n'
                  'struct T { unsigned int welded:2; } relay_diag;')
        for value in ('NO~LIMIT', '[NO, LIMIT]', 'NO to LIMIT', '0U~3U'):
            with self.subTest(value=value):
                row = self.compare(source, value)
                self.assertEqual((row.status, row.expected, row.existing), ('MATCH', '0~3', value))
                self.assertIn('→ 0~3', row.reason)

    def test_registered_macro_without_include_can_supply_excel_bound(self):
        row = self.compare('struct T { unsigned int welded:2; } relay_diag;', 'NO~END',
                           definitions='#define NO 0U\n#define END 3U')
        self.assertEqual(row.status, 'MATCH')

    def test_missing_commented_or_conflicting_definitions_stay_review(self):
        for source, definitions in [
            ('', ''), ('', '//#define NO 0U\n'),
            ('#include "com_def.h"\n#define NO 0U\n', '#define NO 2U\n'),
        ]:
            with self.subTest(source=source, definitions=definitions):
                row = self.compare(source + 'struct T { unsigned int welded:2; } relay_diag;',
                                   'NO', definitions=definitions)
                self.assertEqual((row.status, row.candidate, row.existing), ('REVIEW', False, 'NO'))

    def test_xml_selects_the_active_macro_in_each_c_context(self):
        header = self.write('com_def.h', '#if MODEL == 1\n#define FIRST 0U\n#define LAST 3U\n'
                            '#else\n#define FIRST 1U\n#define LAST 3U\n#endif\n')
        first = self.write('first.c', '#define MODEL 1\n#include "com_def.h"\n'
                           'struct T { unsigned int welded:2; } relay_diag; void Main(void) {}')
        second = self.write('second.c', '#define MODEL 2\n#include "com_def.h"\n'
                            'struct T { unsigned int welded:2; } relay_diag; void Main(void) {}')
        options = load_build_options(self.write('build.xml', xml()))
        catalog = catalog_headers([header, first, second], build_options=options)
        for name, status, expanded in [('first', 'MATCH', '0~3'), ('second', 'MISMATCH', '1~3')]:
            book = Workbook(str(self.root / (name + '_Detail.xlsx')), 'openpyxl', [
                sheet('1.Main', [('Name', 'Type', 'Range'), ('relay_diag.welded', '', 'FIRST~LAST')])])
            row, = compare_book(None, catalog, book)[0]
            self.assertEqual((row.status, row.existing), (status, 'FIRST~LAST'))
            self.assertIn('→ ' + expanded, row.reason)
            self.assertIn('build.xml', row.reason)

    def test_xml_macro_value_is_used_and_source_undef_can_replace_it(self):
        options = load_build_options(self.write('build.xml', xml(macro('NO', '2U'))))
        source = 'struct T { unsigned int welded:2; } relay_diag;'
        row = self.compare(source, 'NO', options=options)
        self.assertIn("'NO' → 2", row.reason)
        row = self.compare('#undef NO\n#define NO 3U\n' + source, 'NO', options=options)
        self.assertIn("'NO' → 3", row.reason)

    def test_large_and_negative_bounds_use_constant_domain_not_width_limit(self):
        row = self.compare('#define MAX 4294967295U\n'
            'struct T { unsigned int welded:32; } relay_diag;', '0U~MAX')
        self.assertEqual((row.status, row.expected), ('MATCH', '0~4294967295'))
        row = self.compare('#define MIN (-128)\n#define MAX 127\n'
            'struct T { signed int welded:8; } relay_diag;', 'MIN~MAX')
        self.assertEqual((row.status, row.expected), ('MATCH', '-128~127'))
        row = self.compare('#define MIN (0U - 1U)\n'
            'struct T { signed int welded:8; } relay_diag;', 'MIN~127')
        self.assertEqual((row.status, row.candidate), ('REVIEW', False))
