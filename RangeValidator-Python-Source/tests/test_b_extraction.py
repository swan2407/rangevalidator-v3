from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.extraction_diagnostics import diagnose_session
from rangevalidator.workflow import run_comparison
from tools.excel4_synthetic import build_b_safety, build_b_workbook


FIELDS = ('sheet', 'section', 'section_cell', 'name', 'name_cell', 'range_cell', 'comparison_status')


def observed(report):
    return [{key: var[key] for key in FIELDS} for sheet in report['sheets']
            for var in sheet['variables'] if var['extracted']]


class BExtractionTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root / 'types.h'
        fields = ' '.join(f'unsigned int v{i:03d}:3;' for i in range(1, 201))
        self.source.write_text(f'typedef struct {{ {fields} }} B_t; extern B_t obj;\n', encoding='utf-8')

    def compare(self, path):
        session = run_comparison(None, [self.source], [path], backend='openpyxl')
        return session, diagnose_session(session)[0]

    def test_sectioned_workbooks_with_one_ten_and_one_hundred_sheets(self):
        for count in (1, 10, 100):
            with self.subTest(count=count):
                path = self.root / f'b_{count}.xlsx'
                truth = build_b_workbook(path, count)
                _, report = self.compare(path)
                self.assertEqual(observed(report), truth)
                self.assertEqual(report['counts']['function_candidates'], count)
                for sheet in (s for s in report['sheets'] if s['function_sheet']):
                    empty, *tables = sheet['tables']
                    self.assertEqual(empty['section'], 'Parameter(Input)')
                    self.assertIsNone(empty['start_row'])
                    self.assertEqual(empty['end_reason'], 'NEXT_SECTION')
                    self.assertEqual(len(tables), 4)
                    self.assertEqual(tables[1]['end_reason'], 'NEXT_HEADER')
                    self.assertEqual(tables[1]['section_cell'], tables[2]['section_cell'])
                    self.assertGreaterEqual(tables[0]['end_row'], tables[0]['last_variable_row'])

    def test_incomplete_sections_and_uncertain_boundary_do_not_reuse_range(self):
        path = self.root / 'safety.xlsx'
        truth = build_b_safety(path)
        session, report = self.compare(path)
        self.assertEqual(observed(report), truth)
        variables = {var['name']: var for var in report['sheets'][0]['variables']}
        self.assertEqual(variables['obj.v005']['reason'], 'RANGE_CELL_UNRESOLVED')
        self.assertEqual(variables['obj.v006']['reason'], 'DUPLICATE_HEADER')
        self.assertEqual(variables['obj.v009']['reason'], 'RANGE_CELL_UNRESOLVED')
        for name in ('obj.v005', 'obj.v006', 'obj.v009'):
            self.assertIsNone(variables[name]['range_cell'])
            self.assertFalse(variables[name]['extracted'])
        self.assertNotIn('obj.v008', variables)
        self.assertEqual([r.cell for r in session.results if r.candidate], ['E6', 'E8', 'H23'])
        self.assertEqual(report['sheets'][0]['tables'][3]['end_reason'], 'UNCERTAIN_BOUNDARY')
