"""Photo-derived pointer declarations and safeguards against wrong-scope matches."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook as ExcelWorkbook
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet


FUNCTION = 'pressure_diagnosis_process_can_data'
TITLE = ('1.' + FUNCTION)[:31]
TYPES = '''typedef unsigned char u1;
typedef unsigned int uint;
typedef struct {
    uint BpsRollingCounter : 8;
    uint BpsCheckSum : 8;
} BpsDataStateCycle_t;
'''
SIGNATURE = '(const u1* can_recv_buf, const BpsDataStateCycle_t* bps_data)'


class SheetFunctionLinkageTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def source(self, text='', name='pressure_sensor.c'):
        path = self.root / name
        path.write_text(TYPES + text, encoding='utf-8')
        return path

    def compare(self, paths, title=TITLE, book='unrelated_design', expressions=None):
        expressions = expressions or ['bps_data->BpsCheckSum', 'bps_data->BpsRollingCounter']
        workbook = Workbook(str(self.root / (book + '.xlsx')), 'openpyxl', [
            sheet(title, [('Name', 'Type', 'Range'),
                          *((e, '', '0~4294967295') for e in expressions)])])
        return compare_book(None, catalog_headers(paths), workbook)[0]

    def assert_bitfields(self, rows):
        self.assertTrue(rows)
        self.assertEqual([(r.status, r.expected) for r in rows], [('MISMATCH', '0~255')] * len(rows))

    def test_photo_signature_with_real_excel_and_shadowed_global(self):
        source = self.source('BpsDataStateCycle_t bps_data;\nvoid ' + FUNCTION + SIGNATURE + ''' {
#if defined(BPS_HONEYWELL)
    u1 calculated_checksum = 0;
    if (calculated_checksum != bps_data->BpsCheckSum) { }
#else
    u1 checksum_error = bps_data->BpsCheckSum;
#endif
    u1 current_counter = bps_data->BpsRollingCounter;
}''')
        path = self.root / 'unrelated_design.xlsx'
        w = ExcelWorkbook(); ws = w.active; ws.title = TITLE
        for col, value in ((4, 'Name'), (6, 'Type'), (8, 'Range')):
            ws.cell(28, col, value)
        for row, member in ((29, 'BpsCheckSum'), (30, 'BpsRollingCounter')):
            ws.cell(row, 4, 'bps_data->' + member)
            ws.cell(row, 8, '0~4294967295')
        w.save(path); w.close()
        before = path.read_bytes()
        result = run_comparison(None, [source], [path], backend='openpyxl')
        self.assertEqual(result.errors, [])
        self.assert_bitfields(result.results)
        self.assertEqual([r.cell for r in result.results], ['H29', 'H30'])
        self.assertTrue(all(FUNCTION in r.reason for r in result.results))
        self.assertEqual(path.read_bytes(), before)

    def test_unrelated_file_hint_does_not_hide_unique_function(self):
        source = self.source('void ' + FUNCTION + SIGNATURE + ' {}')
        unrelated = self.source('BpsDataStateCycle_t bps_data;', 'other.c')
        self.assert_bitfields(self.compare([source, unrelated], book='other'))

    def test_case_and_whitespace_keep_local_pointer_precedence(self):
        source = self.source('BpsDataStateCycle_t BPS_DATA;\nvoid ' + FUNCTION + SIGNATURE + ' {}')
        rows = self.compare([source], title=TITLE.upper(), expressions=[' BPS_DATA  -> BPSCHECKSUM '])
        self.assert_bitfields(rows)

    def test_same_prefix_functions_cannot_be_chosen_by_variable_name(self):
        source = self.source('BpsDataStateCycle_t bps_data;\nvoid ' + FUNCTION + SIGNATURE +
                             ' {}\nvoid pressure_diagnosis_process_cancel(void) {}')
        rows = self.compare([source])
        self.assertTrue(all(r.status == 'REVIEW' and not r.candidate for r in rows))
        self.assertTrue(all('함수 후보가 여러 개' in r.reason for r in rows))

    def test_file_hint_can_disambiguate_distinct_functions(self):
        first = self.source('void ' + FUNCTION + SIGNATURE + ' {}')
        second = self.source('void pressure_diagnosis_process_cancel(void) {}', 'other.c')
        self.assert_bitfields(self.compare([first, second], book='pressure_sensor'))
        self.assertTrue(all(r.status == 'REVIEW' for r in self.compare([first, second])))

    def test_same_static_function_in_different_sources_stays_ambiguous(self):
        first = self.source('static void ' + FUNCTION + SIGNATURE + ' {}')
        second = self.source('static void ' + FUNCTION + SIGNATURE + ' {}', 'other.c')
        self.assertTrue(all(r.status == 'REVIEW' for r in self.compare([first, second])))
        self.assert_bitfields(self.compare([first, second], book='pressure_sensor'))

    def test_missing_function_cannot_silently_use_same_named_global(self):
        source = self.source('BpsDataStateCycle_t bps_data;\nvoid Run' + SIGNATURE + ' {}')
        for expr in ('bps_data.BpsCheckSum', 'bps_data->BpsCheckSum'):
            rows = self.compare([source], title='1.Unrecognized', expressions=[expr])
            self.assertEqual(rows[0].status, 'REVIEW')
            self.assertIn('전역 변수로 대신 연결하지 않음', rows[0].reason)

    def test_global_only_rows_still_compare_without_function(self):
        source = self.source('BpsDataStateCycle_t bps_data;')
        self.assert_bitfields(self.compare([source], title='1.Unrecognized', expressions=['bps_data.BpsCheckSum']))

    def test_short_prefix_does_not_guess_a_function(self):
        source = self.source('BpsDataStateCycle_t bps_data;\nvoid ' + FUNCTION + SIGNATURE + ' {}')
        rows = self.compare([source], title='1.pressure', expressions=['bps_data.BpsCheckSum'])
        self.assertEqual(rows[0].status, 'REVIEW')

    def test_recovered_function_applies_to_every_row_not_just_parameters(self):
        source = self.source('BpsDataStateCycle_t global_data;\nvoid ' + FUNCTION +
            '(void) { BpsDataStateCycle_t local_data; }')
        self.assert_bitfields(self.compare([source], expressions=[
            'local_data.BpsCheckSum', 'global_data.BpsRollingCounter']))

    def test_wrong_pointer_operator_still_reports_actual_declaration(self):
        source = self.source('void ' + FUNCTION + SIGNATURE + ' {}')
        rows = self.compare([source], expressions=['bps_data.BpsCheckSum'])
        self.assertEqual(rows[0].status, 'REVIEW')
        self.assertIn('C 선언', rows[0].reason)

    def test_xml_selected_function_keeps_pointer_and_macro_context(self):
        from rangevalidator.build_options import load_build_options
        from test_build_options import xml, macro
        source = self.source('BpsDataStateCycle_t bps_data;\n#ifdef SELECT_BPS\n'
            '#define LIMIT 255U\nvoid ' + FUNCTION + SIGNATURE + ' {}\n#endif\n')
        path = self.root / 'build.xml'
        path.write_text(xml(macro('SELECT_BPS', '1')), encoding='utf-8')
        catalog = catalog_headers([source], build_options=load_build_options(path))
        book = Workbook(str(self.root / 'unrelated.xlsx'), 'openpyxl', [sheet(TITLE,
            [('Name', 'Type', 'Range'), ('bps_data->BpsCheckSum', '', '0~LIMIT')])])
        rows, _ = compare_book(None, catalog, book)
        self.assertEqual([(r.status, r.expected) for r in rows], [('MATCH', '0~255')])
        self.assertIn('C 문맥 pressure_sensor.c', rows[0].reason)


if __name__ == '__main__':
    unittest.main()
