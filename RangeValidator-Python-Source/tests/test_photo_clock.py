"""Photographed source, real workbook reading, and independent bit-width expectations."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.ch_diagnostics import CHDiagnostics
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet
from tools.photo_clock_audit import PHOTO_STRUCT, PHOTO_ROWS, SCENARIOS, make_case


class PhotoClockTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def compare(self, text, expression='can_hu_clock_01.hu_displayday'):
        header = self.root / 'clock.h'
        header.write_text(text, encoding='utf-8')
        book = Workbook(str(self.root / 'clock.xlsx'), 'openpyxl', [
            sheet('1.Update', [('Name', 'Type', 'Range'), (expression, '', '0~4294967295')])])
        return compare_book(None, catalog_headers([header]), book)[0][0]

    def test_photo_lines_real_excel_rows_and_separate_condition_scenarios(self):
        for scenario in SCENARIOS:
            with self.subTest(scenario=scenario):
                paths, workbook = make_case(self.root / scenario, scenario)
                original = workbook.read_bytes()
                session = run_comparison(None, paths, [workbook], backend='openpyxl')
                self.assertEqual(session.errors, [])
                self.assertEqual(len(session.results), 4)
                source_lines = paths[0].read_text(encoding='utf-8').splitlines()
                diagnostic = CHDiagnostics(paths)
                for result, (member, width, expected, row, line) in zip(session.results, PHOTO_ROWS):
                    self.assertIn(f'{member}:{width}', source_lines[line - 1])
                    self.assertEqual((result.expression, result.cell), (f'can_hu_clock_01.{member}', f'H{row}'))
                    if scenario == 'conflicting_structure' and member == 'hu_displayday':
                        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
                        self.assertIn('0~255', result.reason)
                        self.assertIn('0~15', result.reason)
                        continue
                    self.assertEqual((result.status, result.expected, result.candidate), ('MISMATCH', expected, True))
                    trace = diagnostic.trace_expression(result.expression, paths[1], 'Update')
                    self.assertEqual((trace['code'], trace['range']), ('TYPE_MATCH', [0, (1 << width) - 1]))
                    self.assertNotIn('CONDITIONAL', result.reason)
                    if scenario == 'photo_only':
                        self.assertIn('uint 정의 미수집', result.reason)
                self.assertEqual(workbook.read_bytes(), original)

    def test_explicit_signed_uint_overrides_unsigned_convention(self):
        result = self.compare('typedef signed int uint;\n' + PHOTO_STRUCT)
        self.assertEqual((result.status, result.expected), ('MISMATCH', '-128~127'))
        self.assertNotIn('관용 타입', result.reason)

    def test_type_aliases_with_same_target_compare_across_directives(self):
        text = '#ifdef FEATURE\ntypedef unsigned int ClockWord;\n#endif\n'
        text += 'typedef ClockWord uint;\ntypedef unsigned int uint;\n' + PHOTO_STRUCT
        result = self.compare(text)
        self.assertEqual((result.status, result.expected), ('MISMATCH', '0~255'))

    def test_uint_definition_in_separate_registered_header_without_include_wins(self):
        paths, workbook = make_case(self.root, 'photo_only')
        types = self.root / 'platform.h'
        types.write_text('typedef signed int uint;', encoding='utf-8')
        session = run_comparison(None, [*paths, types], [workbook], backend='openpyxl')
        self.assertEqual([r.expected for r in session.results], ['-128~127', '-8~7', '-2~1', '-128~127'])
        self.assertTrue(all('관용 타입' not in r.reason for r in session.results))

    def test_conflicting_or_unresolved_uint_is_never_replaced_by_convention(self):
        for preamble in ('#if FEATURE\ntypedef signed int uint;\n#else\ntypedef unsigned int uint;\n#endif\n',
                         'typedef Unknown uint;\n',
                         '#if FEATURE\ntypedef unsigned int *uint;\n#endif\n'):
            with self.subTest(preamble=preamble):
                result = self.compare(preamble + PHOTO_STRUCT)
                self.assertEqual((result.status, result.expected, result.candidate), ('REVIEW', '', False))
                self.assertNotIn('관용 타입', result.reason)

    def test_conditional_width_macro_requires_one_known_body(self):
        body = PHOTO_STRUCT.replace('hu_displayday:8', 'hu_displayday:DAY_BITS')
        for second, expected in [('', '0~255'), ('#else\n#define DAY_BITS 8\n', '0~255'),
                                 ('#else\n#define DAY_BITS 4\n', '')]:
            with self.subTest(second=second):
                result = self.compare('#if FEATURE\n#define DAY_BITS 8\n' + second + '#endif\n' + body)
                self.assertEqual((result.status, result.expected), ('MISMATCH' if expected else 'REVIEW', expected))

    def test_member_branch_conflict_only_blocks_conflicting_member(self):
        body = PHOTO_STRUCT.replace('uint hu_displayday:8;',
            '#if FEATURE\n    uint hu_displayday:8;\n#else\n    uint hu_displayday:4;\n#endif\n')
        result = self.compare(body)
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
        self.assertIn('0~255', result.reason)
        self.assertIn('0~15', result.reason)
        same = self.compare(body.replace('hu_displayday:4', 'hu_displayday:8'))
        self.assertEqual((same.status, same.expected), ('MISMATCH', '0~255'))
        month = self.compare(body, 'can_hu_clock_01.hu_displaymonth')
        self.assertEqual((month.status, month.expected), ('MISMATCH', '0~15'))

    def test_conflicting_conditional_prototype_is_not_silently_accepted(self):
        text = PHOTO_STRUCT + '\ntypedef struct { unsigned int hu_displayday:3; } Other;\n'
        text += '#if FEATURE\nvoid Update(Other *ptr);\n#endif\nvoid Update(can_hu_clock_01_t *ptr) {}\n'
        result = self.compare(text, 'ptr->hu_displayday')
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
        self.assertIn('AMBIGUOUS', result.reason)

    def test_conditional_pointer_alias_retains_pointer_access(self):
        text = PHOTO_STRUCT + '\n#if FEATURE\ntypedef can_hu_clock_01_t *ClockPtr;\n#endif\nClockPtr ptr;'
        good = self.compare(text, 'ptr->hu_displayday')
        bad = self.compare(text, 'ptr.hu_displayday')
        self.assertEqual((good.status, good.expected), ('MISMATCH', '0~255'))
        self.assertEqual((bad.status, bad.candidate), ('REVIEW', False))

    def test_alternate_function_definitions_compare_only_the_row_in_question(self):
        paths, workbook = make_case(self.root, 'known_uint_no_variable')
        paths[0].write_text(paths[0].read_text(encoding='utf-8') + '\ntypedef struct {\n'
            'unsigned short hu_displayday:4; unsigned short hu_displaymonth:4;\n'
            'unsigned int hu_displayyear2:8;\n} Other;\n', encoding='utf-8')
        paths[1].write_text('#include "canfd_comm.h"\n#ifdef FEATURE\n'
            'void Update(void) { can_hu_clock_01_t can_hu_clock_01; }\n#else\n'
            'void Update(void) { Other can_hu_clock_01; }\n#endif\n', encoding='utf-8')
        session = run_comparison(None, paths, [workbook], backend='openpyxl')
        self.assertEqual(session.errors, [])
        self.assertEqual([(r.cell, r.status, r.expected) for r in session.results], [
            ('H31', 'REVIEW', ''), ('H32', 'MISMATCH', '0~15'),
            ('H33', 'REVIEW', ''), ('H34', 'MISMATCH', '0~255')])
        self.assertIn('0~255', session.results[0].reason)
        self.assertIn('0~15', session.results[0].reason)
        self.assertIn('MEMBER_NOT_FOUND', session.results[2].reason)
