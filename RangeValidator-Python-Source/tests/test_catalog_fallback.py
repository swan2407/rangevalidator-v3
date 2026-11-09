"""Document comparisons using registered struct members without include/declaration requirements."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


class CatalogFallbackTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def write(self, name, content):
        path = self.root / name
        path.write_text(content, encoding='utf-8')
        return path

    def compare(self, paths, expressions):
        book = Workbook(str(self.root / 'design.xlsx'), 'openpyxl', [
            sheet('1.Main', [('Name', 'Type', 'Range'), *((expr, '', value) for expr, value in expressions)])])
        return compare_book(None, catalog_headers(paths), book)[0]

    def test_no_include_still_links_typedefs_macros_and_variable(self):
        types = self.write('base.h', '#define UINT unsigned int\n#define WIDTH 3\n'
                           'typedef unsigned short u2;')
        header = self.write('device.h', 'typedef STRUCT { uint ready:width; U2 current; } Device_T;')
        source = self.write('device.c', 'Extern device_t can_chg; Extern CONST DEVICE_T * Const ptr;')
        results = self.compare([source, types, header],
            [('CAN_CHG . READY', '0~7'), ('PTR - > Ready', '0~1'), ('can_chg.current', '0~10')])
        self.assertEqual([r.status for r in results], ['MATCH', 'MISMATCH', 'SKIP'])
        self.assertEqual([r.expected for r in results], ['0~7', '0~7', ''])

    def test_struct_only_supports_dot_arrow_array_and_ordinary_member(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; int current; } can_chg_t;')
        results = self.compare([header], [('CAN_CHG . READY', '0~1'), ('can_chg_p - > ready', '0~0'),
            ('can_chg [ 0 ] . ready', None), ('can_chg.current', '0~10')])
        self.assertEqual([r.status for r in results], ['MATCH', 'MISMATCH', 'CREATE', 'SKIP'])
        self.assertEqual([r.candidate for r in results], [False, True, True, False])
        self.assertTrue(all('변수 선언 연결 없이' in r.reason for r in results))
        self.assertEqual(results[0].source, str(header.resolve()))

    def test_identical_candidate_types_and_ranges_can_be_compared(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:3; } can_chg_t;')
        duplicate = self.write('copy.h', 'typedef STRUCT { unsigned int ready:3; } can_chg_t;')
        results = self.compare([header, duplicate], [('can_chg.ready', '0~1')])
        self.assertEqual((results[0].status, results[0].expected), ('MISMATCH', '0~7'))
        self.assertIn('동일', results[0].reason)
        self.assertEqual(results[0].struct, 'can_chg_t')

    def test_conflicting_candidate_ranges_or_member_kinds_remain_review(self):
        for second in ('unsigned int ready:3;', 'int ready;', 'unsigned int ready:UNKNOWN;'):
            with self.subTest(second=second):
                header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } can_chg_t;')
                other = self.write('other.h', 'typedef STRUCT { ' + second + ' } can_chg_t;')
                result = self.compare([header, other], [('can_chg.ready', '0~1')])[0]
                self.assertEqual((result.status, result.candidate, result.expected), ('REVIEW', False, ''))
                self.assertIn('후보', result.reason)

    def test_existing_variable_link_takes_priority_over_global_member_candidates(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } A; '
                            'typedef STRUCT { unsigned int ready:3; } B; typedef A can_chg_t; B can_chg;')
        result = self.compare([header], [('can_chg.ready', '0~1')])[0]
        self.assertEqual((result.status, result.expected, result.struct), ('MISMATCH', '0~7', 'B'))
        self.assertNotIn('변수 선언 연결 없이', result.reason)

    def test_known_type_with_missing_member_is_not_rebound_to_another_structure(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int other:1; } A; '
                            'typedef STRUCT { unsigned int ready:3; } B; A can_chg;')
        result = self.compare([header], [('can_chg.ready', '0~1')])[0]
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
        self.assertIn('MEMBER_NOT_FOUND', result.reason)

    def test_missing_variable_and_member_are_still_explained(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } A;')
        result = self.compare([header], [('can_chg.missing', '0~1')])[0]
        self.assertEqual(result.status, 'REVIEW')
        self.assertIn('DECL_NOT_COLLECTED', result.reason)
        self.assertIn('대응하는 구조체 타입', result.reason)

    def test_diagnostic_trace_uses_the_same_struct_only_matching(self):
        from rangevalidator.ch_diagnostics import CHDiagnostics
        header = self.write('device.h', 'typedef STRUCT { unsigned int Ready:1; int current; } can_chg_t;')
        diagnostic = CHDiagnostics([header])
        trace = diagnostic.trace_expression('CAN_CHG_P - > READY', None, 'Main')
        self.assertEqual((trace['code'], trace['range']), ('TYPE_MATCH', [0, 1]))
        self.assertIn('변수 선언 연결 없이', trace['production_match']['reason'])
        ordinary = diagnostic.trace_expression('CAN_CHG . CURRENT', None, 'Main')
        self.assertEqual((ordinary['code'], ordinary['bitfield']), ('NOT_BITFIELD', False))

    def test_every_root_resolves_its_own_type_before_selecting_the_member(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } can_chg_t; '
                            'typedef Struct { unsigned int ready:5; int status; } device_t; '
                            'typedef struct { unsigned int value:3; } sensor_t;')
        results = self.compare([header], [('can_chg.ready', '0~1'), ('DEVICE.READY', '0~1'),
            ('sensor_p - > value', '0~7'), ('device.status', '0~10'), ('can_chg[0].ready', '0~1')])
        self.assertEqual([r.status for r in results], ['MATCH', 'MISMATCH', 'MATCH', 'SKIP', 'MATCH'])
        self.assertEqual([r.expected for r in results], ['0~1', '0~31', '0~7', '', '0~1'])

    def test_other_structs_member_is_never_used_when_root_type_is_missing(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } unrelated_t;')
        result = self.compare([header], [('can_chg.ready', '0~1')])[0]
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
        self.assertIn('대응하는 구조체 타입', result.reason)

    def test_other_structs_member_is_never_used_when_root_type_lacks_member(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int other:1; } can_chg_t; '
                            'typedef STRUCT { unsigned int ready:1; } unrelated_t;')
        result = self.compare([header], [('can_chg.ready', '0~1')])[0]
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))
        self.assertIn('MEMBER_NOT_FOUND', result.reason)
        self.assertIn('can_chg_t', result.reason)

    def test_aliases_and_struct_tags_can_supply_the_inferred_type(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int ready:1; } Internal; '
                            'typedef Internal can_chg_t; STRUCT sensor { unsigned int value:3; };')
        results = self.compare([header], [('can_chg.ready', '0~1'), ('sensor_p->value', '0~7')])
        self.assertEqual([r.status for r in results], ['MATCH', 'MATCH'])
        self.assertIn('can_chg_t', results[0].reason)

    def test_same_type_candidates_with_case_variants_of_member_can_agree(self):
        header = self.write('device.h', 'typedef STRUCT { unsigned int Ready:1; } can_chg_t;')
        duplicate = self.write('other.h', 'typedef STRUCT { unsigned int ready:1; } can_chg_t;')
        result = self.compare([header, duplicate], [('CAN_CHG.READY', '0~1')])[0]
        self.assertEqual((result.status, result.expected), ('MATCH', '0~1'))
