"""Regression cases for photographed declaration shapes, using synthetic sources."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


class StructMembersTests(unittest.TestCase):
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def write(self, name, content):
        path = self.root / name
        path.write_text(content, encoding='utf-8')
        return path

    def compare(self, paths, function='Main', expressions=None):
        expressions = expressions or ['can_chg.obc_ready', 'can_chg.qcs_maxcurr',
                                      'can_chg_p->obc_ready', 'can_chg_p->qcs_maxcurr']
        book = Workbook(str(self.root / 'device.xlsx'), 'openpyxl', [
            sheet('1.' + function, [('Name', 'Type', 'Range'),
                                   *((expr, '', '0~1') for expr in expressions)])])
        return compare_book(None, catalog_headers(paths), book)[0]

    def test_guard_forms_and_included_keyword_macros_with_trailing_extern(self):
        types = self.write('platform.h', '#ifndef PLATFORM_H\n#define PLATFORM_H 1\n'
                           '#define STRUCT struct\n#define uint unsigned int\n'
                           'typedef unsigned short u2;\n#endif\n')
        for opening in ('#ifndef DEVICE_H', '#if !defined(DEVICE_H)', '#if !defined DEVICE_H'):
            for prefix in ('', '#pragma once\n'):
                with self.subTest(opening=opening, prefix=prefix):
                    header = self.write('device.h', prefix + opening + ' // guard\n#define DEVICE_H 1\n'
                        '#include "platform.h"\ntypedef STRUCT { uint obc_ready:1; u2 qcs_maxcurr; } can_chg_t;\n'
                        'extern can_chg_t can_chg;\nextern const can_chg_t * const can_chg_p;\n'
                        '#ifdef OPT_SDV\nextern int optional;\n#endif\nvoid Main(void);\n#endif\n')
                    source = self.write('device.c', '#include "device.h"\nvoid Main(void) {}\n')
                    results = self.compare([source, header, types])
                    self.assertEqual([r.status for r in results], ['MATCH', 'SKIP', 'MATCH', 'SKIP'])
                    self.assertEqual([r.expected for r in results], ['0~1', '', '0~1', ''])
                    self.assertTrue(all(not r.candidate for r in results))
                    self.assertIn('일반 멤버', results[1].reason)
                    self.assertNotIn('미확정', results[1].reason)

    def test_real_conditions_allow_declared_bit_width_and_label_known_ordinary(self):
        for target in ('variable', 'type', 'function', 'prototype'):
            with self.subTest(target=target):
                parts = {'type': 'typedef struct { unsigned int obc_ready:1; int qcs_maxcurr; } T;',
                         'variable': 'extern T can_chg; extern const T * const can_chg_p;',
                         'prototype': 'void Main(void);', 'function': 'void Main(void) {}'}
                parts[target] = '#if FEATURE\n' + parts[target] + '\n#endif'
                source = self.write('device.c', '\n'.join(parts.values()))
                results = self.compare([source])
                self.assertEqual([r.status for r in results], ['MATCH', 'SKIP', 'MATCH', 'SKIP'])
                self.assertNotIn('CONDITIONAL', results[0].reason)
                self.assertIn('일반 멤버', results[1].reason)
                self.assertNotIn('빌드 조건 미확정', results[1].reason)
                self.assertEqual([r.expected for r in results], ['0~1', '', '0~1', ''])
                self.assertTrue(all(not r.candidate for r in results))

    def test_branch_with_ordinary_and_bitfield_cannot_be_labeled_ordinary(self):
        source = self.write('device.c', '#if FEATURE\n'
            'typedef struct { int qcs_maxcurr; } T;\n#else\n'
            'typedef struct { unsigned int qcs_maxcurr:3; } T;\n#endif\nT can_chg;')
        results = self.compare([source], expressions=['can_chg.qcs_maxcurr'])
        self.assertEqual([(r.status, r.candidate) for r in results], [('REVIEW', False)])

    def test_member_condition_with_conflicting_kinds_stays_review(self):
        source = self.write('device.c', 'typedef struct {\n#if FEATURE\nint qcs_maxcurr;\n'
            '#else\nunsigned int qcs_maxcurr:3;\n#endif\n} T;\nT can_chg;')
        results = self.compare([source], expressions=['can_chg.qcs_maxcurr'])
        self.assertEqual([(r.status, r.candidate) for r in results], [('REVIEW', False)])

    def test_feature_wrapper_is_not_a_whole_file_guard(self):
        source = self.write('device.c', '#if !defined(FEATURE)\n#define DIFFERENT 1\n'
            'typedef struct { unsigned int obc_ready:1; } T; T can_chg;\n#endif\n')
        result = self.compare([source], expressions=['can_chg.obc_ready'])[0]
        self.assertEqual((result.status, result.expected, result.candidate), ('MATCH', '0~1', False))
        self.assertEqual(catalog_headers([source]).bitfields.lookup_variable(source, 'can_chg').status,
                         'CONDITIONAL')
        self.assertNotIn('CONDITIONAL', result.reason)

    def test_struct_keyword_does_not_require_macro_evidence(self):
        for prefix in ('', '#if FEATURE\n#define STRUCT struct\n#endif\n',
                       '#define STRUCT struct\n#undef STRUCT\n',
                       '#define STRUCT struct\n#include "not_registered.h"\n'):
            with self.subTest(prefix=prefix):
                source = self.write('device.c', prefix +
                    'typedef STRUCT { unsigned int obc_ready:1; int qcs_maxcurr; } T; T can_chg;')
                results = self.compare([source], expressions=['can_chg.obc_ready', 'can_chg.qcs_maxcurr'])
                self.assertEqual([r.status for r in results], ['MATCH', 'SKIP'])

    def test_struct_keyword_does_not_require_an_included_macro(self):
        self.write('platform.h', '#define STRUCT struct\n')
        source = self.write('device.c', '#include "platform.h"\n'
            'typedef STRUCT { unsigned int obc_ready:1; } T; T can_chg;')
        result = self.compare([source], expressions=['can_chg.obc_ready'])[0]
        self.assertEqual((result.status, result.candidate), ('MATCH', False))

    def test_mixed_case_keywords_typedefs_variables_functions_and_members(self):
        source = self.write('device.c', 'TypeDef\tStruct { UnSigned\n INT Obc_Ready : 1; '
            'unsigned SHORT Qcs_MaxCurr; } Can_Chg_T;\n'
            'EXTERN can_chg_t CAN_CHG;\nExtern CONST CAN_CHG_T * Const CAN_CHG_P;\n'
            'Void Main(VOID) {}')
        results = self.compare([source], function='MAIN')
        self.assertEqual([r.status for r in results], ['MATCH', 'SKIP', 'MATCH', 'SKIP'])
        self.assertEqual(results[0].member, 'Obc_Ready')
        self.assertEqual(results[0].struct, 'Can_Chg_T')
        self.assertIn('Qcs_MaxCurr', results[1].reason)

    def test_case_insensitive_parameter_binding_and_width_macro(self):
        header = self.write('types.h', '#DEFINE WIDTH 3\n'
            'TYPEDEF STRUCT { UNSIGNED INT Ready:width; } Device_T;\n'
            'VOID Update(DEVICE_T *);')
        source = self.write('device.c', '#INCLUDE "TYPES.H"\n'
            'VOID Update(device_t *PTR) {}')
        result = self.compare([header, source], function='UPDATE', expressions=['ptr -> READY'])[0]
        self.assertEqual((result.status, result.expected, result.candidate), ('MISMATCH', '0~7', True))

    def test_exact_spelling_wins_and_case_collisions_remain_visible(self):
        source = self.write('device.c', 'typedef STRUCT { unsigned int Ready:1; unsigned int ready:3; } T; '
            'T obj; T Obj;')
        results = self.compare([source], expressions=['obj.Ready', 'obj.ready', 'obj.READY', 'OBJ.Ready'])
        self.assertEqual([r.status for r in results], ['MATCH', 'MISMATCH', 'REVIEW', 'REVIEW'])
        self.assertIn('대소문자', results[2].reason)
        self.assertTrue(all(not r.candidate for r in results[2:]))

    def test_case_colliding_type_names_are_not_arbitrarily_selected(self):
        source = self.write('device.c', 'typedef STRUCT { unsigned int ready:1; } Device; '
            'typedef STRUCT { unsigned int ready:3; } device; DEVICE obj;')
        result = self.compare([source], expressions=['obj.ready'])[0]
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))

    def test_spaced_member_operators_and_line_breaks_in_both_excel_formats(self):
        from openpyxl import Workbook as ExcelWorkbook
        from rangevalidator.workflow import run_comparison

        source = self.write('device.c', 'typedef STRUCT { unsigned int obc_ready:1; int qcs_maxcurr; } T; '
            'T can_chg; const T * const can_chg_p; T arr[2];')
        expressions = ['CAN_CHG . OBC_READY', 'can_chg_p \t- \n> Obc_Ready',
                       'ARR [ \n0\t ] . OBC_READY', 'CAN_CHG_P - > QCS_MAXCURR',
                       'can_chg_p - > obc_ready']
        for form in ('A', 'B'):
            with self.subTest(form=form):
                path = self.root / f'design_{form}.xlsx'
                book = ExcelWorkbook()
                book.active.title = '1.Main'
                book.active.append(['INPUT / OUTPUT', 'PROPERTY', ' NAME ', 'DATA TYPE', 'SCOPE']
                                   if form == 'A' else [' NAME ', ' TYPE ', ' RANGE '])
                for i, expr in enumerate(expressions):
                    row = [expr, 'UNSIGNED\n  INT', '0~1'] if i < 4 else [expr, None, None]
                    book.active.append([None, None, *row] if form == 'A' else row)
                book.save(path)
                book.close()
                session = run_comparison(None, [source], [path], 'openpyxl')
                self.assertEqual(session.errors, [])
                self.assertEqual([r.status for r in session.results], ['MATCH', 'MATCH', 'MATCH', 'SKIP', 'CREATE'])
                self.assertEqual([r.expression for r in session.results], expressions)
                self.assertNotIn('설계서 Type', session.results[0].reason)
                self.assertEqual([r.cell for r in session.results],
                                 [f'{"E" if form == "A" else "C"}{n}' for n in range(2, 7)])

    def test_conditional_local_cannot_hide_global_bitfield(self):
        source = self.write('device.c', 'typedef struct { int qcs_maxcurr; } Normal;\n'
            'typedef struct { unsigned int qcs_maxcurr:2; } Bits; Bits can_chg;\n'
            'void Main(void) {\n#if FEATURE\nNormal can_chg;\n#endif\n}\n')
        result = self.compare([source], expressions=['can_chg.qcs_maxcurr'])[0]
        self.assertEqual((result.status, result.candidate), ('REVIEW', False))

    def test_ordinary_members_are_counted_filtered_and_reported(self):
        from openpyxl import Workbook as ExcelWorkbook, load_workbook
        from rangevalidator.workflow import run_comparison
        from rangevalidator.report import create_report
        from rangevalidator.result_index import ResultIndex

        source = self.write('device.c', 'typedef struct { int qcs_maxcurr; } T; T can_chg;')
        path = self.root / 'device.xlsx'
        book = ExcelWorkbook()
        book.active.title = '1.Main'
        book.active.append(['Name', 'Type', 'Range'])
        book.active.append(['can_chg.qcs_maxcurr', 'int', None])
        book.save(path)
        book.close()
        original = path.read_bytes()
        session = run_comparison(None, [source], [path], 'openpyxl')
        index = ResultIndex(session)
        key = next(iter(index.files))
        self.assertEqual(index.counts['SKIP'], 1)
        self.assertEqual(index.visible(key, '1.Main', '일반 멤버', ''), session.results)
        self.assertFalse(session.results[0].candidate)
        self.assertEqual(session.results[0].expected, '')
        report_path = self.root / 'report.xlsx'
        create_report(session, {}, report_path)
        report = load_workbook(report_path)
        try:
            self.assertEqual(report['Results']['K2'].value, '일반 멤버 (SKIP)')
            self.assertIn('일반 멤버', report['Results']['M2'].value)
            self.assertEqual(report['Results']['L2'].value, 'NOT_APPLICABLE')
            self.assertEqual(report['Review'].max_row, 1)
            self.assertEqual(report['Mismatch'].max_row, 1)
        finally:
            report.close()
        self.assertEqual(original, path.read_bytes())
