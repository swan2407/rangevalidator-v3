"""Build contexts must select real branches without mixing translation units."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.build_options import load_build_options, BuildOptionsError, preprocess_sources
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


def xml(macros='', includes=''):
    return ('<cdtprojectproperties><section name="org.eclipse.cdt.internal.ui.wizards.settingswizards.Macros">'
            '<language name="C Source File">' + macros + '</language></section>'
            '<section name="org.eclipse.cdt.internal.ui.wizards.settingswizards.IncludePaths">'
            '<language name="C Source File">' + includes + '</language></section></cdtprojectproperties>')


def macro(name, value=''):
    return f'<macro><name>{name}</name><value>{value}</value></macro>'


BPS = '''#ifndef BPS_H
#define BPS_H
typedef unsigned int uint;
#ifdef OPT_PRESSURE_SENSOR
typedef struct { uint BpsPressureData:12; uint BpsFwMajorVer:4; } BpsDataStateCycle_t;
#else
typedef struct { unsigned short BpsPressureData; unsigned char BpsFwMajorVer; } BpsDataStateCycle_t;
#endif
void pressure_update_status_flags(BpsDataStateCycle_t *data);
#endif
'''
C = '#include "bps.h"\nvoid pressure_update_status_flags(BpsDataStateCycle_t *data) {}\n'


class BuildOptionsTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def options(self, macros='', includes=''):
        path = self.root / 'build.xml'
        path.write_text(xml(macros, includes), encoding='utf-8')
        return load_build_options(path)

    def files(self, sources):
        paths = []
        for name, value in sources.items():
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(value, encoding='utf-8')
            paths.append(path)
        return paths

    def compare(self, catalog, name='pressure_update_status_flags', expression='data->BpsPressureData',
                value='0~4095', source='pressure'):
        book = Workbook(str(self.root / (source + '_Detail.xlsx')), 'openpyxl',
                        [sheet('1. ' + name, [('Name', 'Type', 'Range'), (expression, 'uint', value)])])
        return compare_book(None, catalog, book)[0][0]

    def test_photographed_shape_selected_and_original_location(self):
        paths = self.files({'bps.h': BPS, 'pressure.c': C})
        row = self.compare(catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR'))))
        self.assertEqual((row.status, row.expected, row.source_line), ('MATCH', '0~4095', 5))
        self.assertEqual(row.source, str(self.root / 'bps.h'))
        self.assertIn('build.xml', row.reason)
        self.assertIn('pressure.c:2', row.reason)

    def test_disabled_option_selects_ordinary_member_even_with_bit_comment(self):
        paths = self.files({'bps.h': BPS.replace('unsigned short BpsPressureData;', 'unsigned short BpsPressureData; /* 12bit */'), 'pressure.c': C})
        row = self.compare(catalog_headers(paths, build_options=self.options()))
        self.assertEqual(row.status, 'SKIP')
        self.assertFalse(row.candidate)

    def test_without_options_keeps_legacy_review(self):
        paths = self.files({'bps.h': BPS, 'pressure.c': C})
        self.assertEqual(self.compare(catalog_headers(paths)).status, 'REVIEW')

    def test_numeric_elif_nested_and_function_macro(self):
        opts = self.options(macro('MODEL', '2') + macro('FEATURE'))
        paths = self.files({'unit.c': '''#define DOUBLE(x) ((x) * 2)
#if MODEL == 1
typedef struct { unsigned int value:1; } Item;
#elif MODEL == 2 && defined(FEATURE)
#ifndef DISABLED
typedef struct { unsigned int value:DOUBLE(3); } Item;
#endif
#else
typedef struct { unsigned int value:2; } Item;
#endif
void Probe(Item *data) {}
'''})
        row = self.compare(catalog_headers(paths, build_options=opts), 'Probe', 'data->value', '0~63', 'unit')
        self.assertEqual((row.status, row.expected), ('MATCH', '0~63'))

    def test_define_undef_and_inactive_error(self):
        paths = self.files({'unit.c': '''#undef FEATURE
#if defined(FEATURE)
#error not active
#include "missing.h"
#else
#define WIDTH 3
#endif
typedef struct { unsigned int value:WIDTH; } Item;
void Probe(Item *data) {}
'''})
        row = self.compare(catalog_headers(paths, build_options=self.options(macro('FEATURE'))),
                           'Probe', 'data->value', '0~7', 'unit')
        self.assertEqual(row.status, 'MATCH')

    def test_include_order_and_relative_xml_include_path(self):
        paths = self.files({'inc/options.h': '#define WIDTH 5\n', 'unit.c': '#include <options.h>\n'
                            'typedef struct { unsigned int value:WIDTH; } Item;\nvoid Probe(Item *data) {}'})
        opts = self.options(includes='<includepath workspace_path="false">inc</includepath>')
        row = self.compare(catalog_headers(paths[1:], build_options=opts), 'Probe', 'data->value', '0~31', 'unit')
        self.assertEqual(row.status, 'MATCH')

    def test_registered_unique_header_fallback(self):
        paths = self.files({'inc/bps.h': BPS, 'src/pressure.c': C})
        row = self.compare(catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR'))))
        self.assertEqual(row.status, 'MATCH')

    def test_macro_state_isolated_between_c_files(self):
        paths = self.files({'bps.h': BPS,
                            'pressure.c': '#define OPT_PRESSURE_SENSOR\n' + C,
                            'other.c': '#include "bps.h"\nvoid Other(BpsDataStateCycle_t *data) {}'})
        cat = catalog_headers(paths, build_options=self.options())
        self.assertEqual(self.compare(cat).status, 'MATCH')
        self.assertEqual(self.compare(cat, 'Other', source='other').status, 'SKIP')

    def test_same_header_reincluded_with_changed_macro_without_guard(self):
        paths = self.files({'make.h': 'typedef struct { unsigned int value:WIDTH; } NAME;\n',
                            'unit.c': '#define WIDTH 2\n#define NAME First\n#include "make.h"\n'
                            '#undef WIDTH\n#undef NAME\n#define WIDTH 5\n#define NAME Second\n'
                            '#include "make.h"\nvoid Probe(Second *data) {}'})
        row = self.compare(catalog_headers(paths, build_options=self.options()), 'Probe', 'data->value', '0~31', 'unit')
        self.assertEqual(row.status, 'MATCH')

    def test_unrelated_translation_unit_cannot_supply_missing_variable_or_type(self):
        paths = self.files({'other.c': 'typedef struct { unsigned int value:3; } Missing;\nMissing data;\n',
                            'unit.c': 'void Probe(void) {}'})
        cat = catalog_headers(paths, build_options=self.options())
        for expression in ('data.value', 'Missing.value'):
            self.assertEqual(self.compare(cat, 'Probe', expression, '0~7', 'unit').status, 'REVIEW')

    def test_header_guard_and_pragma_once(self):
        paths = self.files({'bps.h': BPS.replace('#ifndef BPS_H\n#define BPS_H', '#pragma once').rsplit('#endif', 1)[0],
                            'pressure.c': '#include "bps.h"\n' + C})
        self.assertEqual(self.compare(catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR')))).status, 'MATCH')

    def test_missing_include_blocks_only_affected_source(self):
        paths = self.files({'bps.h': BPS, 'pressure.c': C,
                            'bad.c': '#include "missing.h"\nvoid Bad(void) {}'})
        cat = catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR')))
        self.assertEqual(self.compare(cat).status, 'MATCH')
        row = self.compare(cat, 'Bad', source='bad')
        self.assertEqual((row.status, row.expected, row.candidate), ('REVIEW', '', False))
        self.assertIn('missing.h', row.reason)
        unknown = self.compare(cat, source='unknown')
        self.assertEqual(unknown.status, 'REVIEW')

    def test_ambiguous_include_is_not_chosen_arbitrarily(self):
        paths = self.files({'one/bps.h': BPS, 'two/bps.h': BPS, 'pressure.c': C})
        row = self.compare(catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR'))))
        self.assertEqual(row.status, 'REVIEW')
        self.assertIn('중복 헤더', row.reason)

    def test_empty_numeric_macro_is_not_assumed_one(self):
        paths = self.files({'unit.c': '#if FEATURE\nint a;\n#endif\n'})
        bundle = preprocess_sources(paths, self.options(macro('FEATURE')))
        self.assertTrue(bundle.errors)

    def test_unknown_compiler_macro_is_not_assumed_zero(self):
        paths = self.files({'unit.c': '#if __GNUC__\nint a;\n#endif\n'})
        bundle = preprocess_sources(paths, self.options())
        self.assertIn('__GNUC__', next(iter(bundle.errors.values())))

    def test_error_and_unsupported_condition_do_not_certify(self):
        for code in ('#error invalid build\n', '#if UNKNOWN_FUNC(2)\nint a;\n#endif\n',
                     '#if 1\nint a;\n', '#pragma push_macro("FEATURE")\n'):
            with self.subTest(code=code):
                bundle = preprocess_sources(self.files({'unit.c': code}), self.options())
                self.assertTrue(bundle.errors)
                self.assertEqual(bundle.sources[str(self.root / 'unit.c')], '')

    def test_function_signature_conflict_stays_review(self):
        paths = self.files({'bps.h': BPS, 'pressure.c': C.replace('BpsDataStateCycle_t *data', 'BpsDataStateCycle_t **data')})
        row = self.compare(catalog_headers(paths, build_options=self.options(macro('OPT_PRESSURE_SENSOR'))))
        self.assertEqual(row.status, 'REVIEW')

    def test_only_c_language_macros_loaded_and_conflicts_rejected(self):
        opts = self.options(macro('A') + macro('B', '3'))
        self.assertEqual(dict(opts.macros), {'A': '', 'B': '3'})
        with self.assertRaises(BuildOptionsError):
            self.options(macro('A', '1') + macro('A', '2'))
        with self.assertRaises(BuildOptionsError):
            self.options(includes='<includepath workspace_path="true">/project</includepath>')
        path = self.root / 'wrong.xml'
        for data in ('<other/>', '<cdtprojectproperties/>', 'not xml'):
            path.write_text(data)
            with self.assertRaises(BuildOptionsError):
                load_build_options(path)

    def test_session_build_options_are_signed(self):
        from rangevalidator.workflow import Session
        session = Session((), ())
        session.build_options = self.options(macro('A')).receipt()
        session._signature = session.signature()
        session.assert_unchanged()
        session.build_options['macros']['A'] = '9'
        with self.assertRaises(RuntimeError):
            session.assert_unchanged()


if __name__ == '__main__':
    unittest.main()
