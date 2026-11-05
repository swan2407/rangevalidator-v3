from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from rangevalidator.common import Common, parse_declaration
from rangevalidator.engine import catalog_headers, compare_book, build_bitfield_catalog
from rangevalidator.models import Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet


class CatalogComparisonTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.header = self.write('types.h', """
typedef unsigned char u1;
typedef struct { u1 active:1; u1 mode:3; int status; } Device_t;
typedef struct { u1 active:3; u1 status:5; } Sensor_t;
typedef struct { char sign:2; u1 width:UNKNOWN; } Uncertain_t;
typedef Device_t DeviceAlias_t;
typedef struct { u1 mode:3; } guessed_t;
""")
        self.source = self.write('motor.c', """
#include "types.h"
extern DeviceAlias_t device;
extern DeviceAlias_t *device_p;
extern DeviceAlias_t device_array[2];
extern DeviceAlias_t *pointer_array[2];
extern DeviceAlias_t **double_pointer;
extern DeviceAlias_t matrix[2][3];
extern Sensor_t sensor;
extern Uncertain_t uncertain;
#if FEATURE
extern DeviceAlias_t conditional;
#endif
""")
        self.paths = [self.header, self.source]

    def write(self, name, text):
        path = self.root / name
        path.write_text(text, encoding='utf-8')
        return path

    def workbook(self, expressions, path='motor.xlsx'):
        data = [('Name', 'Type', 'Range'), *((expr, '', value) for expr, value in expressions)]
        return Workbook(str(self.root / path), 'com', [
            sheet('Design', data), sheet('commonparameter', data), sheet('1.Main', data),
            sheet('2.UnknownLayout', [('Unknown layout',)]),
        ])

    def test_filters_candidates_and_preserves_range_decisions(self):
        rows = [
            ('device.active', '0~1'), ('device.mode', '0~2'), ('sensor.active', '0~7'),
            ('sensor.status', '0~31'), ('device.status', '0~31'), ('device.normal', '0~2'),
            ('count', '0~1'), ('device_array[i]', '0~1'), ('get_status()', '0~1'),
            ('device.active + 1', '0~1'), ('device.active = 1', '0~1'),
            ('description device.active', '0~1'), ('device.active[0]', '0~1'),
            ('device_array[i + 1].active', '0~1'), ('device_p -> active', '0~1'),
            ('device_array [ i ] . active', '0~1'), ('pointer_array[0]->active', '0~1'),
            ('unknown.active', '0~1'), ('unknown.temperature', '0~1'),
            ('uncertain.width', '0~1'), ('uncertain.sign', '0~1'), ('guessed.mode', '0~7'),
            ('conditional.active', '0~1'), ('double_pointer->active', '0~1'), ('matrix[i].active', '0~1'),
        ]
        results, coverage = compare_book(None, catalog_headers(self.paths), self.workbook(rows))
        self.assertEqual({r.sheet for r in results}, {'1.Main'})
        self.assertEqual([item[:2] for item in coverage],
                         [('1.Main', 'TARGET'), ('2.UnknownLayout', 'REVIEW')])
        self.assertEqual([(r.expression, r.status, r.expected, r.candidate) for r in results], [
            ('device.active', 'MATCH', '0~1', False),
            ('device.mode', 'MISMATCH', '0~7', True),
            ('sensor.active', 'MATCH', '0~7', False),
            ('sensor.status', 'MATCH', '0~31', False),
            ('device_p -> active', 'MATCH', '0~1', False),
            ('device_array [ i ] . active', 'MATCH', '0~1', False),
            ('pointer_array[0]->active', 'MATCH', '0~1', False),
            ('unknown.active', 'REVIEW', '', False),
            ('uncertain.width', 'REVIEW', '', False),
            ('uncertain.sign', 'REVIEW', '', False),
            ('guessed.mode', 'REVIEW', '', False),
            ('conditional.active', 'REVIEW', '', False),
            ('double_pointer->active', 'REVIEW', '', False),
            ('matrix[i].active', 'REVIEW', '', False),
        ])

    def test_common_alias_is_ignored_even_in_real_source_context(self):
        good = parse_declaration('motor.c', 'via_common', 'DeviceAlias_t')
        good.sheet, good.row = 'Global', 4
        missing = parse_declaration('missing.c', 'orphan', 'DeviceAlias_t')
        common = Common(None, [good, missing])
        book = self.workbook([('via_common.active', '0~1'), ('orphan.active', '0~1')], 'unmapped.xlsx')
        results, _ = compare_book(common, catalog_headers(self.paths), book)
        self.assertEqual([(r.status, r.common_sheet, r.common_row, r.candidate) for r in results],
                         [('REVIEW', '', 0, False), ('REVIEW', '', 0, False)])

    def test_context_distinguishes_conflicting_source_declarations(self):
        other = self.write('other.c', '#include "types.h"\nextern Sensor_t device;')
        catalog = catalog_headers([*self.paths, other])
        for path, status, expected in [('motor.xlsx', 'MATCH', '0~1'),
                                       ('other.xlsx', 'MISMATCH', '0~7'),
                                       ('unknown.xlsx', 'REVIEW', '')]:
            results, _ = compare_book(None, catalog, self.workbook([('device.active', '0~1')], path))
            self.assertEqual([(r.status, r.expected) for r in results], [(status, expected)])

    def test_range_formula_and_non_anchor_merge_still_block_candidates(self):
        book = self.workbook([('device.active', '0~0'), ('device.mode', '0~0')])
        target = book.sheets[2]
        target.cells['C2'] = replace(target.cells['C2'], is_formula=True, formula='="0~0"')
        target.cells['C3'] = replace(target.cells['C3'], merged_range='B3:C3', merge_anchor='B3')
        results, _ = compare_book(None, catalog_headers(self.paths), book)
        self.assertEqual([(r.status, r.candidate) for r in results],
                         [('REVIEW', False), ('REVIEW', False)])

    def test_conditional_variable_after_a_spliced_line(self):
        source = self.write('spliced.c', '#include "types.h"\n#define OTHER ' + chr(92)
                            + '\n  4\n#if FEATURE\nextern Device_t flag;\n#endif')
        catalog = catalog_headers([self.header, source])
        results, _ = compare_book(None, catalog, self.workbook([('flag.active', '0~0')], 'spliced.xlsx'))
        self.assertEqual([(r.status, r.candidate) for r in results], [('REVIEW', False)])

    def test_reuses_catalog_for_multiple_documents(self):
        details = [self.write('motor.xlsx', ''), self.write('second.xlsx', '')]
        with patch('rangevalidator.engine.build_bitfield_catalog', wraps=build_bitfield_catalog) as build:
            with patch('rangevalidator.workflow.read_auto', side_effect=lambda path, **_:
                       self.workbook([('device.active', '0~1')], Path(path).name)):
                session = run_comparison(None, self.paths, details, 'auto')
        build.assert_called_once()
        self.assertEqual(session.errors, [])
        self.assertEqual([r.status for r in session.results], ['MATCH', 'MATCH'])
