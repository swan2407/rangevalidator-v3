from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


TYPES = """
typedef struct { unsigned int active:1; unsigned int mode:3; } Device;
typedef struct { unsigned int active:3; } Sensor;
typedef struct { int active; } Normal;
typedef Device Alias;
typedef Device *Handle;
"""


class FunctionScopeTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def write(self, name, contents):
        path = self.root / name
        path.write_text(contents, encoding='utf-8')
        return path

    def compare(self, paths, book, sheets):
        workbook = Workbook(str(self.root / (book + '.xlsx')), 'com', [
            sheet(name, [('Name', 'Type', 'Range'), *((expr, '', value) for expr, value in entries)])
            for name, entries in sheets
        ])
        rows, _ = compare_book(None, catalog_headers(paths), workbook)
        return [(r.sheet, r.expression, r.status, r.expected, r.candidate) for r in rows]

    def test_parameters_and_locals_with_header_prototype(self):
        header = self.write('device.h', TYPES +
                            'void Process(Device, const Device *, Handle, Device *);')
        source = self.write('device.c', """
#include "device.h"
void Process(Device value, const Device *const ptr, Handle handle, Device array[])
{
    Alias local;
    Device *const local_p = &local;
    Device devices[3] = {0};
    Handle pointers[2];
}
""")
        name = '3. Process'
        expressions = [
            ('value.active', '0~1'), ('ptr->mode', '0~3'), ('handle->active', '0~1'),
            ('array[i].active', '0~1'), ('local.active', '0~1'), ('local_p->mode', '0~7'),
            ('devices[i].mode', '0~7'), ('pointers[i]->active', '0~1'),
        ]
        results = self.compare([header, source], 'device', [(name, expressions), ('Process', expressions)])
        self.assertEqual(results, [(name, expr, 'MISMATCH' if expr == 'ptr->mode' else 'MATCH',
                                    '0~7' if expr.endswith('mode') else '0~1', expr == 'ptr->mode')
                                   for expr, _ in expressions])

    def test_function_and_block_shadowing(self):
        source = self.write('scope.c', TYPES + """
Device state;
Device loop;
void First(void) { Sensor state; Device other; { Sensor other; Device unique; } Device good; }
void Second(void) { Device state; }
void Third(void) { for (Sensor loop; 0; ) { } }
""")
        results = self.compare([source], 'scope', [
            ('1.First', [('state.active', '0~7'), ('other.active', '0~1'), ('good.active', '0~1'), ('unique.active', '0~1')]),
            ('2.Second', [('state.active', '0~1'), ('other.active', '0~1')]),
            ('3.Third', [('loop.active', '0~1')]),
        ])
        self.assertEqual(results, [
            ('1.First', 'state.active', 'MATCH', '0~7', False),
            ('1.First', 'other.active', 'REVIEW', '', False),
            ('1.First', 'good.active', 'MATCH', '0~1', False),
            ('1.First', 'unique.active', 'MATCH', '0~1', False),
            ('2.Second', 'state.active', 'MATCH', '0~1', False),
            ('2.Second', 'other.active', 'REVIEW', '', False),
            ('3.Third', 'loop.active', 'REVIEW', '', False),
        ])

    def test_same_static_function_in_two_translation_units(self):
        first = self.write('first.c', TYPES + 'static void Run(void) { Device local; }')
        second = self.write('second.c', TYPES + 'static void Run(void) { Sensor local; }')
        sheets = [('12.Run', [('local.active', '0~1')])]
        for book, status, expected in [('first', 'MATCH', '0~1'),
                                       ('second', 'MISMATCH', '0~7'),
                                       ('unknown', 'REVIEW', '')]:
            self.assertEqual(self.compare([first, second], book, sheets),
                             [('12.Run', 'local.active', status, expected, status == 'MISMATCH')])

    def test_member_exclusion_precedes_access_operator(self):
        source = self.write('device.c', TYPES + """
Device device;
Normal normal;
void Process(void) { Normal local; Device good; float scalar; }
""")
        results = self.compare([source], 'device', [('1.Process', [
            ('device.active', '0~1'), ('device->active', '0~1'), ('normal.active', '0~1'),
            ('normal->active', '0~1'), ('local->active', '0~1'), ('unknown.active', '0~1'), ('scalar.active', '0~1'),
            ('normal.other', '0~1'), ('normal.mode', '0~7'), ('good.active + 1', '0~1'),
        ])])
        self.assertEqual(results, [
            ('1.Process', 'device.active', 'MATCH', '0~1', False),
            ('1.Process', 'device->active', 'REVIEW', '', False),
            ('1.Process', 'normal.active', 'SKIP', '', False),
            ('1.Process', 'normal->active', 'SKIP', '', False),
            ('1.Process', 'local->active', 'SKIP', '', False),
            ('1.Process', 'unknown.active', 'REVIEW', '', False),
            ('1.Process', 'normal.other', 'REVIEW', '', False),
            ('1.Process', 'normal.mode', 'REVIEW', '', False),
        ])

    def test_conflicting_prototype_blocks_but_conditional_definition_can_compare(self):
        header = self.write('device.h', TYPES + 'void Conflict(Sensor *);')
        source = self.write('device.c', """
#include "device.h"
void Conflict(Device *ptr) {}
#if FEATURE
void Conditional(Device *ptr) {}
#endif
""")
        results = self.compare([header, source], 'device', [
            (name, [('ptr->active', '0~0')]) for name in ('1.Conflict', '2.Conditional')
        ])
        self.assertEqual(results, [('1.Conflict', 'ptr->active', 'REVIEW', '', False),
                                   ('2.Conditional', 'ptr->active', 'MISMATCH', '0~1', True)])

    def test_unsupported_pointer_parameter_does_not_fall_back_to_global(self):
        source = self.write('device.c', TYPES + """
Device *ptr;
void Process(Device (*ptr)[2], Device value) {}
""")
        self.assertEqual(self.compare([source], 'device', [('1.Process', [
            ('ptr->active', '0~0'), ('value.active', '0~1')
        ])]), [
            ('1.Process', 'ptr->active', 'REVIEW', '', False),
            ('1.Process', 'value.active', 'MATCH', '0~1', False),
        ])
