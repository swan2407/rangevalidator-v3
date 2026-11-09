from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.bitfield_catalog import build_bitfield_catalog
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


class VariableBindingTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def source(self, name, text):
        path = self.root / name
        path.write_text(text, encoding='utf-8')
        return path

    def compare(self, sources, name, expressions):
        rows = [('Name', 'Type', 'Range'), *((e, '', value) for e, value in expressions)]
        workbook = Workbook(str(self.root / (name + '.xlsx')), 'com', [sheet('1.Main', rows)])
        results, _ = compare_book(None, catalog_headers(sources), workbook)
        return [(r.expression, r.status, r.expected, r.candidate) for r in results]

    def test_forward_declaration_extern_definition_and_access(self):
        header = self.source('device.h', """
typedef unsigned char u1;
typedef unsigned short u2;
typedef unsigned int u4;
typedef u1 U8;
typedef struct DeviceTag Device_t;
typedef Device_t DeviceAlias_t;
typedef DeviceAlias_t AnotherDevice_t;
typedef AnotherDevice_t *DevicePtr_t;
typedef DevicePtr_t const ConstHandle;
extern Device_t device;
extern const volatile Device_t *const device_p;
extern DevicePtr_t handles[];
""")
        source = self.source('device.c', """
#include "device.h"
struct DeviceTag { U8 active:1; u2 mode:3; u4 flags:4; };
Device_t device;
const volatile Device_t *const device_p = &device;
AnotherDevice_t device_array[10];
DevicePtr_t handles[2];
ConstHandle handle;
DevicePtr_t *indirect;
struct DeviceTag initialized = { 0 };
struct DeviceTag factory(void) { Device_t local; return local; }
Device_t after_factory;
""")
        paths = [header, source]
        catalog = build_bitfield_catalog(paths)
        self.assertEqual(catalog, build_bitfield_catalog(reversed(paths)))
        self.assertEqual(catalog.lookup(source, 'DevicePtr_t').pointer_depth, 1)
        self.assertEqual(catalog.lookup_variable(source, 'local').status, 'MISSING')
        pointer = catalog.lookup_variable(source, 'device_p')
        self.assertEqual(pointer.type_link.qualifiers, (('const', 'volatile'), ('const',)))
        self.assertEqual({d.kind for d in pointer.declarations}, {'DECLARATION', 'DEFINITION'})
        self.assertEqual({d.storage for d in pointer.declarations}, {'extern', ''})
        self.assertEqual({Path(d.origin.file) for d in pointer.declarations},
                         {header.resolve(), source.resolve()})
        self.assertEqual(catalog.lookup_variable(source, 'device_array').declarations[0].arrays, ('10',))
        self.assertEqual(catalog.lookup_variable(source, 'indirect').type_link.pointer_depth, 2)
        self.assertEqual(catalog.lookup_variable(source, 'handle').type_link.qualifiers, ((), ('const',)))
        self.assertEqual(catalog.lookup_variable(source, 'handle').type_link.definition,
                         catalog.lookup_variable(source, 'device').type_link.definition)
        self.assertEqual(self.compare(paths, 'device', [
            ('device.active', '0~1'), ('device_p->mode', '0~0'),
            ('device_array[i].active', '0~1'), ('handles[0]->flags', '0~15'),
            ('handle->active', '0~1'), ('initialized.active', '0~1'),
            ('after_factory.active', '0~1'), ('indirect->active', '0~1'),
            ('device->active', '0~1'), ('device_p.active', '0~1'), ('local.active', '0~1'),
        ]), [
            ('device.active', 'MATCH', '0~1', False), ('device_p->mode', 'MISMATCH', '0~7', True),
            ('device_array[i].active', 'MATCH', '0~1', False), ('handles[0]->flags', 'MATCH', '0~15', False),
            ('handle->active', 'MATCH', '0~1', False), ('initialized.active', 'MATCH', '0~1', False),
            ('after_factory.active', 'MATCH', '0~1', False), ('indirect->active', 'REVIEW', '', False),
            ('device->active', 'REVIEW', '', False), ('device_p.active', 'REVIEW', '', False),
            ('local.active', 'REVIEW', '', False),
        ])

    def test_empty_macro_restores_struct_and_variable_declarations_in_source_order(self):
        source = self.source('device.c', """\
extern const DeviceB_t * cinst before;
#define cinst /* intentionally empty */
cinst typedef struct { unsigned int state : 5; } DeviceB_t;
extern const DeviceB_t * cinst can_data1;
#if OPTION
#define cinst volatile
#endif
extern const DeviceB_t * cinst conditional;
#define cinst
extern const DeviceB_t * cinst restored;
#include "unknown.h"
extern const DeviceB_t * cinst after_include;
#define cinst
#undef cinst
extern const DeviceB_t * cinst after_undef;
#define cinst()
extern const DeviceB_t * cinst function_like;
""")
        catalog = build_bitfield_catalog([source])
        binding = catalog.lookup_variable(source, 'can_data1')
        self.assertEqual(binding.status, 'STRUCT')
        self.assertEqual(binding.type_link.definition.line, 3)
        self.assertEqual(binding.declarations[0].origin.line, 4)
        self.assertEqual(binding.type_link.qualifiers, (('const',), ()))
        unresolved = ('before', 'conditional', 'after_include', 'after_undef', 'function_like')
        expressions = [('can_data1->state', '0~7'), ('restored->state', '0~31')]
        expressions.extend((f'{name}->state', '0~31') for name in unresolved)
        self.assertEqual(self.compare([source], 'device', expressions), [
            ('can_data1->state', 'MISMATCH', '0~31', True),
            ('restored->state', 'MATCH', '0~31', False),
            *((f'{name}->state', 'REVIEW', '', False) for name in unresolved),
        ])

    def test_internal_linkage_and_distinct_translation_units(self):
        first = self.source('first.c', """
typedef struct { unsigned int status:2; } Local;
static Local state;
Local external_state;
""")
        second = self.source('second.c', """
typedef struct { unsigned int status:5; } Local;
static Local state;
Local external_state;
static Local private_only;
struct { int status; } normal;
""")
        paths = [first, second]
        self.assertEqual(self.compare(paths, 'first', [('state.status', '0~3')]),
                         [('state.status', 'MATCH', '0~3', False)])
        self.assertEqual(self.compare(paths, 'second', [('state.status', '0~31'), ('normal.status', '0~0')]),
                         [('state.status', 'MATCH', '0~31', False), ('normal.status', 'SKIP', '', False)])
        self.assertEqual(self.compare(paths, 'unknown', [
            ('state.status', '0~0'), ('external_state.status', '0~0'),
            ('private_only.status', '0~0'), ('unknown.status', '0~0'),
        ]), [(name + '.status', 'REVIEW', '', False)
             for name in ('state', 'external_state', 'private_only', 'unknown')])

    def test_shared_extern_conflict_and_pointer_typedef_conflict(self):
        header = self.source('types.h', """
typedef struct { unsigned int status:2; } A;
typedef struct { unsigned int status:5; } B;
extern A value;
typedef A *Handle;
typedef A **Handle;
typedef Cycle2 Cycle1;
typedef Cycle1 Cycle2;
""")
        source = self.source('first.c', '#include "types.h"\nA value; A good; Handle handle; Cycle1 cycle;')
        other = self.source('second.c', '#include "types.h"\nB value;')
        paths = [header, source, other]
        catalog = build_bitfield_catalog(paths)
        self.assertEqual(catalog.lookup_variable(source, 'value').status, 'AMBIGUOUS')
        self.assertEqual(catalog.lookup(source, 'Handle').status, 'AMBIGUOUS')
        self.assertEqual(catalog.lookup_variable(source, 'cycle').status, 'UNRESOLVED')
        self.assertEqual(self.compare(paths, 'first', [
            ('value.status', '0~0'), ('handle->status', '0~0'),
            ('cycle.status', '0~0'), ('good.status', '0~3'),
        ]), [
            ('value.status', 'REVIEW', '', False), ('handle->status', 'REVIEW', '', False),
            ('cycle.status', 'REVIEW', '', False), ('good.status', 'MATCH', '0~3', False),
        ])
