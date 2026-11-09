from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.bitfield_catalog import build_bitfield_catalog


class BitfieldCatalogTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def files(self, **sources):
        paths = []
        for name, source in sources.items():
            path = self.root / name.replace('_', '.')
            path.write_text(source, encoding='utf-8')
            paths.append(path)
        return paths

    def test_links_distributed_definitions_aliases_and_integer_types(self):
        types, aliases, implementation = self.files(
            ints_h="""#ifndef INTS_H
#define INTS_H
typedef unsigned char u1;
typedef unsigned short u2;
typedef unsigned int u4;
typedef signed int Signed;
typedef u1 U8;
typedef U8 BYTE_TYPE;
#define FIRST 1+2
#define WIDTH (FIRST * 2)
#endif
""",
            aliases_h="""#include "ints.h"
struct DeviceTag;
typedef struct DeviceTag Device_t;
typedef Device_t DeviceAlias_t;
typedef DeviceAlias_t DeviceStatus_t;
typedef Device_t DeviceAlias_t;
""",
            device_c="""#include "aliases.h"
struct DeviceTag { BYTE_TYPE active:WIDTH; u2 mode:3; u4 count:4; Signed offset:4; };
""")
        paths = [types, aliases, implementation]
        catalog = build_bitfield_catalog(paths)
        self.assertEqual(catalog, build_bitfield_catalog(reversed(paths)))
        link = catalog.lookup(aliases, 'DeviceStatus_t')
        self.assertEqual(link.status, 'STRUCT')
        self.assertEqual(link, catalog.lookup(implementation, 'DeviceStatus_t'))
        self.assertEqual({Path(origin.file) for origin in link.evidence}, {aliases.resolve(), implementation.resolve()})
        definition = catalog.definitions[link.definition]
        self.assertEqual(definition.tag, 'DeviceTag')
        self.assertEqual(definition.aliases, ('DeviceAlias_t', 'DeviceStatus_t', 'Device_t'))
        self.assertEqual(catalog.lookup(implementation, 'struct DeviceTag').definition, link.definition)
        self.assertEqual([(f.declaration.member, f.declaration.width, f.declaration.signed,
                           f.declaration.minimum, f.declaration.maximum, f.status)
                          for f in definition.fields],
                         [('active', 5, False, 0, 31, 'RESOLVED'),
                          ('mode', 3, False, 0, 7, 'RESOLVED'),
                          ('count', 4, False, 0, 15, 'RESOLVED'),
                          ('offset', 4, True, -8, 7, 'RESOLVED')])
        self.assertEqual(definition.fields[0].type_link.integer_type, 'unsigned char')

    def test_anonymous_and_tagged_typedefs_keep_independent_members(self):
        source, = self.files(a_h="""typedef unsigned char u1;
typedef CAN_A_t Early;
typedef struct { u1 status:2; } CAN_A_t, Alias; typedef struct B { u1 status:5; } CAN_B_t;
""")
        catalog = build_bitfield_catalog([source])
        a = catalog.lookup(source, 'Early').definition
        b = catalog.lookup(source, 'CAN_B_t').definition
        self.assertNotEqual(a, b)
        self.assertEqual(a, catalog.lookup(source, 'Alias').definition)
        self.assertEqual(b, catalog.lookup(source, 'struct B').definition)
        self.assertEqual([catalog.definitions[key].fields[0].declaration.maximum for key in (a, b)], [3, 31])

    def test_conflicting_translation_units_stay_distinct(self):
        header, first, second, unlinked = self.files(
            common_h='typedef struct Shared Local;',
            first_c='#include "common.h"\nstruct Shared { unsigned int status:2; };',
            second_c='#include "common.h"\nstruct Shared { signed int status:5; };',
            unlinked_c='typedef struct Shared Local;')
        catalog = build_bitfield_catalog([header, first, second, unlinked])
        self.assertEqual(catalog.lookup(header, 'Local').status, 'AMBIGUOUS')
        self.assertEqual(catalog.lookup(unlinked, 'Local').status, 'AMBIGUOUS')
        self.assertNotEqual(catalog.lookup(first, 'Local').definition,
                            catalog.lookup(second, 'Local').definition)

    def test_conflicting_integer_typedef_in_visible_headers(self):
        first, second, source = self.files(
            one_h='typedef unsigned int u1;',
            two_h='typedef signed int u1;',
            main_c='#include "one.h"\n#include "two.h"\nstruct T { u1 status:2; };')
        catalog = build_bitfield_catalog([first, second, source])
        field = next(iter(catalog.definitions.values())).fields[0]
        self.assertEqual(field.status, 'AMBIGUOUS')
        self.assertIsNone(field.declaration.maximum)

    def test_unresolved_declarations_do_not_block_good_fields(self):
        source, = self.files(a_h="""typedef B A;
typedef A B;
typedef struct Missing Missing;
#define LOOP LOOP
#define FUNCTION(x) x
struct T {
    u1 no_typedef:3;
    char unspecified:3;
    int plain:3;
    A cycle:3;
    unsigned int width:LOOP;
    unsigned int call:FUNCTION(3);
    unsigned int *pointer:3;
    unsigned int valid:3;
};
""")
        catalog = build_bitfield_catalog([source])
        self.assertEqual(catalog.lookup(source, 'A').status, 'UNRESOLVED')
        self.assertEqual(catalog.lookup(source, 'Missing').status, 'INCOMPLETE')
        fields = {f.declaration.member: f for f in next(iter(catalog.definitions.values())).fields}
        self.assertEqual(fields['valid'].declaration.maximum, 7)
        self.assertEqual(fields['plain'].status, 'SIGNEDNESS_UNRESOLVED')
        for name in ('no_typedef', 'unspecified', 'cycle', 'width', 'call', 'pointer'):
            self.assertIsNone(fields[name].declaration.maximum, name)

    def test_conditional_definitions_and_macros_are_uncertain(self):
        source, = self.files(a_h="""#if OPTION
struct T { unsigned int status:2; };
#define WIDTH 2
#else
struct T { unsigned int status:5; };
#define WIDTH 5
#endif
struct U { unsigned int unknown:WIDTH; unsigned int known:1; };
""")
        catalog = build_bitfield_catalog([source])
        self.assertEqual(catalog.lookup(source, 'struct T').status, 'CONDITIONAL')
        definitions = list(catalog.definitions.values())
        self.assertEqual([d.status for d in definitions], ['CONDITIONAL', 'CONDITIONAL', 'RESOLVED'])
        self.assertEqual([f.status for f in definitions[-1].fields], ['WIDTH_UNRESOLVED', 'RESOLVED'])

    def test_common_header_identity_and_local_typedef_shadow(self):
        shared, first, second, both = self.files(
            shared_h='typedef struct { unsigned int value:1; } Shared;',
            first_c='#include "shared.h"\ntypedef Shared Local;',
            second_c='#include "shared.h"\n'
                     'typedef struct { unsigned int value:4; } Local;',
            both_c='#include "shared.h"\n#include "shared.h"')
        catalog = build_bitfield_catalog([shared, first, second, both])
        shared_id = catalog.lookup(shared, 'Shared').definition
        self.assertEqual(catalog.lookup(both, 'Shared').definition, shared_id)
        self.assertEqual(catalog.lookup(first, 'Local').definition, shared_id)
        self.assertNotEqual(catalog.lookup(second, 'Local').definition, shared_id)
        self.assertEqual(len(catalog.definitions), 2)

    def test_conditional_include_retains_condition_and_calculates_declared_width(self):
        header, source = self.files(
            types_h='typedef unsigned int Word;\n#define WIDTH 3',
            main_c='#if FEATURE\n#include "types.h"\n#endif\n'
                   'struct T { Word value:1; unsigned int flags:WIDTH; };')
        catalog = build_bitfield_catalog([header, source])
        self.assertEqual(catalog.lookup(source, 'Word').status, 'CONDITIONAL')
        fields = next(iter(catalog.definitions.values())).fields
        self.assertEqual([f.status for f in fields], ['CONDITIONAL', 'CONDITIONAL'])
        self.assertEqual([f.declaration.maximum for f in fields], [1, 7])

    def test_missing_input_is_recorded(self):
        source, = self.files(a_h='struct T { unsigned int status:2; };')
        catalog = build_bitfield_catalog([source, self.root / 'missing.h'])
        self.assertEqual(len(catalog.definitions), 1)
        self.assertEqual(catalog.diagnostics[0].code, 'SOURCE_READ_ERROR')
