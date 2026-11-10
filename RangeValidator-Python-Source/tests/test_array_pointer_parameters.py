"""Function linkage must preserve pointer/array order without blocking valid rows."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from test_function_sheets import sheet


class ArrayPointerParameterTests(unittest.TestCase):
    def catalog(self, prototype, definition, extra=''):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        header, source = root / 'api.h', root / 'api.c'
        header.write_text('typedef unsigned char u1;\ntypedef unsigned short u2;\n'
                          'typedef struct { unsigned int flag:1; } Item;\n'
                          + extra + '\nextern Item global;\nvoid Probe(' + prototype + ');\n', encoding='utf-8')
        source.write_text('#include "api.h"\nvoid Probe(' + definition + ') {}\n', encoding='utf-8')
        catalog = catalog_headers([header, source])
        return catalog, catalog.bitfields.lookup_function(str(source), 'Probe'), root

    def test_equivalent_parameter_shapes(self):
        cases = [
            ('u1 (*p)[16]', 'u1 (*data)[16]', ''),
            ('u1 (*p)[SIZE]', 'u1 (*data)[16]', '#define SIZE (8 * 2)\n'),
            ('u1 (*p)[2][16]', 'u1 (*data)[2][16]', ''),
            ('const u1 (*p)[16]', 'const u1 (*data)[16]', ''),
            ('volatile u1 (*p)[16]', 'volatile u1 (*data)[16]', ''),
            ('u1 (*const p)[16]', 'u1 (*data)[16]', ''),
            ('Item (*p)[16]', 'Item (*data)[16]', ''),
            ('u1 (*)[16]', 'u1 (*data)[16]', ''),
            ('u1 p[][16]', 'u1 (*data)[16]', ''),
            ('u1 p[2][16]', 'u1 data[999][16]', ''),
            ('BlockPtr p', 'u1 (*data)[16]', 'typedef u1 (*BlockPtr)[16];\n'),
            ('Alias p', 'u1 (*data)[16]', 'typedef u1 (*BlockPtr)[16];\ntypedef BlockPtr Alias;\n'),
            ('Row *p', 'u1 (*data)[16]', 'typedef u1 Row[16];\n'),
            ('const BlockPtr p', 'u1 (*data)[16]', 'typedef u1 (*BlockPtr)[16];\n'),
            ('const Row *p', 'const u1 (*data)[16]', 'typedef u1 Row[16];\n'),
            ('u1 *p', 'u1 data[16]', ''),
            ('u1 **p', 'u1 *data[16]', ''),
            ('u1 (*p)[16], u2 event, u1 entry', 'u1 (*data)[16], u2 id, u1 slot', ''),
        ]
        for prototype, definition, extra in cases:
            with self.subTest(prototype=prototype, definition=definition):
                _, scope, _ = self.catalog(prototype, definition, extra)
                self.assertEqual(scope.status, 'RESOLVED')

    def test_conflicts_and_unknown_types_stay_unresolved(self):
        cases = [
            ('u1 (*p)[16]', 'u2 (*p)[16]', ''),
            ('u1 (*p)[16]', 'u1 (*p)[32]', ''),
            ('u1 p[][16]', 'u1 p[][32]', ''),
            ('u1 (*p)[2][16]', 'u1 (*p)[3][16]', ''),
            ('u1 (*p)[16]', 'u1 **p', ''),
            ('u1 *p[16]', 'u1 (*p)[16]', ''),
            ('const u1 (*p)[16]', 'u1 (*p)[16]', ''),
            ('u1 (*p)[16]', 'u1 (*p)[2][16]', ''),
            ('u1 (*p)[UNKNOWN]', 'u1 (*p)[UNKNOWN]', ''),
            ('Missing (*p)[16]', 'Missing (*p)[16]', ''),
            ('Alias p', 'Alias p', 'typedef Alias Alias;\n'),
            ('Choice *p', 'Choice *p', '#ifdef FEATURE\n'
             'typedef struct { unsigned int flag:1; } Choice;\n#else\n'
             'typedef struct { unsigned char flag; } Choice;\n#endif\n'),
            ('void (*cb)(u1)', 'void (*cb)(u1)', ''),
        ]
        for prototype, definition, extra in cases:
            with self.subTest(prototype=prototype, definition=definition):
                _, scope, _ = self.catalog(prototype, definition, extra)
                self.assertEqual(scope.status, 'AMBIGUOUS')

    def test_valid_array_parameter_unblocks_global_bitfield(self):
        catalog, _, root = self.catalog('u1 (*p)[16], u2 id, u1 entry',
                                       'u1 (*data)[16], u2 id, u1 entry')
        book = Workbook(str(root / 'api.xlsx'), 'openpyxl', [sheet('1.Probe', [
            ('Name', 'Type', 'Range'), ('global.flag', 'unsigned int', '0~7')])])
        rows, _ = compare_book(None, catalog, book)
        self.assertEqual([(r.status, r.expected, r.candidate) for r in rows], [('MISMATCH', '0~1', True)])

    def test_array_pointer_binding_keeps_required_array_access(self):
        for prototype, definition, extra in [
            ('Item (*p)[16]', 'Item (*p)[16]', ''),
            ('Handle p', 'Handle p', 'typedef Item (*Handle)[16];\n'),
        ]:
            with self.subTest(prototype=prototype):
                _, scope, _ = self.catalog(prototype, definition, extra)
                binding = scope.variables['p']
                self.assertEqual(binding.status, 'STRUCT')
                self.assertEqual(binding.declarations[0].arrays, ('', '16'))
                self.assertEqual(binding.type_link.pointer_depth, 0)


if __name__ == '__main__':
    unittest.main()
