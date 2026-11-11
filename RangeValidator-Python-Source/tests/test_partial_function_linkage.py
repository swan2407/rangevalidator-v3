"""Unsupported sibling parameters must not hide independently known members."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook as ExcelWorkbook
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.models import Workbook
from rangevalidator.workflow import run_comparison
from test_function_sheets import sheet


TYPES = '''typedef unsigned char u1;
typedef unsigned int u4;
typedef enum { IDLE, RUNNING } SystemState;
typedef enum { PHASE_IDLE, PHASE_ACTIVE } RecoveryPhase;
typedef struct RecoveryContext RficRecoveryContext_t;
typedef void (*StateFuncPtr_t)(RficRecoveryContext_t*);
struct RecoveryContext {
    SystemState currentSystemState;
    StateFuncPtr_t currentStateFuncPtr;
    RecoveryPhase currentRecoveryPhase;
    u4 tickCount;
    u1 firstEnterFlag;
    unsigned int testBit : 3;
};
typedef struct { unsigned int testBit:1; } Other;
'''
PARAMETERS = 'RficRecoveryContext_t* ctx, SystemState newStateEnum, StateFuncPtr_t newFuncPtr'


class PartialFunctionLinkageTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def make(self, parameters=PARAMETERS, prototype=None, extra='', body='', name='StateTransition'):
        header, source = self.root / 'rf_recovery.h', self.root / 'rf_recovery.c'
        header.write_text(TYPES + '\n' + extra, encoding='utf-8')
        proto = parameters if prototype is None else prototype
        source.write_text('#include "rf_recovery.h"\nstatic void ' + name + '(' + proto + ');\n'
                          'static void ' + name + '(' + parameters + ') { ' + body + ' }\n', encoding='utf-8')
        catalog = catalog_headers([header, source])
        return catalog, source

    def rows(self, catalog, *expressions, title='1.StateTransition', book='rf_recovery'):
        workbook = Workbook(str(self.root / (book + '.xlsx')), 'openpyxl', [sheet(title,
            [('Name', 'Type', 'Range'), *((expr, '', '0~255') for expr in expressions)])])
        return compare_book(None, catalog, workbook)[0]

    def test_photo_function_signature_and_five_rows_in_real_workbook(self):
        # The signature/prototype and five expressions are visible in the photos.
        # Their missing struct definition is modeled explicitly as ordinary fields.
        catalog, source = self.make(body='ctx->currentSystemState = newStateEnum; '
            'ctx->currentStateFuncPtr = newFuncPtr; ctx->currentRecoveryPhase = PHASE_IDLE; '
            'ctx->tickCount = 0U; ctx->firstEnterFlag = 0;')
        book = self.root / 'rf_recovery.xlsx'
        w = ExcelWorkbook(); ws = w.active; ws.title = '1.StateTransition'
        for column, value in ((4, 'Name'), (6, 'Type'), (8, 'Range')):
            ws.cell(32, column, value)
        fields = ('currentRecoveryPhase', 'currentStateFuncPtr', 'currentSystemState', 'firstEnterFlag', 'tickCount')
        for row, field in enumerate(fields, 33):
            ws.cell(row, 4, 'ctx->' + field)
            ws.cell(row, 8, '-' if row < 36 else '0~255' if row == 36 else '0~4294967295')
        w.save(book); w.close(); before = book.read_bytes()
        session = run_comparison(None, [source, source.with_suffix('.h')], [book], backend='openpyxl')
        self.assertEqual(session.errors, [])
        self.assertEqual([(r.cell, r.status) for r in session.results], [(f'H{i}', 'SKIP') for i in range(33, 38)])
        self.assertTrue(all('일반 멤버' in r.reason for r in session.results))
        self.assertEqual(book.read_bytes(), before)

    def test_bitfield_survives_unsupported_callback_typedef(self):
        catalog, source = self.make()
        scope = catalog.bitfields.lookup_function(str(source), 'StateTransition')
        self.assertEqual(scope.status, 'PARTIAL')
        self.assertEqual(scope.variables['ctx'].status, 'STRUCT')
        row = self.rows(catalog, 'ctx->testBit')[0]
        self.assertEqual((row.status, row.expected), ('MISMATCH', '0~7'))
        self.assertIn('newFuncPtr', row.reason)

    def test_other_unresolved_parameter_shapes_do_not_block_ctx(self):
        for other, extra in [
            ('void (*cb)(int)', ''),
            ('Missing extra', ''),
            ('u1 (*buffer)[UNKNOWN_SIZE]', ''),
            ('Recursive value', 'typedef Recursive Recursive;'),
            ('void callback(int)', ''),
            ('...', ''),
            ('StateFuncPtr_t const callbacks[3]', ''),
        ]:
            with self.subTest(other=other):
                catalog, _ = self.make('RficRecoveryContext_t *ctx, ' + other, extra=extra)
                row = self.rows(catalog, 'ctx->testBit')[0]
                self.assertEqual((row.status, row.expected), ('MISMATCH', '0~7'))

    def test_parameter_names_may_differ_in_prototype(self):
        catalog, _ = self.make(prototype='RficRecoveryContext_t*, SystemState, StateFuncPtr_t')
        self.assertEqual(self.rows(catalog, 'ctx->testBit')[0].status, 'MISMATCH')

    def test_unknown_target_never_falls_back_to_global(self):
        for parameters, extra, expression in [
            ('RficRecoveryContext_t *ctx, StateFuncPtr_t newFuncPtr', 'RficRecoveryContext_t *newFuncPtr;', 'newFuncPtr->testBit'),
            ('RficRecoveryContext_t *ctx, Missing unknown', 'RficRecoveryContext_t unknown;', 'unknown.testBit'),
            ('RficRecoveryContext_t *ctx, void (*cb)(int)', 'RficRecoveryContext_t *cb;', 'cb->testBit'),
            ('RficRecoveryContext_t *ctx, void callback(int)', 'RficRecoveryContext_t callback;', 'callback.testBit'),
        ]:
            with self.subTest(expression=expression):
                catalog, _ = self.make(parameters, extra=extra)
                good, bad = self.rows(catalog, 'ctx->testBit', expression)
                self.assertEqual(good.status, 'MISMATCH')
                self.assertEqual((bad.status, bad.candidate, bad.expected), ('REVIEW', False, ''))

    def test_unknown_prototype_target_blocks_only_affected_parameter(self):
        catalog, _ = self.make('RficRecoveryContext_t *ctx, Other *other',
                              prototype='Unknown *ctx, Other *other')
        bad, good = self.rows(catalog, 'ctx->testBit', 'other->testBit')
        self.assertEqual(bad.status, 'REVIEW')
        self.assertIn("'ctx'", bad.reason)
        self.assertIn('Unknown * ctx', bad.reason)
        self.assertEqual((good.status, good.expected), ('MISMATCH', '0~1'))

    def test_known_conflict_is_not_hidden_by_unknown_sibling(self):
        for prototype in (
            'Other *ctx, SystemState s, StateFuncPtr_t f',
            'RficRecoveryContext_t ctx, SystemState s, StateFuncPtr_t f',
            'RficRecoveryContext_t *ctx, StateFuncPtr_t f',
            'RficRecoveryContext_t *ctx, SystemState s, ...',
        ):
            with self.subTest(prototype=prototype):
                catalog, source = self.make(prototype=prototype)
                self.assertEqual(catalog.bitfields.lookup_function(str(source), 'StateTransition').status, 'AMBIGUOUS')
                row = self.rows(catalog, 'ctx->testBit')[0]
                self.assertEqual((row.status, row.candidate), ('REVIEW', False))
                self.assertIn('불일치', row.reason)

    def test_local_and_global_members_are_independent_of_unknown_callback(self):
        catalog, _ = self.make(extra='RficRecoveryContext_t global;', body='RficRecoveryContext_t local;')
        for row in self.rows(catalog, 'local.testBit', 'global.testBit'):
            self.assertEqual((row.status, row.expected), ('MISMATCH', '0~7'))

    def test_truncated_sheet_and_different_filename_keep_partial_scope(self):
        name = 'recovery_process_transition_with_callback'
        catalog, _ = self.make(name=name)
        row = self.rows(catalog, 'ctx->testBit', title=('1.' + name)[:31], book='other_design')[0]
        self.assertEqual((row.status, row.expected), ('MISMATCH', '0~7'))

    def test_case_insensitive_target_cannot_escape_its_issue(self):
        catalog, _ = self.make(extra='RficRecoveryContext_t *NEWFUNCPTR;')
        good, bad = self.rows(catalog, 'CTX->testBit', 'NEWFUNCPTR->testBit')
        self.assertEqual(good.status, 'MISMATCH')
        self.assertEqual(bad.status, 'REVIEW')

    def test_xml_and_macro_ranges_keep_function_context(self):
        from rangevalidator.build_options import load_build_options
        from test_build_options import xml
        _, source = self.make(extra='#define LIMIT 7U')
        path = self.root / 'build.xml'; path.write_text(xml(), encoding='utf-8')
        catalog = catalog_headers([source, source.with_suffix('.h')], build_options=load_build_options(path))
        book = Workbook('other_design.xlsx', 'openpyxl', [sheet('1.StateTransition',
            [('Name', 'Type', 'Range'), ('ctx->testBit', '', '0~LIMIT')])])
        row = compare_book(None, catalog, book)[0][0]
        self.assertEqual((row.status, row.expected), ('MATCH', '0~7'))
        self.assertIn('C 문맥 rf_recovery.c', row.reason)

    def test_multiple_static_definitions_are_not_resolved_by_partial_linkage(self):
        _, source = self.make()
        other = self.root / 'other.c'
        other.write_text(source.read_text(encoding='utf-8'), encoding='utf-8')
        catalog = catalog_headers([source, other, source.with_suffix('.h')])
        row = self.rows(catalog, 'ctx->testBit', book='unrelated_design')[0]
        self.assertEqual((row.status, row.candidate), ('REVIEW', False))

    def test_return_type_conflict_still_blocks_partial_parameters(self):
        _, source = self.make()
        code = source.read_text(encoding='utf-8')
        source.write_text(code.replace('static void StateTransition', 'static int StateTransition', 1), encoding='utf-8')
        catalog = catalog_headers([source, source.with_suffix('.h')])
        row = self.rows(catalog, 'ctx->testBit')[0]
        self.assertEqual(row.status, 'REVIEW')
        self.assertIn('반환 타입 불일치', row.reason)


if __name__ == '__main__':
    unittest.main()
