from collections import Counter
import json
from pathlib import Path
import tempfile
import unittest

from rangevalidator.ch_diagnostics import CHDiagnostics
from rangevalidator.engine import catalog_headers, compare_book
from rangevalidator.readers import read_workbook
from tools.diagnose_ch1 import audit_clang_truth, fixtures
from tools.clang1_synthetic import generate


class CHDiagnosticsTests(unittest.TestCase):
    def test_real_reader_observations_and_unchanged_comparison(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'probe'
            workbook = fixtures(root)
            paths = sorted(root.glob('*.[ch]'))
            book = read_workbook(workbook, 'openpyxl')
            before, coverage = compare_book(None, catalog_headers(paths), book)
            diagnostic = CHDiagnostics(paths, build_options={root / 'probe.c': ['-DSELECT_MODE=1']})
            observed = diagnostic.workbook(book)
            self.assertEqual((before, coverage), compare_book(None, diagnostic.catalog, book))
            self.assertEqual(observed['production_counts'], Counter(r.status for r in before))
            rows = {(row['function'], row['name']): row for row in observed['rows']}
            self.assertEqual(len(rows), 34)
            for fn, expr, bounds in [
                ('global', 'can_data1.state', [0, 7]), ('global', 'can_data_p->state', [0, 7]),
                ('global', 'local_static.active', [0, 1]), ('global', 'can_data_array[0].state', [0, 7]),
                ('global', 'control.mode', [0, 7]), ('global', 'same_control.enable', [0, 1]),
                ('global', 'later.active', [0, 3]), ('global', 'device.state', [0, 3]),
                ('func', 'device.state', [0, 31]), ('update', 'param_p->state', [0, 7]),
                ('global', 'param_p->state', [0, 3]), ('update', 'local_data.state', [0, 7]),
                ('other', 'local_data.state', [0, 31]),
            ]:
                with self.subTest(function=fn, expression=expr):
                    self.assertEqual(rows[fn, expr]['trace']['range'], bounds)
                    self.assertEqual(rows[fn, expr]['comparison_status'], 'MATCH')
            self.assertEqual(rows['global', 'can_data1.state']['name_cell'], 'A2')
            self.assertEqual(rows['func', 'device.state']['range_cell'], 'E2')
            expected = {
                'wrapped.status': 'NOT_BITFIELD',
                'wrapped.status.mode': 'MEMBER_PATH_UNRESOLVED',
                'wrapped_p->status.fault': 'MEMBER_PATH_UNRESOLVED',
                'wrapped_array[0].status.mode': 'MEMBER_PATH_UNRESOLVED',
                'never->active': 'STRUCT_NOT_FOUND', 'lost.state': 'TYPEDEF_UNRESOLVED',
                'unknown.state': 'VARIABLE_TYPE_UNRESOLVED',
                'unknown_p->state': 'POINTER_TYPE_UNRESOLVED',
                'unknown_a[0].state': 'ARRAY_ELEMENT_TYPE_UNRESOLVED',
                'can_chg_p->active': 'TYPE_MATCH',
                'widths.unknown_width': 'BITFIELD_WIDTH_UNRESOLVED',
                'can_data1.no_such_member': 'MEMBER_NOT_FOUND',
                'widths.literal': 'EXCEL_RANGE_REVIEW',
                'conditional_width.mode': 'PREPROCESSOR_UNRESOLVED',
                'optional.active': 'TYPE_MATCH',
            }
            for expr, code in expected.items():
                self.assertEqual(rows['global', expr]['code'], code, expr)
            self.assertEqual(rows['global', 'wrapped.status.mode']['comparison_status'], 'REVIEW')
            self.assertIsNone(rows['global', 'can_data1.no_such_member']['trace']['bitfield'])
            self.assertFalse(rows['global', 'wrapped.status']['trace']['bitfield'])
            self.assertEqual(rows['global', 'widths.literal']['trace']['range'], [0, 15])
            self.assertEqual(diagnostic.trace_expression('control.mode', root / 'alternate.c', 'global')['range'],
                             [0, 127])
            self.assertEqual(diagnostic.trace_expression('can_chg_p->active', root / 'undefined.c', 'global')['code'],
                             'TYPE_MATCH')
            self.assertEqual(diagnostic.trace_expression('local_static.active', None, 'update')['range'], [0, 1])
            self.assertEqual(sum(r.get('duplicate_tag', False) for r in diagnostic.records), 2)

    def test_include_paths_order_and_option_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sources = {
                'sdk/types.h': 'typedef struct { unsigned int state:2; } Remote;',
                'one.h': '#define WIDTH 2', 'two.h': '#define WIDTH 5',
                'option.c': '#include "types.h"\nextern Remote remote;',
                'first.c': '#include "one.h"\n#include "two.h"\n'
                           'struct A { unsigned int state:WIDTH; } v;',
                'second.c': '#include "two.h"\n#include "one.h"\n'
                            'struct B { unsigned int state:WIDTH; } v;',
                'branch.c': '#ifdef SELECT\nstruct C { unsigned int state:2; } v;\n#endif',
                'broken.c': 'struct Good { unsigned int state:2; } good;\n'
                            'struct Broken { unsigned int lost:1;',
            }
            for name, content in sources.items():
                path = root / name
                path.parent.mkdir(exist_ok=True)
                path.write_text(content + '\n', encoding='ascii')
            paths = [*(root / file for file in sources), root / 'missing.h']
            diagnostic = CHDiagnostics(paths, build_options={root / 'option.c': ['-I', 'sdk']})
            files = {Path(f['file']).name: f for f in diagnostic.files}
            self.assertFalse(files['missing.h']['read'])
            self.assertEqual(files['option.c']['unmatched_include_observations'][0]['line'], 1)
            self.assertEqual(files['option.c']['build_options'], 'BUILD_OPTIONS_PROVIDED_NOT_APPLIED')
            for file in ('first.c', 'second.c'):
                self.assertEqual(diagnostic.trace_expression('v.state', root / file, 'global')['code'],
                                 'BITFIELD_WIDTH_UNRESOLVED')
            branch = diagnostic.trace_expression('v.state', root / 'branch.c', 'global')
            self.assertEqual((branch['code'], branch['build_options']),
                             ('TYPE_MATCH', None))
            self.assertEqual(branch['range'], [0, 3])
            self.assertNotIn('CONDITIONAL', branch['production_match']['reason'])
            broken = diagnostic.trace_expression('good.lost', root / 'broken.c', 'global')
            self.assertEqual((broken['code'], broken['bitfield']), ('PARSE_INCOMPLETE', None))
            self.assertTrue(any(Path(d['file']).name == 'broken.c' and d['line'] == 2
                                for d in diagnostic.failures))

    def test_stress_truth_attributes_all_49_unresolved_widths(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp).resolve()
            generate(root, 'stress100')
            jobs = json.loads((root / 'truth.json').read_text(encoding='utf-8'))
            paths = list(root.rglob('*.c')) + list(root.rglob('*.h'))
            diagnostic = CHDiagnostics(paths, build_options={root / j['source']: j['args'] for j in jobs})
            audit = audit_clang_truth(diagnostic, root, jobs)
            self.assertEqual(audit['counts']['width_unresolved'], 49)
            self.assertEqual(audit['counts']['variable_type_correct'], 197)
            self.assertNotIn('variable_type_wrong', audit['counts'])
            for job in jobs:
                trace = diagnostic.trace_expression('local.selected', root / job['source'],
                                                    job['accesses'][1]['function'])
                self.assertEqual((trace['code'], trace['build_options']),
                                 ('PREPROCESSOR_UNRESOLVED', 'BUILD_OPTIONS_PROVIDED_NOT_APPLIED'))
