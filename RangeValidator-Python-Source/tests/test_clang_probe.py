import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest

from rangevalidator import clang_probe
from tools.clang1_synthetic import generate
from tools.compare_clang_parser import score_clang


class ClangProbeTests(unittest.TestCase):
    def test_compile_database_preserves_build_flags_and_rejects_ambiguous_variants(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            generate(root, 'functional')
            database = root / 'compile_commands.json'
            jobs = clang_probe.compilation_jobs(database)
            self.assertIn('-Dcinst=', jobs[root / 'app/main.c'])
            args = jobs[root / 'options/config.c']
            self.assertEqual(args[0], f'-working-directory={root}')
            self.assertEqual(args[args.index('-I') + 1], 'include')
            self.assertEqual(args[args.index('-isystem') + 1], 'sdk')
            self.assertNotIn('ignore.obj', args)
            entries = json.loads(database.read_text(encoding='utf-8'))
            database.write_text(json.dumps([*entries, entries[0]]), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'Multiple compile variants'):
                clang_probe.compilation_jobs(database)

    def test_unsupported_options_stop_before_library_loading(self):
        root = Path.cwd()
        result = clang_probe.analyze(root / 'example.c', root=root,
                                     args=['-I', '-DOTHER', '--rv-vendor-mode', '@vendor.rsp'])
        self.assertEqual(result['status'], 'PARSE_INCOMPLETE')
        self.assertFalse(result['trusted'])
        self.assertFalse(result['ast_created'])
        self.assertEqual(result['unsupported_options'], ['-I', '-DOTHER', '--rv-vendor-mode', '@vendor.rsp'])

    def test_stress_truth_covers_requested_sizes(self):
        for count in (10, 100):
            with self.subTest(count=count), tempfile.TemporaryDirectory() as folder:
                root = Path(folder)
                jobs = generate(root, f'stress{count}')
                self.assertEqual(len(list(root.rglob('*.c'))) + len(list(root.rglob('*.h'))), count)
                self.assertGreater(max(map(len, (root / 'unit_00.h').read_text().splitlines())), 2000)
                self.assertGreater(len((root / 'unit_00.c').read_text().splitlines()), 2500)
                self.assertEqual(len(jobs), (count - 2) // 2)

    @unittest.skipUnless(os.getenv('RV_LIBCLANG_DLL') or importlib.util.find_spec('clang'),
                         'Clang dependency unavailable: clang.cindex is absent')
    def test_functional_ast_matches_independent_truth(self):
        dll = os.getenv('RV_LIBCLANG_DLL')
        clang_probe.library_info(dll)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            jobs = generate(root, 'functional')
            db = clang_probe.compilation_jobs(root / 'compile_commands.json')
            for job in jobs:
                with self.subTest(job=job['name']):
                    source = root / job['source']
                    args = db[source] if job['db'] else [f'-working-directory={root}', *job['args']]
                    result = clang_probe.analyze(source, root=root, args=args, dll=dll,
                                                 config_confirmed=job['confirmed'])
                    self.assertEqual(score_clang(job, result), [])
