import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from rangevalidator.analysis import analyze, match_expression

ROOT = Path(__file__).resolve().parents[1]


class V1FixtureTests(unittest.TestCase):
    def test_original_range_fixture(self):
        p = ROOT / 'tests/fixtures/synthetic_status.h'
        m = analyze(p.read_text(encoding='utf-8-sig'), str(p))
        self.assertFalse(m.reviews)
        self.assertEqual([(f.member, f.width, f.minimum, f.maximum) for f in m.fields], [
            ('active',1,0,1), ('mode',2,0,3), ('state',3,0,7),
            ('temperature',8,0,255), ('voltage',12,0,4095), ('current',16,0,65535),
            ('active',3,0,7), ('error',4,0,15), ('offset',8,-128,127)])
        self.assertTrue(all(f.file == str(p) and f.line > 0 for f in m.fields))

    def test_original_matching_fixture(self):
        p = ROOT / 'tests/fixtures/synthetic_matching.h'
        m = analyze(p.read_text(encoding='utf-8-sig'), str(p))
        self.assertFalse(m.reviews)
        self.assertEqual(len(m.variables), 4)
        for expr, maximum in [('device_p->active',1), ('sensor_p->active',7), ('sensor_p->offset',127), ('device.mode',3), ('sensor_tag_p->error',15)]:
            with self.subTest(expr=expr):
                result = match_expression(m, expr)
                self.assertEqual(result.status, 'TYPE_MATCH')
                self.assertEqual(result.field.maximum, maximum)


class CliTests(unittest.TestCase):
    def test_analyze_and_error_json(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / '합성 공백.h'
            p.write_text('typedef struct { unsigned int flag:1; } A;', encoding='utf-8')
            run = subprocess.run([sys.executable, '-m', 'rangevalidator', 'analyze', str(p)], cwd=ROOT, capture_output=True, encoding='utf-8')
            self.assertEqual(run.returncode, 0, run.stderr)
            self.assertEqual(json.loads(run.stdout)['fields'][0]['maximum'], 1)
            run = subprocess.run([sys.executable, '-m', 'rangevalidator', 'analyze', str(p.with_name('missing.h'))], cwd=ROOT, capture_output=True, encoding='utf-8')
            self.assertEqual(run.returncode, 2)
            self.assertEqual(json.loads(run.stderr)['error']['stage'], 'source.read')
