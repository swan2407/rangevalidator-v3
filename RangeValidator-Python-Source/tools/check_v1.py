"""Compile a read-only V1 oracle into THIS project's build directory.
Only named synthetic fixtures and literals from V1 tests; no production discovery.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / 'tests'))
from test_analysis import BASIC, NEGATIVE
from rangevalidator.analysis import analyze, match_expression


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--v1', type=Path, default=Path('C:/Dev/RangeValidator'))
    p.add_argument('--compiler', default='C:/msys64/ucrt64/bin/g++.exe')
    args = p.parse_args()
    references = [args.v1 / s for s in ['src/bitfield_parser.cpp', 'src/type_matcher.cpp',
        'include/bitfield_parser.h', 'include/type_matcher.h', 'tests/parser_tests.cpp',
        'tests/range_tests.cpp', 'tests/matching_tests.cpp', 'samples/synthetic_status.h',
        'samples/synthetic_matching.h']]
    hashes = {str(f): sha(f) for f in references}
    out = ROOT / 'build'
    out.mkdir(exist_ok=True)
    exe = out / 'v1_probe.exe'
    env = dict(os.environ, PATH=str(Path(args.compiler).parent) + os.pathsep + os.environ.get('PATH', ''), TEMP=str(out), TMP=str(out), TMPDIR=str(out))
    run = subprocess.run([args.compiler, '-std=c++17', '-O2', '-static', '-I' + str(args.v1 / 'include'),
            str(ROOT / 'tools/v1_probe.cpp'), str(references[0]), str(references[1]), '-o', str(exe)],
            capture_output=True, text=True, env=env)
    if run.returncode:
        print(run.stderr, file=sys.stderr)
        return run.returncode
    cases = [('basic', BASIC, 'p->active'),
        ('boundary', 'struct E { unsigned int u:32; signed int s:32; signed int one:1; };', 'p->u'),
        ('multiline', 'typedef struct Tag {\nunsigned int flag:1;\nsigned int state:3;\n} A; A *p;', 'p->state'),
        ('tag', 'struct Tag { unsigned int active:3; }; struct Tag *p;', 'p->active'),
        ('comments', '// fake\n' + BASIC + ' /* A *p; */', 'p->active'),
        ('extern', 'typedef struct { unsigned int active:1; } A; extern A *p;', 'p->active'),
        ('include_guard', '#ifndef A_H\n#define A_H\n' + BASIC + '\n#endif', 'p->active'),
        ('wrong_operator', BASIC, 'p.active'),
        ('unknown_base', BASIC, 'other->active'),
        ('ordinary_member', BASIC, 'p->normal')]
    cases += [(f'negative_{i:02}', source, 'p->active') for i, source in enumerate(NEGATIVE)]
    for name in ('synthetic_status.h', 'synthetic_matching.h'):
        cases.append((name, (ROOT / 'tests/fixtures' / name).read_text(encoding='utf-8-sig'), 'device_p->active'))
    records, differences, unexpected = [], [], []
    for name, source, expr in cases:
        run = subprocess.run([str(exe), expr], input=source, text=True, capture_output=True, check=True)
        lines = [s.split('\t') for s in run.stdout.splitlines()]
        m = analyze(source, 'fixture.h')
        pyfields = [['FIELD', f.struct_name, f.struct_tag, f.member, f.declared_type, str(f.width), str(f.line),
                     '?' if f.signed is None else str(int(f.signed)),
                     '?' if f.minimum is None else str(f.minimum),
                     '?' if f.maximum is None else str(f.maximum)] for f in m.fields]
        v1fields = [s for s in lines if s[0] == 'FIELD']
        v1blocked = next(s[1] == '1' for s in lines if s[0] == 'BLOCKED')
        v1status = next(s[1] for s in lines if s[0] == 'MATCH')
        py_status = match_expression(m, expr).status
        record = dict(case=name, fields_equal=v1fields == pyfields,
                      binding_gate_equal=v1blocked == bool(m.reviews), v1_status=v1status, python_status=py_status)
        records.append(record)
        if not record['fields_equal'] or not record['binding_gate_equal'] or v1status != py_status:
            # Parser-only V1 accepts a keyword field; binding model blocks it.
            reason = None
            if name == 'ordinary_member' and (v1status, py_status) == ('REVIEW', 'SKIP'):
                reason = 'Requested ordinary integer exclusion uses SKIP; legacy single-header matcher used REVIEW.'
            elif name == 'negative_32' and v1blocked and m.reviews and v1status == py_status == 'REVIEW':
                reason = 'Python rejects C keyword member during extraction; V1 parser extracts it but binding model blocks it.'
            detail = dict(record, reason=reason, v1_fields=v1fields, python_fields=pyfields)
            differences.append(detail)
            if reason is None:
                unexpected.append(detail)
    preserved = hashes == {str(f): sha(f) for f in references}
    report = dict(cases=len(records), unexpected_differences=len(unexpected), reference_files_unchanged=preserved,
                  reference_sha256=hashes, results=records, differences=differences)
    (ROOT / 'docs/V1_DIFFERENTIAL_RESULTS.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('cases', 'unexpected_differences', 'reference_files_unchanged')}))
    print(json.dumps(differences, ensure_ascii=False, indent=2))
    return int(bool(unexpected) or not preserved)


if __name__ == '__main__':
    raise SystemExit(main())
