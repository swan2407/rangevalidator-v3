"""Run the isolated Clang-1 experiment and write per-suite JSON receipts."""
import argparse
from collections import Counter
import json
from pathlib import Path
import platform
import subprocess
import sys
from time import perf_counter
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tools.clang1_synthetic import generate
from rangevalidator import clang_probe
from rangevalidator.bitfield_catalog import build_bitfield_catalog
from rangevalidator.engine import MEMBER_ACCESS, _catalog_match


SUITES = ('functional', 'stress10', 'stress100')


def save(path, result):
    path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def location(row):
    return [row['file'], row['line']] if row else None


def score_clang(job, result):
    issues = []
    records = {r['id']: r for r in result['records']}
    definitions = [r for r in records.values() if r['definition']]

    def record_location(identity):
        return location(records.get(identity, {}).get('definition'))

    def check(label, expected, actual):
        if expected != actual:
            issues.append(dict(check=label, expected=expected, actual=actual))

    status = result['status']
    if job['status'] == 'PARSE_INCOMPLETE_OR_FAILED':
        check('status', True, status in ('PARSE_INCOMPLETE', 'PARSE_FAILED'))
    else:
        check('status', job['status'], status)
    if job.get('reason'):
        check('incomplete reason', True, job['reason'] in result['incomplete_reasons'])
    if not job.get('check_inventory', True):
        if job.get('marker'):
            evidence = str(result['diagnostics']) + str(result['unsupported_options'])
            check('diagnostic marker', True, job['marker'] in evidence)
        if job.get('diagnostic_location'):
            check('diagnostic location', True, any(
                d['severity'] >= 3 and location(d.get('location')) == job['diagnostic_location']
                for d in result['diagnostics']))
        if job['name'] == 'missing_header':
            check('missing header', ['missing_rv.h'], result['missing_headers'])
        if job['name'] == 'missing_macro':
            check('unresolved identifier', ['WIDTH_MISSING'], result['unresolved_identifiers'])
        if job['status'] != 'PARSE_WARNING':
            check('trusted recovery', False, result['trusted'])
            check('range emitted during recovery', False,
                  any(f['range'] is not None for r in definitions for f in r['fields']))
        return issues

    check('definition count', len(job['records']), len(definitions))
    check('forward-only count', job.get('unresolved_records', 0), len(records) - len(definitions))
    for expected in job['records']:
        at = expected['location']
        candidates = [r for r in definitions if location(r['definition']) == at and r['tag'] == expected['tag']]
        if len(candidates) != 1:
            check(f'definition at {at}', 1, len(candidates))
            continue
        row = candidates[0]
        check(f'aliases at {at}', sorted(expected['aliases']), sorted(set(row['aliases'])))
        check(f'forward declarations at {at}', expected['forwards'],
              len({d['offset'] for d in row['declarations'] if d != row['definition']}))
        check(f'parent at {at}', expected['parent'], record_location(row['parent_record_id']))
        fields = {f['name']: [f['width'], f['signed'], f['range']] for f in row['fields']}
        check(f'fields at {at}', expected['fields'], fields)
        for f in row['fields']:
            if f['name'] in expected['fields']:
                check(f'bitfield flag at {at} {f["name"]}', expected['fields'][f['name']][0] is not None,
                      f['bitfield'])

    check('variable count', len(job['variables']), len(result['variables']))
    for expected in job['variables']:
        at, name = expected['location'], expected['name']
        rows = [v for v in result['variables'] if v['name'] == name and location(v['location']) == at]
        if len(rows) != 1:
            check(f'variable {name} at {at}', 1, len(rows))
            continue
        row = rows[0]
        typ = row['type']
        for key, actual in dict(record=record_location(typ['record_id']), kind=row['kind'],
                                function=(row['function'] or {}).get('name'), pointers=typ['pointer_depth'],
                                arrays=typ['array_dimensions'], storage=row['storage']).items():
            check(f'{name} at {at} {key}', expected[key], actual)
        check(f'{name} at {at} trusted link', expected['record'] is not None, typ['record_link_trusted'])
        if expected['qualifiers']:
            check(f'{name} qualifiers', expected['qualifiers'], [layer['qualifiers'] for layer in typ['layers']])
        if expected['spelling']:
            check(f'{name} spelling', expected['spelling'], typ['original'])

    for expected in job['accesses']:
        at, expr = expected['location'], expected['expression']
        rows = [a for a in result['member_accesses'] if location(a['location']) == at and a['expression'] == expr]
        if len(rows) != 1:
            check(f'access {expr} at {at}', 1, len(rows))
            continue
        row = rows[0]
        check(f'{expr} declaration at {at}', expected['variable'], location(row['variable_location']))
        check(f'{expr} owner at {at}', expected['member_record'], record_location(row['member_record_id']))
        check(f'{expr} range at {at}', expected['range'], row['range'])
        check(f'{expr} function at {at}', expected['function'], (row['function'] or {}).get('name'))
        check(f'{expr} trusted access at {at}', True, row['trusted'])
    return issues


def legacy_results(root, paths, jobs):
    catalog = build_bitfield_catalog(paths)

    def origin(at):
        return [Path(at.file).relative_to(root).as_posix(), at.line] if at else None

    def bounds(field):
        return [field.minimum, field.maximum] if field.minimum is not None else None

    reports = []
    for job in jobs:
        if not job.get('check_inventory', True):
            continue
        source = root / job['source']
        counts = Counter()
        differences = []

        def check(metric, expected, actual, context):
            outcome = 'correct' if expected == actual else 'unresolved' if actual is None else 'wrong'
            counts[f'{metric}_{outcome}'] += 1
            if outcome != 'correct':
                differences.append(dict(metric=metric, context=context, expected=expected, actual=actual))

        visible = catalog.source_scopes.get(str(source), {})
        for expected in job['records']:
            at = expected['location']
            candidates = [r for r in catalog.definitions.values()
                          if r.origin.file in visible and origin(r.origin) == at and r.tag == expected['tag']]
            row = candidates[0] if len(candidates) == 1 else None
            check('record', at, origin(row.origin) if row else None, at)
            for alias in expected['aliases']:
                link = catalog.lookup(source, alias)
                check('alias', at, origin(link.definition) if link.status == 'STRUCT' else None, alias)
            for name, (width, signed, interval) in expected['fields'].items():
                fields = [f.declaration for f in row.fields if f.declaration.member == name] if row else []
                field = fields[0] if len(fields) == 1 else None
                if row:
                    check('bitfield_flag', width is not None,
                          bool(fields) if len(fields) < 2 else None, [at, name])
                if interval is not None:
                    check('width', width, (field.width or None) if field else None, [at, name])
                    check('range', interval, bounds(field) if field else None, [at, name])
                    check('signedness', signed, field.signed if field else None, [at, name])

        def binding(name, fn, scope=source):
            function = catalog.lookup_function(scope, fn)
            result = function.variables.get(name) if fn else catalog.lookup_variable(scope, name)
            return function, result

        for expected in job['variables']:
            at, name = expected['location'], expected['name']
            _, row = binding(name, expected['function'], root / at[0])
            actual = None
            if row and row.status == 'STRUCT' and any(origin(d.origin) == at for d in row.declarations):
                actual = origin(row.type_link.definition)
            metric = 'variable_link' if expected['record'] is not None else 'incomplete_variable'
            check(metric, expected['record'], actual, [at, name])

        for expected in job['accesses']:
            expr = expected['expression']
            access = MEMBER_ACCESS.fullmatch(expr)
            if access:
                function, row = binding(access['base'], expected['function'])
                row = row or catalog.lookup_variable(source, access['base'])
                status, _, field = _catalog_match(SimpleNamespace(bitfields=catalog), access, str(source), function)
                if status == 'TYPE_MATCH':
                    definition = origin(row.type_link.definition)
                    check('access_owner', expected['member_record'], definition, expected)
                    declarations = [origin(d.origin) for d in row.declarations]
                    actual = expected['variable'] if expected['variable'] in declarations else declarations[0]
                    check('access_variable', expected['variable'], actual, expected)
                    check('access_range', expected['range'], bounds(field), expected)
                    continue
                if status == 'SKIP' and expected['range'] is None:
                    counts['access_excluded'] += 1
                    continue
            counts['access_unresolved'] += 1
            differences.append(dict(metric='access_unresolved', context=expected))
        reports.append(dict(job=job['name'], counts=dict(counts), differences=differences))

    inventory = Counter(records=len(catalog.definitions),
                        bitfields=sum(len(r.fields) for r in catalog.definitions.values()),
                        aliases=sum(len(r.aliases) for r in catalog.definitions.values()))
    globals_seen = {(d.origin, d.name) for rows in catalog.variables.values() for d in rows}
    scoped = {(d.origin, d.name, d.kind) for scopes in catalog.functions.values()
              for scope in scopes for b in scope.variables.values() for d in b.declarations
              if (d.origin, d.name) not in globals_seen}
    inventory.update(global_variables=len(globals_seen),
                     local_variables=sum(kind != 'PARAMETER' for _, _, kind in scoped),
                     parameters=sum(kind == 'PARAMETER' for _, _, kind in scoped))
    return dict(status='MEASURED', inventory=dict(inventory), jobs=reports,
                diagnostics=[dict(code=d.code, location=[Path(d.file).relative_to(root).as_posix(), d.line],
                                  reason=d.reason) for d in catalog.diagnostics])


def peak_working_set():
    if sys.platform != 'win32':
        return None
    import ctypes
    from ctypes import wintypes

    class MemoryCounters(ctypes.Structure):
        _fields_ = [('size', wintypes.DWORD), ('faults', wintypes.DWORD)] + [
            (name, ctypes.c_size_t) for name in
            ('peak', 'current', 'peak_paged', 'paged', 'peak_nonpaged', 'nonpaged', 'pagefile', 'peak_pagefile')]

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetCurrentProcess.restype = wintypes.HANDLE
    query = kernel.K32GetProcessMemoryInfo
    query.argtypes = [wintypes.HANDLE, ctypes.POINTER(MemoryCounters), wintypes.DWORD]
    query.restype = wintypes.BOOL
    counters = MemoryCounters()
    counters.size = ctypes.sizeof(counters)
    if not query(kernel.GetCurrentProcess(), ctypes.byref(counters), counters.size):
        raise ctypes.WinError(ctypes.get_last_error())
    return counters.peak


def worker(root, engine, dll):
    jobs = json.loads((root / 'truth.json').read_text(encoding='utf-8'))
    paths = sorted(p for p in root.rglob('*') if p.suffix in ('.c', '.h'))
    started = perf_counter()
    if engine == 'legacy':
        report = legacy_results(root, paths, jobs)
        report['totals'] = dict(sum((Counter(j['counts']) for j in report['jobs']), Counter()))
    else:
        try:
            library = clang_probe.library_info(dll)
        except ModuleNotFoundError as exc:
            if exc.name != 'clang':
                raise
            report = dict(status='DEPENDENCY_UNAVAILABLE', error=str(exc), jobs=[])
        except (ImportError, OSError, RuntimeError) as exc:
            report = dict(status='DEPENDENCY_LOAD_FAILED', error=str(exc), jobs=[])
        else:
            db = clang_probe.compilation_jobs(root / 'compile_commands.json')
            rows = []
            for job in jobs:
                source = root / job['source']
                args = db[source] if job['db'] else [f'-working-directory={root}', *job['args']]
                result = clang_probe.analyze(source, root=root, args=args, dll=dll,
                                             config_confirmed=job['confirmed'])
                rows.append(dict(job=job['name'], issues=score_clang(job, result), result=result))
            report = dict(status='DIFFERENCES' if any(r['issues'] for r in rows) else 'MEASURED',
                          library=library, jobs=rows, totals=clang_totals(rows))
    if not report['status'].startswith('DEPENDENCY_'):
        report.update(seconds=round(perf_counter() - started, 6), peak_working_set_bytes=peak_working_set())
    report.update(files=len(paths), python=sys.version, architecture=platform.machine())
    save(root / f'{engine}.json', report)
    return report['status']


def clang_totals(rows):
    """Count observed entities by source location, and links per translation unit."""
    records, aliases, fields, variables = set(), set(), set(), set()
    totals = Counter()
    for row in rows:
        result = row['result']
        totals.update(diagnostic_errors=sum(d['severity'] >= 3 for d in result['diagnostics']),
                      incomplete_units=int(result['status'] in ('PARSE_INCOMPLETE', 'PARSE_FAILED')))
        if not result['trusted']:
            continue
        for record in result['records']:
            if record['definition']:
                records.add(record['id'])
                fields.update((f['location']['file'], f['location']['offset'])
                              for f in record['fields'] if f['bitfield'] and f['location'])
        for alias in result['typedefs']:
            if alias['type']['record_link_trusted']:
                aliases.add((alias['location']['file'], alias['location']['offset']))
        for var in result['variables']:
            variables.add((var['kind'], var['location']['file'], var['location']['offset']))
            key = 'variable_links_trusted' if var['type']['record_link_trusted'] else 'variable_links_unresolved'
            totals[key] += 1
    totals.update(records=len(records), aliases=len(aliases), bitfields=len(fields),
                  global_variables=sum(kind == 'GLOBAL' for kind, _, _ in variables),
                  local_variables=sum(kind == 'LOCAL' for kind, _, _ in variables),
                  parameters=sum(kind == 'PARAMETER' for kind, _, _ in variables),
                  truth_differences=sum(len(row['issues']) for row in rows))
    return dict(totals)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path, help='new output directory')
    parser.add_argument('--suite', choices=SUITES, action='append')
    parser.add_argument('--dll', type=Path)
    parser.add_argument('--worker', choices=('clang', 'legacy'), help=argparse.SUPPRESS)
    args = parser.parse_args()
    root = args.output.resolve()
    if args.worker:
        worker(root, args.worker, args.dll)
        return 0
    root.mkdir(parents=True)
    summary = {}
    for suite in args.suite or SUITES:
        folder = root / suite
        folder.mkdir()
        generate(folder, suite)
        summary[suite] = {}
        for engine in ('legacy', 'clang'):
            command = [sys.executable, '-B', __file__, str(folder), '--worker', engine]
            if args.dll:
                command.extend(('--dll', str(args.dll.resolve())))
            subprocess.run(command, check=True)
            report = json.loads((folder / f'{engine}.json').read_text(encoding='utf-8'))
            summary[suite][engine] = {k: v for k, v in report.items() if k not in ('jobs', 'diagnostics')}
    save(root / 'summary.json', summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return int(any(row['clang']['status'] != 'MEASURED' for row in summary.values()))


if __name__ == '__main__':
    raise SystemExit(main())
