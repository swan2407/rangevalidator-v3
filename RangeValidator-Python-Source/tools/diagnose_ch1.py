"""Generate and diagnose synthetic C/H and Excel inputs. No company files are loaded."""
import argparse
from collections import Counter
import json
from pathlib import Path
import sys

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rangevalidator.ch_diagnostics import CHDiagnostics
from rangevalidator.readers import read_workbook
from tools.clang1_synthetic import generate


def save(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')


def fixtures(root):
    root.mkdir()
    files = {
        'defs.h': '''struct testmode_control_status_s {
    unsigned int enable : 1;
    unsigned int mode : 3;
};
typedef struct testmode_control_status_s testmode_control_status_t;
typedef struct { unsigned int state:3; unsigned int active:1; } CAN_DATA_t;
typedef struct { unsigned int fault:1; unsigned int mode:3; } Inner_t;
typedef struct { Inner_t status; unsigned int valid:1; } Outer_t;
typedef struct { unsigned int state:2; } DeviceA_t;
typedef struct { unsigned int state:5; } DeviceB_t;
typedef CAN_DATA_t can_chg_t;
''',
        'aliases.h': '''#include "defs.h"
typedef struct testmode_control_status_s Control;
typedef Control ControlChain;
typedef struct Later Later_t;
struct Later { unsigned int active:2; };
typedef struct Never Never_t;
typedef UnknownType Lost_t;
void update(CAN_DATA_t *param_p);
''',
        'probe.c': '''#include "aliases.h"
extern CAN_DATA_t can_data1;
extern const CAN_DATA_t *can_data_p;
static CAN_DATA_t local_static;
CAN_DATA_t can_data_array[4];
extern Outer_t wrapped;
extern Outer_t *wrapped_p;
extern Outer_t wrapped_array[4];
extern DeviceA_t device;
extern DeviceA_t *param_p;
extern ControlChain control;
extern testmode_control_status_t same_control;
extern Later_t later;
extern Never_t *never;
extern Lost_t lost;
extern struct NoTag missing_record;
extern Unavailable unknown;
extern Unavailable *unknown_p;
extern Unavailable unknown_a[2];
#define cinst
extern const can_chg_t * cinst can_chg_p;
#define WIDTH 4
struct Widths { unsigned int literal:WIDTH; unsigned int unknown_width:ABSENT; } widths;
#ifndef SELECT_MODE
#define WIDTH_ALT 2
#elif SELECT_MODE == 1
#define WIDTH_ALT 3
#else
#define WIDTH_ALT 5
#endif
struct Condition { unsigned int mode:WIDTH_ALT; } conditional_width;
#ifdef FEATURE
struct Optional { unsigned int active:2; } optional;
#endif
void func(void) { DeviceB_t device; }
void update(CAN_DATA_t *param_p) { CAN_DATA_t local_data; local_data.state = param_p->state; }
void other(void) { DeviceB_t local_data; }
''',
        'undefined.c': '''#include "defs.h"
extern const can_chg_t * cinst can_chg_p;
''',
        'alternate.c': '''struct testmode_control_status_s { unsigned int mode:7; };
extern struct testmode_control_status_s control;
''',
    }
    for name, content in files.items():
        (root / name).write_text(content, encoding='ascii')

    # Independent C semantics. Missing/conditional inputs have no fixed range.
    truth = [
        ('global', 'can_data1.state', [0, 7]), ('global', 'can_data_p->state', [0, 7]),
        ('global', 'local_static.active', [0, 1]), ('global', 'can_data_array[0].state', [0, 7]),
        ('global', 'wrapped.valid', [0, 1]), ('global', 'wrapped.status.fault', [0, 1]),
        ('global', 'wrapped.status.mode', [0, 7]), ('global', 'wrapped_p->status.fault', [0, 1]),
        ('global', 'wrapped_array[0].status.mode', [0, 7]), ('global', 'wrapped.status', None),
        ('global', 'device.state', [0, 3]), ('func', 'device.state', [0, 31]),
        ('update', 'param_p->state', [0, 7]), ('global', 'param_p->state', [0, 3]),
        ('update', 'local_data.state', [0, 7]), ('other', 'local_data.state', [0, 31]),
        ('global', 'control.mode', [0, 7]), ('global', 'same_control.enable', [0, 1]),
        ('global', 'later.active', [0, 3]), ('global', 'never->active', None),
        ('global', 'lost.state', None), ('global', 'missing_record.state', None),
        ('global', 'unknown.state', None), ('global', 'unknown_p->state', None),
        ('global', 'unknown_a[0].state', None), ('global', 'can_chg_p->active', [0, 1]),
        ('global', 'widths.literal', [0, 15]), ('global', 'widths.unknown_width', None),
        ('global', 'conditional_width.mode', None), ('global', 'optional.active', None),
        ('global', 'absent.state', None), ('global', 'can_data1.no_such_member', None),
        ('global', 'can_data1->state', None), ('global', 'can_data_array.state', None),
    ]
    save(root / 'truth.json', [dict(function=fn, expression=expr, range=bounds) for fn, expr, bounds in truth])
    path = root / 'probe_Detail.xlsx'
    workbook = Workbook()
    workbook.remove(workbook.active)
    for n, fn in enumerate(dict.fromkeys(fn for fn, _, _ in truth), 1):
        sheet = workbook.create_sheet(f'{n}. {fn}')
        form_a = fn in ('func', 'other')
        sheet.append(['Input/Output', 'Property', 'Name', 'Data Type', 'Scope'] if form_a else
                     ['Name', 'Type', 'Range'])
        for scope, expr, bounds in truth:
            if scope != fn:
                continue
            interval = '~'.join(map(str, bounds)) if bounds is not None else '0~1'
            if expr == 'widths.literal':
                interval = '=15'
            row = [expr, 'unsigned int', interval]
            sheet.append(['Input', 'Global', *row] if form_a else row)
    workbook.save(path)
    workbook.close()
    return path


def audit_clang_truth(diagnostic, root, jobs):
    """Compare the original independent truth by file and line, scoped per C variant."""
    counts = Counter()
    differences = []
    checks = []
    for job in jobs:
        if not job.get('check_inventory', True):
            continue
        source = str(root / job['source'])
        visible = diagnostic.bitfields.source_scopes[source]
        for expected in job['records']:
            file, line = expected['location']
            candidates = [r for r in diagnostic.records if r['defined'] and r['origin']['file'] in visible
                          and r['origin']['file'] == str(root / file) and r['origin']['line'] == line
                          and r['tag'] == expected['tag']]
            record = candidates[0] if len(candidates) == 1 else None
            checks.append((job['name'], 'structure', [file, line], True, True if record else None))
            for alias in expected['aliases']:
                link = diagnostic.bitfields.lookup(source, alias)
                actual = ([Path(link.definition.file).relative_to(root).as_posix(), link.definition.line]
                          if link.status == 'STRUCT' else None)
                checks.append((job['name'], 'typedef', alias, expected['location'], actual))
            for member, (width, signed, bounds) in expected['fields'].items():
                fields = [f for f in record['bitfields'] if f['member'] == member] if record else []
                key = [file, line, member]
                if width is None:
                    ordinary = record and member in record['ordinary_members'] and not fields
                    checks.append((job['name'], 'ordinary', key, True, True if ordinary else None))
                    continue
                checks.append((job['name'], 'bitfield_extracted', key, True, True if fields else None))
                if width == 0:
                    continue
                field = fields[0] if len(fields) == 1 else {}
                for label, expected_value in (('width', width), ('signed', signed), ('range', bounds)):
                    checks.append((job['name'], label, key, expected_value, field.get(label)))
        for expected in job['variables']:
            file, line = expected['location']
            candidates = [v for v in diagnostic.variables if v['origin']['file'] == str(root / file)
                          and v['origin']['line'] == line and v['name'] == expected['name']
                          and v['scope'] == expected['kind'] and v['function'] == expected['function']]
            link = (candidates[0]['type_link'] or {}) if len(candidates) == 1 else {}
            definition = link.get('definition')
            actual = ([Path(definition['file']).relative_to(root).as_posix(), definition['line']]
                      if link.get('status') == 'STRUCT' and definition else None)
            checks.append((job['name'], 'variable_type' if expected['record'] else 'incomplete_variable',
                           [file, line, expected['name']], expected['record'], actual))
        for expected in job['accesses']:
            trace = diagnostic.trace_expression(expected['expression'], source, expected['function'])
            checks.append((job['name'], 'expression_range', expected['expression'], expected['range'],
                           trace['range'] if trace['code'] == 'TYPE_MATCH' else None))
            if expected['range'] is not None:
                member_origin = None
                if trace['code'] == 'TYPE_MATCH':
                    definition = trace['type_link']['definition']
                    member_origin = [Path(definition['file']).relative_to(root).as_posix(), definition['line']]
                checks.append((job['name'], 'expression_member_owner', expected['expression'],
                               expected['member_record'], member_origin))
    for job, metric, item, expected, actual in checks:
        outcome = 'correct' if actual == expected else 'unresolved' if actual is None else 'wrong'
        counts[f'{metric}_{outcome}'] += 1
        if outcome != 'correct':
            differences.append(dict(job=job, metric=metric, item=item, expected=expected, actual=actual))
    return dict(counts=dict(counts), differences=differences)


def run(root, backend='openpyxl'):
    root = root.resolve()
    root.mkdir(parents=True)
    probe = root / 'probe'
    workbook = fixtures(probe)
    diagnostic = CHDiagnostics(probe.glob('*.[ch]'), build_options={probe / 'probe.c': ['-DSELECT_MODE=1']})
    observation = diagnostic.workbook(read_workbook(workbook, backend))
    report = diagnostic.report([observation])
    rows = {(row['function'], row['name']): row for row in observation['rows']}
    outcomes = []
    for expected in json.loads((probe / 'truth.json').read_text(encoding='utf-8')):
        row = rows.get((expected['function'], expected['expression']))
        actual = row['trace']['range'] if row and row['trace']['code'] == 'TYPE_MATCH' else None
        outcomes.append(dict(expected, actual=actual, code=row['code'] if row else 'READER_ROW_NOT_FOUND',
                             outcome='MATCH_TRUTH' if row and actual == expected['range'] else
                             'UNRESOLVED' if actual is None else 'WRONG_RANGE'))
    report['excel_truth_comparison'] = outcomes
    report['macro_keyword_variants'] = {
        name: diagnostic.trace_expression('can_chg_p->active', str(probe / name), 'global')
        for name in ('probe.c', 'undefined.c')}
    for name, line, macro_defined in (('probe.c', 21, True), ('undefined.c', 2, False)):
        variant = report['macro_keyword_variants'][name]
        variant.update(expected_declaration=dict(file=str(probe / name), line=line, variable='can_chg_p',
                                                 type='can_chg_t', empty_macro_defined=macro_defined))
        if not variant['candidates']:
            report['parse_failures'].append(dict(variant['expected_declaration'],
                code='SILENT_DECLARATION_OMISSION' if macro_defined else 'INVALID_UNEXPANDED_KEYWORD',
                evidence='INDEPENDENT_SYNTHETIC_TRUTH'))
    report['duplicate_tag_other_scope'] = diagnostic.trace_expression(
        'control.mode', str(probe / 'alternate.c'), 'global')
    save(root / 'diagnosis.json', report)
    for suite in ('functional', 'stress100'):
        folder = root / suite
        folder.mkdir()
        generate(folder, suite)
        jobs = json.loads((folder / 'truth.json').read_text(encoding='utf-8'))
        paths = [p for p in folder.rglob('*') if p.suffix in ('.c', '.h')]
        analysis = CHDiagnostics(paths, build_options={
            folder / j['source']: j['args'] for j in jobs if j['db']})
        audit = audit_clang_truth(analysis, folder, jobs)
        if suite == 'stress100':
            audit['preprocessor_traces'] = [
                analysis.trace_expression('local.selected', str(folder / job['source']),
                                          job['accesses'][1]['function'])
                for job in jobs]
        save(folder / 'audit.json', audit)
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    parser.add_argument('--backend', choices=('openpyxl', 'com'), default='openpyxl')
    options = parser.parse_args()
    print(json.dumps(run(options.output, options.backend)['counts'], indent=2))
