"""Verify Excel-4 with actual COM using Excel-1 and Excel-3 evidence."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel1_synthetic import assess
from excel4_synthetic import build_b_safety, build_b_workbook
from rangevalidator.extraction_diagnostics import diagnose_session, diagnose_workbook
from rangevalidator.readers import read_com
from rangevalidator.workflow import run_comparison
from rangevalidator.writer import save_selected


FIELDS = ('sheet', 'section', 'section_cell', 'name', 'name_cell', 'range_cell', 'comparison_status')


def load_json(path):
    return json.loads(path.read_text(encoding='utf-8'))


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def entries(reports):
    return {
        (Path(v['file']).name, v['sheet'], v['name'], v['name_cell'], v['range_cell']):
        (v['format'], v['comparison_status'])
        for report in reports for sheet in report['sheets'] for v in sheet['variables'] if v['extracted']
    }


def compare(source, path):
    before = digest(path)
    session = run_comparison(None, [source], [path], backend='com')
    report = diagnose_session(session)[0]
    assert report['backend'] == 'com' and not report['read_errors'], path
    assert digest(path) == before, path
    assert all('SHEET_BODY_READ_STARTED' not in s['states'] for s in report['sheets']
               if not s['function_sheet']), path
    return session, report


def validate_b(report, truth):
    actual = [{key: var[key] for key in FIELDS} for sheet in report['sheets']
              for var in sheet['variables'] if var['extracted']]
    assert actual == truth, report['file']


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    synthetic = output / 'synthetic'
    synthetic.mkdir()
    excel1 = ROOT.parent / 'excel1-diagnostic' / 'synthetic'
    excel3 = ROOT.parent / 'excel3-diagnostic'
    source = excel1 / 'excel1_types.h'
    truth = load_json(excel1 / 'ground_truth.json')
    baseline = entries(load_json(excel3 / 'diagnostics.json'))

    reports = []
    for name in dict.fromkeys(v['file'] for v in truth):
        _, report = compare(source, excel1 / name)
        reports.append(report)
        print(f'Actual COM baseline: {name}', flush=True)
    current = entries(reports)
    assert all(current.get(key) == value for key, value in baseline.items())
    counts = dict(Counter(value[0] for value in current.values()))
    assert counts == {'A': 3853, 'B': 3935}
    added = sorted(current.keys() - baseline.keys())
    assert len(added) == 10 and all(current[key][0] == 'B' for key in added)
    assessment = assess(truth, reports)
    assert assessment['correct_extractions'] == 7788 and assessment['wrong_cell_links'] == 0
    assert assessment['missing'] == 5
    missing = next(r for r in reports if Path(r['file']).name == 'missing_header.xlsx')
    unresolved = [v for s in missing['sheets'] for v in s['variables']]
    assert len(unresolved) == 5 and all(v['range_cell'] is None and not v['extracted'] for v in unresolved)
    mixed = next(r for r in reports if Path(r['file']).name == 'large_120.xlsx')
    assert mixed['counts']['function_candidates'] == 120
    assert sum(s['extracted_variables'] for s in mixed['sheets']) == 7260
    write_json(output / 'assessment.json', assessment)
    write_json(output / 'diagnostics.json', reports)

    _, matrix = compare(source, excel3 / 'synthetic/a_105_sheets.xlsx')
    matrix_truth = load_json(excel3 / 'matrix_truth.json')
    actual_a = {s['sheet']: [[v['name_cell'], v['range_cell']] for v in s['variables']
                           if v['extracted']] for s in matrix['sheets'] if s['function_sheet']}
    assert actual_a == matrix_truth
    print('Actual COM A matrix: 105 sheets, 15981 pairs', flush=True)

    stress = []
    for count in (1, 10, 100):
        path = synthetic / f'b_{count}.xlsx'
        expected = build_b_workbook(path, count)
        _, report = compare(source, path)
        validate_b(report, expected)
        write_json(output / f'b_{count}_truth.json', expected)
        write_json(output / f'b_{count}_diagnostics.json', report)
        stress.append(dict(sheets=count, variables=len(expected), wrong_links=0))
        print(f'Actual COM B matrix: {count} sheets, {len(expected)} pairs', flush=True)

    path = synthetic / 'b_safety.xlsx'
    expected = build_b_safety(path)
    before = digest(path)
    session, report = compare(source, path)
    validate_b(report, expected)
    selected = [r for r in session.results if r.cell in ('E6', 'E8') and r.candidate]
    assert [(r.cell, r.status) for r in selected] == [('E6', 'MISMATCH'), ('E8', 'CREATE')]
    folder, outcomes = save_selected(session, [r.id for r in selected], output / 'writer')
    assert all(outcomes[r.id]['status'] == 'SAVED_VERIFIED' for r in selected), outcomes
    saved = read_com(folder / 'b_safety_fixed.xlsx', function_sheets_only=True)
    cells = saved.sheets[0].cells
    assert all(cells[r.cell].value == r.expected for r in selected)
    assert cells['E9'].is_formula and cells['E10'].merge_anchor == 'D10'
    assert cells['D10'].value == '0~3' and cells['H23'].value == '0~3'
    assert digest(path) == before
    write_json(output / 'b_safety.json', report)
    print('Actual COM B Writer: E6 MISMATCH, E8 CREATE, both SAVED_VERIFIED', flush=True)

    sample = ROOT.parent / 'B형.xlsx'
    sample_before = digest(sample)
    sample_report = diagnose_workbook(read_com(sample, function_sheets_only=True))
    assert digest(sample) == sample_before
    sample_tables = [t for s in sample_report['sheets'] for t in s['tables'] if t.get('format') == 'B']
    assert [t['section'] for t in sample_tables] == [
        'Parameter(Input)', 'Parameter(Output)', 'Global(Input)', 'Global(Output)', 'Local', 'Return']
    write_json(output / 'sample.json', sample_report)
    write_json(output / 'validation.json', dict(
        counts=counts, recovered=added, legacy_results_preserved=True, unresolved=len(unresolved),
        mixed_sheets=120, mixed_variables=7260, a_matrix_sheets=len(actual_a),
        a_matrix_variables=sum(map(len, actual_a.values())), b_stress=stress,
        wrong_cell_links=0, excluded_body_reads=0, sample_sections=len(sample_tables),
        originals_unchanged=True,
        writer=[dict(cell=r.cell, comparison=r.status, reopened_value=cells[r.cell].value,
                     **outcomes[r.id]) for r in selected]))


if __name__ == '__main__':
    main()
