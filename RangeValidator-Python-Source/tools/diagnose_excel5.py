"""Run staged synthetic Excel-5 measurements with actual desktop Excel."""
import argparse
from collections import Counter
from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys
from time import perf_counter
import traceback

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel1_synthetic import assess
from excel5_metrics import Probe, Trace
from excel5_synthetic import build_limits, build_mixed, build_write_file
from rangevalidator.extraction_diagnostics import diagnose_session
from rangevalidator.models import ReadFailure, ReaderLimits
from rangevalidator.readers import Win32Runtime, read_com
from rangevalidator.workflow import run_comparison
from rangevalidator.writer import save_selected


def read_json(path):
    return json.loads(path.read_text(encoding='utf-8'))


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compare(source, paths, probe):
    return run_comparison(None, [source], paths, backend='com', diagnostics=Trace(probe))


def validate(truth, session, reports):
    result_index = {(Path(r.file).name, r.sheet, r.cell): r for r in session.results}
    observed = {}
    duplicates = 0
    for report in reports:
        for sheet in report['sheets']:
            if not sheet['function_sheet']:
                assert 'SHEET_BODY_READ_STARTED' not in sheet['states']
            for row in sheet['variables']:
                key = (Path(row['file']).name, row['sheet'], row['name_cell'])
                duplicates += key in observed
                observed[key] = row
    counts = Counter(total=len(truth), duplicates=duplicates)
    errors = []
    for expected in truth:
        key = (expected['file'], expected['sheet'], expected['name_cell'])
        row = observed.pop(key, None)
        if row is None:
            counts['missing'] += 1
            errors.append(dict(key=key, issue='missing'))
            continue
        wrong_link = row['range_cell'] != expected['range_cell']
        wrong_format = row['format'] != expected['format']
        counts['wrong_links'] += wrong_link
        counts['format_errors'] += wrong_format
        actual = result_index.get((key[0], key[1], row['range_cell']))
        if expected['status'] == 'UNRESOLVED':
            correct = actual is None and not row['extracted'] and row['state'] == 'RANGE_CELL_UNRESOLVED'
            counts['unresolved'] += correct and not wrong_link
        else:
            correct = actual is not None and row['extracted'] and (
                actual.status, actual.candidate, actual.existing
            ) == (expected['status'], expected['candidate'], expected['original_range'])
            counts['exact'] += correct and not wrong_link and not wrong_format
        if not correct or wrong_link or wrong_format or (row['name'], row['section']) != (expected['name'], expected['section']):
            errors.append(dict(key=key, expected=expected, observed=row,
                               status=actual.status if actual else None))
    counts['unexpected'] = len(observed)
    counts['disagreements'] = len(errors)
    return dict(counts=counts, disagreements=errors[:25], unexpected=list(observed)[:25])


def mixed_case(source, root, count):
    path = root / f'mixed_{count}.xlsx'
    truth = build_mixed(path, count)
    write_json(root / f'mixed_{count}_truth.json', truth)
    before = sha(path)
    start = perf_counter()
    with Probe() as probe:
        session = compare(source, [path], probe)
    elapsed = perf_counter() - start
    reports = diagnose_session(session)
    check = validate(truth, session, reports)
    states = [asdict(f) for f in session.files]
    unchanged = sha(path) == before
    assert unchanged
    return dict(**check, file_size_bytes=path.stat().st_size, function_sheets=count,
                analysis_seconds=elapsed, per_file_seconds=session.excel1_timings,
                originals_unchanged=unchanged, file_status=states, read_errors=session.errors,
                metrics=probe.result())


def writer_case(source, root, count, split):
    tag = f'write_{count}_{"split" if split else "single"}'
    paths = [root / f'{tag}_{index}.xlsx' for index in range(1, 3 if split else 2)]
    expected = []
    for i, path in enumerate(paths):
        expected.extend(build_write_file(path, count // len(paths), 'B' if i == 0 else 'A'))
    write_json(root / f'{tag}_truth.json', expected)
    before = {p: sha(p) for p in paths}
    with Probe() as probe:
        session = compare(source, paths, probe)
        prepared = perf_counter()
        # The first table in each file is the edit batch. Later cells include safety guards.
        wanted = {(v['file'], v['sheet'], v['range_cell']): v for path in paths for v in
                  [v for v in expected if v['file'] == path.name][:count // len(paths)]}
        chosen = [r for r in session.results if (Path(r.file).name, r.sheet, r.cell) in wanted]
        assert len(chosen) == count and all(r.candidate for r in chosen)
        selection_seconds = perf_counter() - prepared
        start = perf_counter()
        folder, outcomes = save_selected(session, [r.id for r in chosen], root / tag)
        writer_total = perf_counter() - start
        reopened = {}
        start = perf_counter()
        for path in paths:
            fixed = folder / (path.stem + '_fixed.xlsx')
            reopened[path.name] = read_com(fixed, function_sheets_only=True, _trace=Trace(probe))
        reread = perf_counter() - start
        verification_start = perf_counter()
        receipts = []
        for r in chosen:
            filename = Path(r.file).name
            sheet = next(s for s in reopened[filename].sheets if s.name == r.sheet)
            cell = sheet.cells.get(r.cell)
            target = wanted[(filename, r.sheet, r.cell)]
            outcome = outcomes[r.id]
            receipts.append(dict(file=filename, sheet=r.sheet, cell=r.cell,
                original=r.existing, expected=r.expected, actual=cell.value if cell else None,
                formula=cell.is_formula if cell else None, status=outcome['status'],
                correct=bool(cell and cell.value == r.expected and not cell.is_formula
                             and r.existing == target['original_range'] and outcome['status'] == 'SAVED_VERIFIED')))
        for v in expected:
            if v['status'] == 'REVIEW':
                sheet = reopened[v['file']].sheets[0]
                cell = sheet.cells[v['range_cell']]
                assert cell.is_formula if v['reason'] == 'FORMULA' else cell.merge_anchor != cell.address
        verify = perf_counter() - verification_start
    check = validate(expected, session, diagnose_session(session))
    unchanged = all(sha(path) == before[path] for path in paths)
    write_json(root / f'{tag}_cells.json', receipts)
    return dict(**check, files=len(paths), file_size_bytes=sum(p.stat().st_size for p in paths),
                requested=count, saved_verified=sum(r['correct'] for r in receipts),
                originals_unchanged=unchanged, output=str(folder), candidate_selection_seconds=selection_seconds,
                writer_total_seconds=writer_total, independent_com_reread_seconds=reread,
                independent_value_check_seconds=verify, metrics=probe.result())


def limits_case(source, root):
    path = root / 'read_limits.xlsx'
    truth = build_limits(path)
    write_json(root / 'read_limits_truth.json', truth)
    before = sha(path)
    with Probe() as probe:
        session = compare(source, [path], probe)
    report, = diagnose_session(session)
    assert session.files[0].status == 'PARTIAL'
    assert len(session.errors) == 1
    error = session.errors[0]
    assert (error['stage'], error['sheet_name']) == ('com.bounds.limit', '1. Wide_limit')
    assert error['attempted_bounds'] == (1, 1, 20, 257)
    assert all(r.sheet != '1. Wide_limit' for r in session.results)
    check = validate([v for v in truth if v['sheet'] == '2. Long_table'], session, [report])
    return dict(**check, file_size_bytes=path.stat().st_size,
                originals_unchanged=sha(path) == before,
                complete_sheet='2. Long_table', incomplete_sheet='1. Wide_limit',
                affected_sections=['Global(Input)'], analysis_complete=False,
                error=error, file_status=asdict(session.files[0]), metrics=probe.result())


def error_case(source, root):
    valid = root / 'after_failures.xlsx'
    build_write_file(valid, 2)
    corrupt = root / 'corrupt.xlsx'
    corrupt.write_bytes(b'This is a synthetic invalid XLSX container.')
    unsupported = root / 'unsupported.csv'
    unsupported.write_text('not,xlsx', encoding='utf-8')
    before = {path: sha(path) for path in (valid, corrupt, unsupported)}
    with Probe() as probe:
        session = compare(source, [root / 'missing.xlsx', corrupt, unsupported, valid], probe)
    states = [asdict(f) for f in session.files]
    assert [f.status for f in session.files] == ['ERROR', 'ERROR', 'ERROR', 'COMPLETED']
    assert session.results and all(Path(r.file) == valid for r in session.results)
    return dict(files=states, errors=session.errors, originals_unchanged=all(sha(p) == h for p, h in before.items()),
                metrics=probe.result())


def sparse_budget_case(root):
    from openpyxl import Workbook

    path = root / 'sparse_budget.xlsx'
    workbook = Workbook()
    sparse = workbook.active
    sparse.title = '1. Sparse'
    sparse['A1'] = 'first'
    sparse['A393217'] = 'last'
    workbook.create_sheet('2. Normal')['A1'] = 'sibling'
    workbook.save(path)
    workbook.close()
    before = sha(path)
    with Probe() as probe:
        bounded = read_com(path, function_sheets_only=True, _trace=Trace(probe))
        raised_budget = read_com(path, ReaderLimits(max_cells=500_000),
                                 function_sheets_only=True, _trace=Trace(probe))
    assert bounded.partial and len(bounded.read_errors) == 1
    error = bounded.read_errors[0]
    assert error['attempted_bounds'] == (1, 1, 393217, 1)
    assert error['sheet_name'] == '1. Sparse' and error['stage'] == 'com.bounds.limit'
    assert not bounded.sheets[0].cells and bounded.sheets[1].cells['A1'].value == 'sibling'
    assert not raised_budget.partial and raised_budget.sheets[0].cells['A393217'].value == 'last'
    return dict(default_cell_budget=250000, raised_test_budget=500000, default_failure=error,
                recovered_tail='A393217', originals_unchanged=sha(path) == before,
                file_size_bytes=path.stat().st_size, metrics=probe.result())


def locked_case(source, root):
    import win32file

    path = root / 'exclusive_lock.xlsx'
    valid = root / 'after_lock.xlsx'
    build_write_file(path, 2)
    build_write_file(valid, 2)
    before = {p: sha(p) for p in (path, valid)}
    handle = win32file.CreateFile(str(path), win32file.GENERIC_READ, 0, None,
                                  win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
    try:
        with Probe() as probe:
            session = compare(source, [path, valid], probe)
    finally:
        handle.Close()
    assert [f.status for f in session.files] == ['ERROR', 'COMPLETED']
    assert all(Path(r.file) == valid for r in session.results)
    return dict(files=[asdict(f) for f in session.files], errors=session.errors,
                originals_unchanged=all(sha(p) == h for p, h in before.items()), metrics=probe.result())


def ownership_case(source, root):
    from queue import Queue
    from threading import Event, Thread

    # Load pywin32 on the main thread before initializing the watcher STA.
    Win32Runtime()
    path = root / 'watch_owned_document.xlsx'
    build_write_file(path, 1)
    before = sha(path)
    ready = Queue()
    finished = Queue()
    verify = Event()

    def watch_document():
        # The other open workbook lives in its own STA, as a separate user Excel would.
        runtime = Win32Runtime()
        runtime.initialize()
        app = watch = None
        pid = None
        owned = False
        failure = None
        try:
            existing = runtime.process_ids()
            app = runtime.dispatch()
            pid = runtime.pid(app)
            assert pid and pid not in existing, 'Cannot prove ownership of watcher Excel'
            owned = True
            app.Visible = False
            app.DisplayAlerts = False
            watch = app.Workbooks.Open(Filename=str(path), ReadOnly=True, UpdateLinks=0, AddToMru=False)

            def state():
                return (app.Hwnd, app.Workbooks.Count, watch.FullName, watch.Saved,
                        watch.Worksheets.Item(1).Range('A1').Value2)

            initial = state()
            ready.put((pid, None))
            while not verify.wait(0.05):
                runtime.pythoncom.PumpWaitingMessages()
            assert state() == initial
        except Exception as exc:
            failure = exc
            if ready.empty() and not verify.is_set():
                ready.put((pid, exc))
        finally:
            try:
                if watch is not None:
                    watch.Close(SaveChanges=False)
            except Exception as exc:
                failure = failure or exc
            finally:
                watch = None
                try:
                    if app is not None and owned:
                        app.Quit()
                except Exception as exc:
                    failure = failure or exc
                finally:
                    app = None
                    runtime.uninitialize()
            try:
                if owned:
                    runtime.wait_exit(pid)
            except Exception as exc:
                failure = failure or exc
            finished.put(failure)

    watcher = Thread(target=watch_document, name='synthetic-excel-watcher')
    watcher.start()
    try:
        pid, opening_error = ready.get(timeout=60)
        if opening_error:
            raise opening_error
        test = root / 'ownership_probe'
        test.mkdir(exist_ok=True)
        receipt = writer_case(source, test, 10, False)
        assert receipt['saved_verified'] == 10 and not receipt['counts']['disagreements']
        assert pid not in receipt['metrics']['excel_pids']
    finally:
        verify.set()
        watcher.join()
    failure = finished.get_nowait()
    if failure:
        raise failure
    return dict(synthetic_open_workbook_preserved=True, watcher_exited=True, watcher_pid=pid,
                writer=receipt, originals_unchanged=sha(path) == before and receipt['originals_unchanged'])


def baseline_case(source):
    original = ROOT.parent / 'excel1-diagnostic' / 'synthetic'
    reference = ROOT.parent / 'excel4-diagnostic'
    truth = read_json(original / 'ground_truth.json')
    paths = [original / name for name in dict.fromkeys(v['file'] for v in truth)]
    before = {p: sha(p) for p in paths}
    with Probe() as probe:
        session = compare(source, paths, probe)
    reports = diagnose_session(session)
    observed = assess(truth, reports)
    summary = Counter(v['format'] for report in reports for sheet in report['sheets']
                      for v in sheet['variables'] if v['extracted'])
    bulk = next(r for r in reports if Path(r['file']).name == 'large_120.xlsx')
    assert observed['correct_extractions'] == 7788 and observed['missing'] == 5
    assert observed['wrong_cell_links'] == 0 and summary == {'A': 3853, 'B': 3935}
    assert bulk['counts']['function_candidates'] == 120
    assert sum(s['extracted_variables'] for s in bulk['sheets']) == 7260
    assert not session.errors
    matrix = ROOT.parent / 'excel3-diagnostic' / 'synthetic/a_105_sheets.xlsx'
    matrix_before = sha(matrix)
    with Probe() as matrix_probe:
        matrix_session = compare(source, [matrix], matrix_probe)
    matrix_report, = diagnose_session(matrix_session)
    positions = {s['sheet']: [[v['name_cell'], v['range_cell']] for v in s['variables'] if v['extracted']]
                 for s in matrix_report['sheets'] if s['function_sheet']}
    assert positions == read_json(ROOT.parent / 'excel3-diagnostic' / 'matrix_truth.json')
    assert not matrix_session.errors
    assert observed == read_json(reference / 'assessment.json')
    return dict(assessment=observed, formats=summary, mixed_120_variables=7260,
                a_matrix_sheets=len(positions), a_matrix_pairs=sum(map(len, positions.values())),
                originals_unchanged=sha(matrix) == matrix_before and all(sha(p) == h for p, h in before.items()),
                metrics=probe.result(), matrix_metrics=matrix_probe.result())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    parser.add_argument('part', choices=('mixed', 'writer', 'edges', 'baseline', 'isolation'))
    parser.add_argument('--case', help='Run only one named case in the selected part')
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    root = output / 'synthetic'
    root.mkdir(exist_ok=True)
    source = ROOT.parent / 'excel1-diagnostic' / 'synthetic/excel1_types.h'
    cases = {
        'mixed': [(f'mixed_{n}', lambda n=n: mixed_case(source, root, n)) for n in (120, 200, 300)],
        'writer': [(f'write_{n}_{mode}', lambda n=n, mode=mode: writer_case(source, root, n, mode == 'split'))
                   for n in (10, 100, 500, 1000) for mode in ('single', 'split')],
        'edges': [('limits', lambda: limits_case(source, root)), ('errors', lambda: error_case(source, root))],
        'isolation': [('sparse_budget', lambda: sparse_budget_case(root)),
                      ('locked', lambda: locked_case(source, root)),
                      ('ownership', lambda: ownership_case(source, root))],
        'baseline': [('baseline', lambda: baseline_case(source))],
    }
    failed = False
    selected = [(name, run) for name, run in cases[args.part] if args.case is None or args.case == name]
    if not selected:
        parser.error(f'Unknown case {args.case!r} for part {args.part!r}')
    for name, run in selected:
        started = perf_counter()
        try:
            result = run()
            counts = result.get('counts', {})
            ok = result.get('originals_unchanged', True) and not any(counts.get(key, 0) for key in
                ('disagreements', 'wrong_links', 'format_errors', 'missing', 'duplicates', 'unexpected'))
            ok = ok and not result.get('read_errors')
            if 'requested' in result:
                ok = ok and result['requested'] == result['saved_verified']
            result['outcome'] = 'SUCCESS' if ok else 'FAILURE'
        except Exception as exc:
            result = dict(outcome='FAILURE', error=repr(exc), traceback=traceback.format_exc())
            if isinstance(exc, ReadFailure):
                result['com_failure'] = exc.as_dict()
        result['elapsed_seconds'] = perf_counter() - started
        result['finished_utc'] = datetime.now(timezone.utc).isoformat()
        write_json(output / f'{name}.json', result)
        print(f'{name}: {result["outcome"]} {result["elapsed_seconds"]:.2f}s '
              f'{result.get("counts", result.get("error", ""))}', flush=True)
        failed |= result['outcome'] != 'SUCCESS'
    return int(failed)


if __name__ == '__main__':
    raise SystemExit(main())
