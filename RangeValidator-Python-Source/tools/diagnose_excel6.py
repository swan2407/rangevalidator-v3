"""Actual COM boundaries, Open diagnostics, Writer scaling and policy checks."""
import argparse
from dataclasses import asdict
from pathlib import Path
import sys
from time import perf_counter
import traceback
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import win32file

from diagnose_excel5 import (baseline_case, compare, limits_case, ownership_case,
                            read_json, sha, sparse_budget_case, write_json)
from excel5_metrics import Probe, Trace
from excel5_synthetic import build_write_file
from rangevalidator.extraction_diagnostics import diagnose_session
from rangevalidator.manual_edits import ReviewRequest
from rangevalidator.models import ReadFailure
from rangevalidator.readers import Win32Runtime, read_com
from rangevalidator.writer import save_selected


def _candidate(sheet):
    sheet.append(['Name', 'Type', 'Range'])
    sheet.append(['obj.v001', 'unsigned int', '0~3'])


def boundary_case(source, root, expected):
    path = root / f'budget_{expected}.xlsx'
    book = Workbook()
    seed = book.active
    seed.title = '1. Seed'
    seed['A1'], seed[f'A{expected - 20}'] = 'first', 'budget witness'
    target = book.create_sheet('2. Boundary')
    _candidate(target)
    target['E4'] = 'extent witness'
    target.cell(900000, 16384).font = Font(bold=True)
    if expected > 250000:
        _candidate(book.create_sheet('3. Sibling'))
    book.save(path)
    book.close()
    before = sha(path)
    with Probe() as probe:
        session = compare(source, [path], probe)
    matches = [(r.sheet, r.cell, r.candidate) for r in session.results]
    if expected <= 250000:
        assert session.files[0].status == 'COMPLETED' and not session.errors
        assert matches == [('2. Boundary', 'C2', True)]
    else:
        assert session.files[0].status == 'PARTIAL'
        assert matches == [('3. Sibling', 'C2', True)]
        error, = session.errors
        assert error['sheet_name'] == '2. Boundary' and error['attempted_bounds'] == (1, 1, 4, 5)
        assert error['stage'] == 'com.bounds.limit'
        rejected = root / 'must_not_save_partial'
        try:
            save_selected(session, [f'{path}|2. Boundary|C2'], rejected)
        except ValueError:
            assert not rejected.exists()
        else:
            raise AssertionError('Reader-incomplete cell was accepted for saving')
    assert sha(path) == before
    return dict(file=str(path), total_content_rectangle_cells_before_sibling=expected,
                rectangles={'1. Seed': [expected-20, 1], '2. Boundary': [4, 5]},
                status=session.files[0].status, editable=matches, errors=session.errors,
                originals_unchanged=True, metrics=probe.result())


def merge_budget_case(source, root):
    path = root / 'merge_budget.xlsx'
    book = Workbook()
    at = book.active
    at.title = '1. Exactly'
    at['A1'] = 'merged title'
    at.merge_cells('A1:A2')
    at['A200000'] = 'extent'
    # The first sheet leaves 50,000 cells. A separate direct read with a
    # higher workbook budget tests the independent 200,000 merge cap.
    over = book.create_sheet('2. Over')
    over['A1'] = 'merged title'
    over.merge_cells('A1:A2')
    over['A200001'] = 'extent'
    _candidate(book.create_sheet('3. Sibling'))
    book.save(path)
    book.close()
    from rangevalidator.models import ReaderLimits
    before = sha(path)
    with Probe() as probe:
        result = read_com(path, ReaderLimits(max_cells=500000),
                          function_sheets_only=True, _trace=Trace(probe))
    error, = result.read_errors
    assert result.partial and error['sheet_name'] == '2. Over'
    assert error['attempted_bounds'] == (1, 1, 200001, 1)
    assert 'Merge scan budget' in error['limit_reason']
    assert result.sheets[0].merged_ranges == ['A1:A2']
    assert result.sheets[0].cells['A200000'].value == 'extent'
    assert result.sheets[-1].cells['C2'].value == '0~3'
    assert sha(path) == before
    return dict(file=str(path), exactly=200000, over=200001, sheet_error=error,
                sibling_preserved=True, originals_unchanged=True, metrics=probe.result())


def column_edge_case(source, root):
    path = root / 'column_edge.xlsx'
    book = Workbook()
    at = book.active
    at.title = '1. At_column_cap'
    _candidate(at)
    at['IV2'] = 'column 256'
    over = book.create_sheet('2. Beyond_column_cap')
    _candidate(over)
    over['IW2'] = 'column 257'
    _candidate(book.create_sheet('3. Sibling'))
    book.save(path)
    book.close()
    before = sha(path)
    with Probe() as probe:
        session = compare(source, [path], probe)
    assert session.files[0].status == 'PARTIAL'
    error, = session.errors
    assert error['attempted_bounds'] == (1, 1, 2, 257)
    assert [(r.sheet, r.cell, r.candidate) for r in session.results] == [
        ('1. At_column_cap', 'C2', True), ('3. Sibling', 'C2', True)]
    assert session.books[str(path)].sheets[0].cells['IV2'].value == 'column 256'
    assert sha(path) == before
    return dict(file=str(path), admitted_column=256, rejected_column=257, status='PARTIAL',
                sheet_error=error, originals_unchanged=True, metrics=probe.result())


def writer_scale(source, root, count):
    folder = root / f'write_{count}_files'
    folder.mkdir(exist_ok=True)
    paths = [folder / f'input_{index+1}.xlsx' for index in range(count)]
    per_file = 1000 // count
    truth = {}
    for index, path in enumerate(paths):
        rows = build_write_file(path, per_file, 'B' if index % 2 == 0 else 'A')
        truth[path] = rows[:per_file]
    hashes = {p: sha(p) for p in paths}
    receipts = []
    with Probe() as probe:
        session = compare(source, paths, probe)
        assert not session.errors
        result_index = {(Path(r.file), r.sheet, r.cell): r for r in session.results}
        targets = [result_index[p, row['sheet'], row['range_cell']]
                   for p, rows in truth.items() for row in rows]
        assert len(targets) == 1000 and all(r.candidate for r in targets)
        start = perf_counter()
        output, outcomes = save_selected(session, [r.id for r in targets], folder)
        total = perf_counter() - start
        start = perf_counter()
        for path in paths:
            fixed = output / (path.stem + '_fixed.xlsx')
            saved = read_com(fixed, function_sheets_only=True, _trace=Trace(probe))
            assert not saved.partial
            sheets = {s.name: s for s in saved.sheets}
            for v in truth[path]:
                r = result_index[path, v['sheet'], v['range_cell']]
                cell = sheets[r.sheet].cells[r.cell]
                assert r.existing == v['original_range'] and r.status == v['status']
                assert cell.value == r.expected and not cell.is_formula
                assert outcomes[r.id]['status'] == 'SAVED_VERIFIED'
                receipts.append(dict(file=path.name, sheet=r.sheet, cell=r.cell, original=r.existing,
                    expected=r.expected, actual=cell.value, formula=cell.is_formula,
                    status=outcomes[r.id]['status']))
        reread = perf_counter() - start
    assert all(sha(p) == hashes[p] for p in paths)
    write_json(folder / 'cells.json', receipts)
    return dict(files=count, cells=len(receipts), saved_verified=len(receipts),
                input_bytes=sum(p.stat().st_size for p in paths), writer_total_seconds=total,
                independent_reread_and_check_seconds=reread, originals_unchanged=True,
                output=str(output), metrics=probe.result())


def writer_policy(source, root):
    folder = root / 'writer_policy'
    folder.mkdir(exist_ok=True)
    bad, good = folder / 'locked.xlsx', folder / 'normal.xlsx'
    bad_truth = build_write_file(bad, 1)
    good_truth = build_write_file(good, 2)
    book = load_workbook(good)
    manual = book.create_sheet('2. Manual')
    manual.append(['Name', 'Type', 'Range'])
    manual.append(['obj.v001', 'unsigned int', 'needs review'])
    book.save(good)
    book.close()
    hashes = {p: sha(p) for p in (bad, good)}
    with Probe() as probe:
        session = compare(source, [bad, good], probe)
        assert not session.errors
        index = {(Path(r.file), r.sheet, r.cell): r for r in session.results}
        bad_result = index[bad, bad_truth[0]['sheet'], bad_truth[0]['range_cell']]
        auto = [index[good, v['sheet'], v['range_cell']] for v in good_truth[:2]]
        manual = index[good, '2. Manual', 'C2']
        assert manual.status == 'REVIEW'
        protected = [index[good, v['sheet'], v['range_cell']] for v in good_truth if v['status'] == 'REVIEW']
        assert len(protected) == 2 and all(r.status == 'REVIEW' and not r.candidate for r in protected)
        assert all(v['range_cell'] is None for v in good_truth if v['status'] == 'UNRESOLVED')
        requests = [ReviewRequest(r.id, '0~7', True) for r in [manual, *protected]]
        lock = win32file.CreateFile(str(bad), win32file.GENERIC_READ, 0, None,
                                   win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
        try:
            output, outcomes = save_selected(session, [r.id for r in [bad_result, *auto]], folder,
                                             manual_requests=requests)
        finally:
            lock.Close()
        assert outcomes[bad_result.id]['status'] == 'SAVE_FAILED'
        assert not (output / 'locked_fixed.xlsx').exists()
        saved = read_com(output / 'normal_fixed.xlsx', function_sheets_only=True, _trace=Trace(probe))
        sheets = {s.name: s for s in saved.sheets}
        for r, kind in zip([*auto, manual], ('AUTO_MISMATCH', 'AUTO_CREATE', 'MANUAL_REVIEW')):
            assert outcomes[r.id]['status'] == 'SAVED_VERIFIED' and outcomes[r.id]['edit_kind'] == kind
            assert sheets[r.sheet].cells[r.cell].value == '0~7'
        for r in protected:
            assert outcomes[r.id]['status'] == 'SKIPPED'
        report = read_com(output / 'RangeValidator_Report.xlsx', _trace=Trace(probe))
        assert len(report.sheets) == 5 and not report.partial
    assert all(sha(p) == hashes[p] for p in (bad, good))
    return dict(originals_unchanged=True, outputs=str(output), report_sheets=[s.name for s in report.sheets],
                outcomes=outcomes, unresolved=sum(v['status'] == 'UNRESOLVED' for v in good_truth),
                metrics=probe.result())


class AuditWorkbook:
    def __init__(self, book, record):
        self.book, self.record = book, record

    def __getattr__(self, name):
        return getattr(self.book, name)

    def Close(self, **kwargs):
        self.record['close_target'] = str(self.book.Name)
        self.record['close_arguments'] = kwargs
        result = self.book.Close(**kwargs)
        self.record['close_returned'] = True
        return result


class AuditBooks:
    def __init__(self, books, pid, audit):
        self.books, self.pid, self.audit = books, pid, audit

    def __getattr__(self, name):
        return getattr(self.books, name)

    def Open(self, **kwargs):
        path = Path(kwargs['Filename'])
        try:
            with path.open('rb'):
                accessible = True
        except OSError:
            accessible = False
        record = dict(file=path.name, exists=path.is_file(), readable_by_runner=accessible,
                      owned_pid=self.pid, dispatch_and_ownership_verified=True, before_count=int(self.books.Count),
                      arguments={**kwargs, 'Filename': path.name})
        self.audit.append(record)
        try:
            result = self.books.Open(**kwargs)
            return AuditWorkbook(result, record) if result is not None else None
        except Exception as exc:
            def scrub(value):
                return str(value).replace(str(path), path.name).replace(str(path.parent), '<synthetic>')
            record.update(code=getattr(exc, 'hresult', None), message=scrub(exc),
                          excepinfo=scrub(getattr(exc, 'excepinfo', None)))
            raise
        finally:
            record['after_count'] = int(self.books.Count)


class AuditApp:
    def __init__(self, app, pid, audit):
        object.__setattr__(self, '_app', app)
        object.__setattr__(self, '_pid', pid)
        object.__setattr__(self, '_audit', audit)

    def __setattr__(self, name, value):
        setattr(self._app, name, value)

    def __getattr__(self, name):
        if name == 'Workbooks':
            return AuditBooks(self._app.Workbooks, self._pid, self._audit)
        return getattr(self._app, name)


def open_audit(source, root):
    absent, corrupt, locked, valid = (root / name for name in
        ('audit_missing.xlsx', 'audit_corrupt.xlsx', 'audit_locked.xlsx', 'audit_normal.xlsx'))
    corrupt.write_bytes(b'Synthetic invalid XLSX')
    build_write_file(locked, 1)
    build_write_file(valid, 1)
    before = {p: sha(p) for p in (corrupt, locked, valid)}
    audit = []
    dispatch = Win32Runtime.dispatch_owned

    def traced(runtime):
        app, pid = dispatch(runtime)
        return AuditApp(app, pid, audit), pid

    lock = win32file.CreateFile(str(locked), win32file.GENERIC_READ, 0, None,
                               win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
    try:
        with Probe() as probe, patch.object(Win32Runtime, 'dispatch_owned', traced):
            session = compare(source, [absent, corrupt, locked, valid], probe)
    finally:
        lock.Close()
    assert [f.status for f in session.files] == ['ERROR', 'ERROR', 'ERROR', 'COMPLETED']
    assert all(item['before_count'] == item['after_count'] == 0 for item in audit[:3])
    assert audit[-1]['before_count'] == 0 and audit[-1]['after_count'] == 1
    assert not any('close_target' in item for item in audit[:3])
    assert audit[-1]['close_target'] == valid.name and audit[-1]['close_returned']
    assert audit[-1]['close_arguments'] == {'SaveChanges': False}
    cleaned = [asdict(event) for event in session.diagnostics.events
               if event.stage in ('com.Close(False)', 'com.Quit', 'com.process.exit')]
    assert sum(e['stage'] == 'com.process.exit' and e['status'] == 'OK' for e in cleaned) == 4
    assert all(sha(p) == h for p, h in before.items())
    return dict(development_synthetic_only=True, originals_unchanged=True, attempts=audit,
                file_errors=session.errors, cleanup=cleaned, metrics=probe.result())


def b100_case(source):
    earlier = ROOT.parent / 'excel4-diagnostic'
    path = earlier / 'synthetic' / 'b_100.xlsx'
    expected = read_json(earlier / 'b_100_truth.json')
    before = sha(path)
    with Probe() as probe:
        session = compare(source, [path], probe)
    report, = diagnose_session(session)
    observed = [{key: var[key] for key in expected[0]} for sheet in report['sheets']
                for var in sheet['variables'] if var['extracted']]
    assert observed == expected and len(expected) == 18075
    assert report['counts']['function_candidates'] == 100 and not session.errors
    assert sha(path) == before
    return dict(function_sheets=100, variables=len(expected), wrong_links=0,
                originals_unchanged=True, metrics=probe.result())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    parser.add_argument('part', choices=('reader', 'writer'))
    parser.add_argument('--case')
    args = parser.parse_args()
    root = args.output.resolve()
    root.mkdir(parents=True, exist_ok=True)
    source = ROOT.parent / 'excel1-diagnostic' / 'synthetic' / 'excel1_types.h'
    cases = {
        'reader': [(f'budget_{n}', lambda n=n: boundary_case(source, root, n))
                   for n in (249999, 250000, 250001)] + [
            ('merge_budget', lambda: merge_budget_case(source, root)),
            ('column_edge', lambda: column_edge_case(source, root)),
            ('wide_long', lambda: limits_case(source, root)),
            ('sparse', lambda: sparse_budget_case(root)),
            ('open_audit', lambda: open_audit(source, root)),
            ('ownership', lambda: ownership_case(source, root)),
            ('baseline', lambda: baseline_case(source)),
            ('b100', lambda: b100_case(source)),
        ],
        'writer': [(f'write_{n}', lambda n=n: writer_scale(source, root, n)) for n in (1, 2, 5, 10)]
                  + [('writer_policy', lambda: writer_policy(source, root))],
    }
    failed = False
    selected = [(name, func) for name, func in cases[args.part] if not args.case or name == args.case]
    if not selected:
        parser.error(f'Unknown case {args.case!r}')
    for name, func in selected:
        start = perf_counter()
        try:
            result = func()
            result['outcome'] = 'SUCCESS'
        except Exception as exc:
            result = dict(outcome='FAILURE', error=repr(exc), traceback=traceback.format_exc())
            if isinstance(exc, ReadFailure):
                result['com_failure'] = exc.as_dict()
            failed = True
        result['elapsed_seconds'] = perf_counter() - start
        write_json(root / f'{name}.json', result)
        print(f'{name}: {result["outcome"]} {result["elapsed_seconds"]:.3f}s '
              f'{result.get("error", "")}', flush=True)
    return int(failed)


if __name__ == '__main__':
    raise SystemExit(main())
