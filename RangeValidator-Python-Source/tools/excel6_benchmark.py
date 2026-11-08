"""Matched COM read benchmark with correctness and merge API evidence."""
import argparse
from collections import Counter
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import random
from statistics import median
import sys
from time import perf_counter
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from diagnose_excel5 import read_json, sha, validate, write_json
from excel5_metrics import Probe, Trace
from excel5_synthetic import _book, _table
from rangevalidator import com_values
from rangevalidator.extraction_diagnostics import diagnose_session


class MergeRange:
    def __init__(self, actual, counts):
        self.actual, self.counts = actual, counts

    @property
    def MergeCells(self):
        self.counts['MergeCells'] += 1
        return self.actual.MergeCells

    @property
    def MergeArea(self):
        self.counts['MergeArea'] += 1
        return MergeRange(self.actual.MergeArea, self.counts)

    @property
    def Address(self):
        self.counts['Address'] += 1
        return self.actual.Address


class MergeCells:
    def __init__(self, actual, counts):
        self.actual, self.counts = actual, counts

    def Item(self, *args):
        self.counts['Cells.Item'] += 1
        return MergeRange(self.actual.Item(*args), self.counts)


class MergeSheet:
    def __init__(self, actual, counts):
        self.actual, self.counts = actual, counts

    @property
    def Cells(self):
        self.counts['Worksheet.Cells'] += 1
        return MergeCells(self.actual.Cells, self.counts)

    def Range(self, *args):
        self.counts['Worksheet.Range'] += 1
        cells = [arg.actual if isinstance(arg, MergeRange) else arg for arg in args]
        return MergeRange(self.actual.Range(*cells), self.counts)


class BenchmarkTrace(Trace):
    def emit(self, reader, stage, status='ENTER', code=None, **metadata):
        if reader == 'COM' and metadata.get('api') == 'com.Value2':
            if status == 'ENTER':
                self.started['value2'] = perf_counter()
            elif 'value2' in self.started:
                self.probe.times['value2_seconds'] += perf_counter() - self.started.pop('value2')
        return super().emit(reader, stage, status, code, **metadata)


def fixtures(root):
    existing = ROOT.parent / 'excel5-diagnostic' / 'synthetic'
    for label, forms, dense in (
        ('plain_a', ['A'], False), ('plain_b', ['B'], False),
        ('light_merge', ['A', 'B'] * 8, False), ('heavy_merge', ['A', 'B'] * 8, True),
    ):
        path = root / f'{label}.xlsx'
        truth_path = root / f'{label}_truth.json'
        if not path.exists():
            book = _book()
            truth = []
            random.Random(606).shuffle(forms)
            for i, fmt in enumerate(forms):
                sheet = book.create_sheet(f'{i+1}. Bench_{i}')
                sheet['A1'] = 'Synthetic function title'
                if dense:
                    sheet.merge_cells('A1:L2')
                _table(sheet, truth, path.name, fmt, 4, 2, 500 if len(forms) == 1 else 80,
                       merged=dense or (label == 'light_merge' and i == 0))
                if dense:
                    # Vertical merges spanning populated, blank, and block-boundary rows.
                    for r in range(8, 80, 5):
                        sheet.merge_cells(start_row=r, start_column=12, end_row=r+3, end_column=13)
                    sheet['N95'] = 'Right boundary'
            book.save(path)
            book.close()
            write_json(truth_path, truth)
        yield label, path, truth_path
    for n in (120, 200, 300):
        yield f'mixed_{n}', existing / f'mixed_{n}.xlsx', existing / f'mixed_{n}_truth.json'


def signature(session):
    data = {
        'rows': [asdict(r) for r in session.results],
        'coverage': session.coverage,
        'sheets': [(s.name, s.bounds, s.merged_ranges,
                    [asdict(c) for _, c in sorted(s.cells.items())])
                   for b in session.books.values() for s in b.sheets],
    }
    return hashlib.sha256(json.dumps(data, sort_keys=True, default=str).encode()).hexdigest()


def run(source, path, truth):
    counts = Counter()
    before = sha(path)
    start = perf_counter()
    with Probe() as probe:
        merge = com_values._merge_map

        def measured(ws, *args, **kwargs):
            return merge(MergeSheet(ws, counts), *args, **kwargs)

        with patch.object(com_values, '_merge_map', measured):
            from rangevalidator.workflow import run_comparison
            session = run_comparison(None, [source], [path], backend='com', diagnostics=BenchmarkTrace(probe))
    elapsed = perf_counter() - start
    assert not session.errors and sha(path) == before
    check = validate(truth, session, diagnose_session(session))
    assert not check['disagreements'] and not check['unexpected'], check
    assert not check['counts']['duplicates']
    result = probe.result()
    result.update(input_sha256=before, file_bytes=path.stat().st_size, total_seconds=elapsed,
                  sheet_timings=next(iter(session.excel1_timings.values())),
                  merge_dispatches=dict(counts), correctness=check['counts'],
                  evidence_sha256=signature(session), original_unchanged=True)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    parser.add_argument('mode', choices=('before', 'after'))
    parser.add_argument('--repeat', type=int, default=3)
    parser.add_argument('--case')
    args = parser.parse_args()
    output = args.output.resolve()
    root = output / 'synthetic'
    root.mkdir(parents=True, exist_ok=True)
    source = ROOT.parent / 'excel1-diagnostic' / 'synthetic' / 'excel1_types.h'
    for label, path, truth_path in fixtures(root):
        if args.case and label != args.case:
            continue
        truth = read_json(truth_path)
        baseline = read_json(output / f'{label}_before.json') if args.mode == 'after' else None
        samples = []
        for index in range(args.repeat):
            sample = run(source, path, truth)
            if baseline:
                assert all(sample[k] == baseline['runs'][0][k] for k in
                           ('input_sha256', 'evidence_sha256', 'correctness')), (label, sample)
            elif samples:
                assert sample['evidence_sha256'] == samples[0]['evidence_sha256']
            samples.append(sample)
            print(f'{args.mode} {label} {index+1}: {sample["total_seconds"]:.3f}s, '
                  f'merge {sample["seconds"]["merge_scan_seconds"]:.3f}s', flush=True)
            write_json(output / f'{label}_{args.mode}.json', dict(
                file=str(path), outcome='SUCCESS' if len(samples) == args.repeat else 'IN_PROGRESS',
                total_median_seconds=median(s['total_seconds'] for s in samples), runs=samples))


if __name__ == '__main__':
    main()
