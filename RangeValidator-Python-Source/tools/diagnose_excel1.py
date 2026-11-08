"""Excel-1 diagnostic runner. Produces only synthetic workbooks and JSON evidence."""
import argparse
import json
from pathlib import Path
import sys
from time import perf_counter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rangevalidator.extraction_diagnostics import diagnose_session, diagnose_workbook
from rangevalidator.readers import read_openpyxl
from rangevalidator.workflow import run_comparison
from excel1_synthetic import assess, build_suite


def write(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--backend", choices=("openpyxl", "com"), default="openpyxl")
    parser.add_argument("--samples", type=Path, help="directory with read-only A형.xlsx/B형.xlsx")
    args = parser.parse_args(argv)
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    source, paths, truth_path = build_suite(output / "synthetic")
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    reports = []
    timings = []
    for path in paths:
        started = perf_counter()
        session = run_comparison(None, [source], [path], backend=args.backend)
        report = diagnose_session(session)[0]
        elapsed = perf_counter() - started
        report["workflow_timings"] = session.excel1_timings.get(str(path.resolve()), {})
        report["wall_seconds"] = elapsed
        reports.append(report)
        timings.append(dict(file=path.name, wall_seconds=elapsed,
                            **report["workflow_timings"],
                            diagnostic_scan_seconds=report["timings"].get("diagnostic_scan_seconds")))
        print(f"{path.name}: {len([v for s in report['sheets'] for v in s['variables'] if v['extracted']])} extracted, {elapsed:.3f}s", flush=True)
    assessment = assess(truth, reports)
    write(output / "diagnostics.json", reports)
    write(output / "assessment.json", assessment)
    write(output / "performance.json", timings)
    if args.samples:
        samples = []
        for name in ("A형.xlsx", "B형.xlsx"):
            path = args.samples / name
            if path.is_file():
                samples.append(diagnose_workbook(read_openpyxl(path)))
        write(output / "sample_observations.json", samples)
    print(json.dumps({k: v for k, v in assessment.items() if not isinstance(v, list)},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
