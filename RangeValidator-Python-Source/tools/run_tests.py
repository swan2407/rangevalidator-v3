"""Local validation runner; never discovers company input files."""
import argparse
import hashlib
from datetime import datetime, timezone
import io
import json
import os
from pathlib import Path
import platform
import sys
import time
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--excel', action='store_true', help='Run actual desktop Excel tests using synthetic workbooks')
    p.add_argument('--gui', action='store_true')
    args = p.parse_args()
    if args.gui:
        os.environ['RV_TEST_GUI'] = '1'
    if args.excel:
        os.environ['RV_TEST_EXCEL'] = '1'
    started = datetime.now(timezone.utc).isoformat()
    clock = time.perf_counter()
    suite = unittest.defaultTestLoader.discover(str(ROOT / 'tests'))
    stream = io.StringIO()
    result = unittest.TextTestRunner(stream=stream, verbosity=2).run(suite)
    receipt = dict(started_utc=started, python=sys.version, platform=platform.platform(),
                   excel_opt_in=os.environ.get('RV_TEST_EXCEL') == '1', tests=result.testsRun,
                   failures=len(result.failures), errors=len(result.errors), skipped=len(result.skipped),
                   elapsed_seconds=round(time.perf_counter() - clock, 3),
                   successful=result.wasSuccessful(),
                   skipped_tests=[dict(test=str(t), reason=why) for t, why in result.skipped],
                   source_sha256={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'rangevalidator').glob('*.py')})
    (ROOT / 'docs/TEST_RUN.log').write_text(stream.getvalue(), encoding='utf-8')
    (ROOT / 'docs/TEST_RUN.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(receipt, ensure_ascii=False, indent=2))
    if not result.wasSuccessful():
        print(stream.getvalue())
    return int(not result.wasSuccessful() or (args.excel and bool(result.skipped)))


if __name__ == '__main__':
    raise SystemExit(main())
