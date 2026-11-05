"""Finalize ONLY the diagnostic test/build/EXE probe sequence."""
from datetime import datetime,timezone
from pathlib import Path
import json,subprocess,sys
ROOT=Path(__file__).resolve().parents[1]

def execute(script):
    run=subprocess.run([sys.executable,str(ROOT/'tools'/script)],cwd=ROOT,capture_output=True)
    if run.returncode:
        print(run.stdout.decode('utf-8',errors='replace'));print(run.stderr.decode('utf-8',errors='replace'))
        raise SystemExit(run.returncode)
    print(script+': OK',flush=True)

execute('run_diagnostic_tests.py')
execute('check_diagnostic_scope.py')
history=ROOT/'build/release-history'/('diagnostic-final-'+datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f'));history.mkdir(parents=True)
for rel in ('dist/RangeValidator-Python-Windows-x64','dist/RangeValidator-Python-Windows-x64.zip','dist/SHA256SUMS.txt'):
    src=(ROOT/rel).resolve();dst=(history/src.name).resolve()
    assert src.is_relative_to(ROOT.resolve()) and dst.is_relative_to(ROOT.resolve())
    assert not dst.exists()
    src.rename(dst)
for name in ('DISTRIBUTION.json','DISTRIBUTION_VALIDATION.json'):
    src=ROOT/'docs'/name
    (history/name).write_bytes(src.read_bytes())
execute('package_windows.py')
execute('verify_diagnostic_distribution.py')
print(json.dumps(json.loads((ROOT/'docs/DISTRIBUTION.json').read_text(encoding='utf-8')),ensure_ascii=False,indent=2))
