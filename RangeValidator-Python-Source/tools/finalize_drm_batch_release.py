"""Related tests, immutable scope, onedir package, read-only frozen batch acceptance."""
from datetime import datetime,timezone
from pathlib import Path
import subprocess,sys,json
ROOT=Path(__file__).resolve().parents[1]

def run(name,*args):
    p=subprocess.run([sys.executable,str(ROOT/'tools'/name),*args],cwd=ROOT,capture_output=True)
    if p.returncode:
        print(p.stdout.decode('utf-8',errors='replace'));print(p.stderr.decode('utf-8',errors='replace'));raise SystemExit(p.returncode)
    print(name+': OK',flush=True)

run('run_drm_batch_tests.py');run('check_drm_batch_scope.py')
tests=json.loads((ROOT/'docs/DRM_BATCH_TEST_RUN.json').read_text(encoding='utf-8'))
assert tests['successful'] and tests['tests']==41 and not tests['skipped']
history=ROOT/'build/release-history'/('drm-batch-'+datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f'));history.mkdir(parents=True)
# Never move/terminate the existing onedir application. It may be open on this PC.
zip_name='RangeValidator-Python-Windows-x64.zip'
for rel in ('dist/RangeValidator-Python-Windows-x64.zip','dist/SHA256SUMS.txt'):
    src=(ROOT/rel).resolve();dst=(history/src.name).resolve()
    assert src.is_relative_to(ROOT.resolve()) and dst.is_relative_to(ROOT.resolve()) and not dst.exists()
    if src.exists():
        try:src.rename(dst)
        except PermissionError:
            (history/src.name).write_bytes(src.read_bytes())
            if src.suffix=='.zip':zip_name='RangeValidator-Python-Windows-x64-'+history.name+'.zip'
package_name='RangeValidator-Python-Windows-x64-'+datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')
for name in ('DISTRIBUTION.json','DISTRIBUTION_VALIDATION.json'):(history/name).write_bytes((ROOT/'docs'/name).read_bytes())
run('package_windows.py','--package-name',package_name,'--zip-name',zip_name);run('verify_drm_batch_distribution.py')
r=json.loads((ROOT/'docs/DISTRIBUTION.json').read_text(encoding='utf-8'))
print(json.dumps({k:r[k] for k in ('exe','exe_sha256','zip','zip_sha256','files')},indent=2))
