"""Frozen ZIP read-only batch probe; no Writer/report, no complete regression."""
from pathlib import Path
import hashlib,json,os,subprocess,time,zipfile
ROOT=Path(__file__).resolve().parents[1]

def main():
    release=json.loads((ROOT/'docs/DISTRIBUTION.json').read_text(encoding='utf-8'))
    stage=ROOT/'build'/('batch-isolated-'+str(time.time_ns()));stage.mkdir()
    with zipfile.ZipFile(release['zip']) as z:z.extractall(stage)
    package=stage/Path(release['exe']).parent.name
    for line in (package/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines():
        digest,name=line.split('  ',1)
        assert hashlib.sha256((package/name).read_bytes()).hexdigest()==digest
    env={k:v for k,v in os.environ.items() if not k.upper().startswith('PYTHON')}
    env['PATH']=os.environ['SystemRoot']+'/System32;'+os.environ['SystemRoot']
    empty=stage/'empty';empty.mkdir();results={}
    for flag in ('smoke-test','batch-read-test'):
        receipt=stage/(flag+'.json')
        p=subprocess.run([str(package/'RangeValidator-Python.exe'),'--'+flag,str(receipt)],cwd=empty,env=env,capture_output=True,timeout=240)
        assert p.returncode==0,(flag,p.returncode,p.stderr)
        results[flag]=json.loads(receipt.read_text(encoding='utf-8'))
    probe=results['batch-read-test']
    assert probe['normal_bitfield_comparison'] and probe['file_error_separate_from_review'] and probe['originals_unchanged']
    assert not probe['writer_executed'] and not probe['report_executed']
    for key in ('pythoncom','win32api'):assert Path(results['smoke-test'][key]).resolve().is_relative_to(package.resolve())
    record=dict(full_suite_run=False,writer_report_tests_run=False,windows_only_path=env['PATH'],zip_hashes_verified=True,
        exe_sha256=release['exe_sha256'],zip_sha256=release['zip_sha256'],source_sha256=release['source_sha256'],
        results=results,actual_company_drm_tested=False,physical_clean_pc_tested=False)
    for name in ('DRM_BATCH_DISTRIBUTION_VALIDATION.json','DISTRIBUTION_VALIDATION.json'):
        (ROOT/'docs'/name).write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(dict(zip_hashes_verified=True,results=probe),ensure_ascii=False,indent=2))
if __name__=='__main__':main()
