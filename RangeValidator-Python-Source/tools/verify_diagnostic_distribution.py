"""Extract and run diagnostic-only EXE checks. No engine/Writer/report integration."""
from pathlib import Path
import hashlib,json,os,subprocess,time,zipfile
ROOT=Path(__file__).resolve().parents[1]

def main():
    release=json.loads((ROOT/'docs/DISTRIBUTION.json').read_text(encoding='utf-8'))
    stage=ROOT/'build'/('diagnostic-isolated-'+str(time.time_ns()));stage.mkdir()
    with zipfile.ZipFile(release['zip']) as z:z.extractall(stage)
    package=stage/Path(release['exe']).parent.name
    for line in (package/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines():
        expected,name=line.split('  ',1)
        assert hashlib.sha256((package/name).read_bytes()).hexdigest()==expected,name
    env={k:v for k,v in os.environ.items() if not k.upper().startswith('PYTHON')}
    env['PATH']=os.environ['SystemRoot']+'/System32;'+os.environ['SystemRoot']
    empty=stage/'empty';empty.mkdir()
    result={}
    for mode in ('smoke-test','diagnostic-test'):
        receipt=stage/(mode+'.json')
        command=[str(package/'RangeValidator-Python.exe'),'--'+mode,str(receipt)]
        p=subprocess.run(command,cwd=empty,env=env,capture_output=True,timeout=240 if mode=='diagnostic-test' else 40)
        assert p.returncode==0,(mode,p.returncode,p.stderr)
        result[mode]=json.loads(receipt.read_text(encoding='utf-8'))
    for field in ('pythoncom','win32api'):
        assert Path(result['smoke-test'][field]).resolve().is_relative_to(package.resolve())
    probe=result['diagnostic-test']
    for field in ('gui_diagnostic_window','view_button','masked','reader_native_success','reader_com_success','reader_com_failure','last_result_retained','original_inputs_unchanged'):
        assert probe[field],field
    assert not probe['writer_executed'] and not probe['report_executed'] and not probe['actual_drm_tested']
    record=dict(diagnostic_only=True,full_suite_run=False,writer_executed=False,report_executed=False,
        windows_only_path=env['PATH'],python_environment_removed=True,zip_hashes_verified=True,
        exe_sha256=release['exe_sha256'],zip_sha256=release['zip_sha256'],source_sha256=release['source_sha256'],
        result=result,actual_drm_tested=False,clean_pc_tested=False)
    for name in ('DIAGNOSTIC_DISTRIBUTION_VALIDATION.json','DISTRIBUTION_VALIDATION.json'):
        (ROOT/'docs'/name).write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(record,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
