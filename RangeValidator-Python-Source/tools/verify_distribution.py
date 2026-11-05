"""ZIP extraction/hash + GUI/COM run with Windows-only PATH, no Python env vars."""
from pathlib import Path
import hashlib,json,os,subprocess,sys,time,zipfile
ROOT=Path(__file__).resolve().parents[1]


def main():
 report=json.loads((ROOT/'docs/DISTRIBUTION.json').read_text(encoding='utf-8'))
 stage=ROOT/'build'/('isolated-'+str(time.time_ns()));stage.mkdir()
 extracted=stage/'extract';extracted.mkdir()
 with zipfile.ZipFile(report['zip']) as z:z.extractall(extracted)
 package=extracted/'RangeValidator-Python-Windows-x64'
 for line in (package/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines():
  digest,name=line.split('  ',1)
  if hashlib.sha256((package/name).read_bytes()).hexdigest()!=digest:raise RuntimeError('Extracted checksum mismatch')
 empty=stage/'empty';empty.mkdir()
 env={k:v for k,v in os.environ.items() if not k.upper().startswith('PYTHON')}
 env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')+';'+os.environ['SystemRoot']
 exe=package/'RangeValidator-Python.exe'
 results={};modules=[]
 for mode in ('smoke-test','integration-test'):
  receipt=stage/(mode+'.json')
  p=subprocess.Popen([str(exe),'--'+mode,str(receipt)],cwd=empty,env=env,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  if mode=='smoke-test':
   try:
    import win32api,win32process
    time.sleep(.25)
    handle=win32api.OpenProcess(0x410,False,p.pid)
    try:modules=[win32process.GetModuleFileNameEx(handle,m) for m in win32process.EnumProcessModules(handle)]
    finally:handle.Close()
   except Exception as exc:modules=['module inspection unavailable: '+type(exc).__name__]
  stdout,stderr=p.communicate(timeout=180 if mode=='integration-test' else 40)
  if p.returncode or not receipt.exists():raise RuntimeError(f'{mode} failed, exit={p.returncode}, stderr={stderr!r}')
  value=json.loads(receipt.read_text(encoding='utf-8'))
  if 'error' in value:raise RuntimeError(value['error'])
  results[mode]=value
 for name in ('pythoncom','win32api'):
  value=Path(results['smoke-test'][name]).resolve()
  if not value.is_relative_to(package.resolve()):raise RuntimeError('External pywin32 loaded')
 prohibited=('c:\\msys64\\',str(ROOT/'.build-python').lower(),str(ROOT/'.runtime').lower())
 if any(any(m.lower().startswith(p) for p in prohibited) for m in modules):raise RuntimeError('External Python/MSYS2 module loaded')
 receipt=dict(path=env['PATH'],python_environment_removed=True,zip_extraction_hashes_verified=True,results=results,loaded_modules=modules,
  clean_pc_tested=False,actual_drm_tested=False,source_sha256=report['source_sha256'])
 (ROOT/'docs/DISTRIBUTION_VALIDATION.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps(dict(gui=results['smoke-test']['gui'],tabs=results['smoke-test']['tabs'],integration=results['integration-test'],zip_hashes=True,clean_pc_tested=False),ensure_ascii=False,indent=2))
if __name__=='__main__':main()
