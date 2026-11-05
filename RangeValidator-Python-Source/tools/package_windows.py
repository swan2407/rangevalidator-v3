"""Build onedir, collect notices, immutable ZIP/manifest verification."""
from pathlib import Path
import argparse,hashlib,json,os,re,shutil,subprocess,sys,zipfile
ROOT=Path(__file__).resolve().parents[1]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--package-name',default='RangeValidator-Python-Windows-x64')
    parser.add_argument('--zip-name',default='RangeValidator-Python-Windows-x64.zip')
    args=parser.parse_args()
    if any(not re.fullmatch(r'[A-Za-z0-9_.-]+',n) or n in ('.','..') for n in (args.package_name,args.zip_name)):raise ValueError('Invalid release name')
    dest=ROOT/'dist';dest.mkdir(exist_ok=True)
    package=dest/args.package_name
    archive=dest/args.zip_name
    if package.exists() or archive.exists():raise FileExistsError('Release already exists; preserve before building a new release')
    command=[sys.executable,'-m','PyInstaller','--noconfirm','--clean','--onedir','--windowed','--name','RangeValidator-Python',
      '--hidden-import','win32timezone','--hidden-import','win32com.client.dynamic','--hidden-import','win32com.client.gencache',
      '--hidden-import','rangevalidator.selfcheck','--distpath','build/frozen','--workpath','build/pyinstaller','--specpath','build','gui_entry.py']
    run=subprocess.run(command,cwd=ROOT,capture_output=True)
    (ROOT/'build/PYINSTALLER_BUILD.log').write_bytes(run.stdout+run.stderr)
    if run.returncode:raise RuntimeError('PyInstaller failed; see build/PYINSTALLER_BUILD.log')
    shutil.copytree(ROOT/'build/frozen/RangeValidator-Python',package)
    shutil.copytree(ROOT/'docs/dependency-licenses',package/'licenses')
    shutil.copy2(ROOT/'docs/DEPENDENCIES.json',package/'licenses/DEPENDENCIES.json')
    shutil.copy2(ROOT/'docs/WINDOWS_GUIDE_KO.md',package/'사용안내.md')
    rebuild_guide=ROOT/'docs/COM_READER_REBUILD.md'
    if rebuild_guide.exists():shutil.copy2(rebuild_guide,package/'Excel-COM-Reader-Guide.md')
    files=sorted(p for p in package.rglob('*') if p.is_file())
    checks=''.join(f'{sha(p)}  {p.relative_to(package).as_posix()}\n' for p in files)
    (package/'SHA256SUMS.txt').write_text(checks,encoding='utf-8')
    with zipfile.ZipFile(archive,'x',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for p in sorted(package.rglob('*')):
            if p.is_file():z.write(p,p.relative_to(dest).as_posix())
    with zipfile.ZipFile(archive) as z:
        expected={p.relative_to(dest).as_posix():sha(p) for p in package.rglob('*') if p.is_file()}
        actual={n:hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist()}
        if len(z.namelist())!=len(set(z.namelist())) or actual!=expected:raise RuntimeError('ZIP verification failed')
    exe=package/'RangeValidator-Python.exe'
    report=dict(exe=str(exe),exe_sha256=sha(exe),zip=str(archive),zip_sha256=sha(archive),files=len(expected),zip_verified=True,
        python=sys.version,source_sha256={str(p.relative_to(ROOT)):sha(p) for p in (ROOT/'rangevalidator').glob('*.py')},
        build_command=command,pyinstaller='6.22.3',architecture='Windows x64 onedir')
    (dest/'SHA256SUMS.txt').write_text(f'{report["zip_sha256"]}  {archive.name}\n{report["exe_sha256"]}  {package.name}/RangeValidator-Python.exe\n',encoding='utf-8')
    (ROOT/'docs/DISTRIBUTION.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
