"""Build a development onedir distribution using the invoking Windows CPython."""
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def main():
    if sys.platform != 'win32' or 'mingw' in sys.version.lower():
        raise RuntimeError('Use the verified Windows CPython environment.')
    (ROOT / 'build').mkdir(exist_ok=True)
    command = [sys.executable, '-m', 'PyInstaller', '--onedir', '--windowed',
               '--name', 'RangeValidator', '--distpath', str(ROOT / 'dist'),
               '--workpath', str(ROOT / 'build' / 'development'),
               '--specpath', str(ROOT / 'build'), str(ROOT / 'gui_entry.py')]
    with (ROOT / 'build' / 'development-build.log').open('wb') as log:
        result = subprocess.run(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    if result.returncode:
        raise RuntimeError('Build failed. See build/development-build.log.')
    print(ROOT / 'dist' / 'RangeValidator' / 'RangeValidator.exe')


if __name__ == '__main__':
    main()
