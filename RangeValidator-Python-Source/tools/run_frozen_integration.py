"""Launch an explicit, actual-Excel frozen acceptance test."""
import argparse
from datetime import datetime
from pathlib import Path
import subprocess


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', action='store_true', required=True)
    parser.add_argument('--exe', type=Path, default=Path('dist/RangeValidator/RangeValidator.exe'))
    args = parser.parse_args()
    receipt = Path('build') / ('acceptance-' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.json')
    receipt = receipt.resolve()
    receipt.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run([str(args.exe.resolve()), '--full-integration-test', str(receipt)],
                            cwd=Path(__file__).resolve().parents[2])
    print(receipt)
    raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
