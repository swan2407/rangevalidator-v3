import json
from collections import Counter
from pathlib import Path
import subprocess
import sys
root = Path(__file__).resolve().parents[1]
command = [sys.executable, '-m', 'rangevalidator', 'compare', str(root / 'tests/fixtures/synthetic_matching.h'), 'C:/Dev/RangeValidator/samples/synthetic_type_matching.xlsx']
run = subprocess.run(command, capture_output=True, encoding='utf-8', cwd=root)
if run.returncode:
    print(run.stderr)
    raise SystemExit(run.returncode)
data = json.loads(run.stdout)
print(json.dumps(dict(cli_compare='ok', rows=len(data['results']), source_reviews=len(data['source_reviews']), statuses=dict(Counter(r['status'] for r in data['results'])))))
