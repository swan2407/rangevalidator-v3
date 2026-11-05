from pathlib import Path
from email.parser import Parser
from hashlib import sha256
import json
import zipfile
root = Path(__file__).resolve().parents[1]
records = []
licenses = root / 'docs/dependency-licenses'
licenses.mkdir(exist_ok=True)
for wheel in sorted((root / 'wheelhouse').glob('*.whl')):
    with zipfile.ZipFile(wheel) as z:
        metadata = next(s for s in z.namelist() if s.endswith('.dist-info/METADATA'))
        m = Parser().parsestr(z.read(metadata).decode('utf-8'))
        license_files = [s for s in z.namelist() if ('license' in s.lower() or 'licence' in s.lower() or 'copying' in s.lower()) and not s.endswith('/')]
        saved = []
        for i, name in enumerate(license_files):
            out = licenses / (m['Name'] + '-' + str(i) + '-' + Path(name).name)
            out.write_bytes(z.read(name))
            saved.append(str(out.relative_to(root)))
        records.append(dict(name=m['Name'], version=m['Version'], license=m['License-Expression'] or m['License'],
                            classifiers=m.get_all('Classifier', []), wheel=wheel.name,
                            sha256=sha256(wheel.read_bytes()).hexdigest(), licenses=saved))
(root / 'docs/DEPENDENCIES.json').write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding='utf-8')
(root / 'wheelhouse/SHA256SUMS').write_text(''.join(f"{r['sha256']}  {r['wheel']}\n" for r in records), encoding='ascii')
print(json.dumps([{k: r[k] for k in ['name','version','license','licenses']} for r in records], ensure_ascii=False))

