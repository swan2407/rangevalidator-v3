"""Immutable scope checks for this batch-read repair; never accesses V1 or V2."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
before=json.loads((ROOT/'build/drm-batch-before/sha256.json').read_text(encoding='utf-8'))
allowed={'rangevalidator\\workflow.py','rangevalidator\\readers.py','rangevalidator\\diagnostics.py','rangevalidator\\models.py','rangevalidator\\gui.py','gui_entry.py','docs\\DISTRIBUTION.json'}
changed=[name for name,h in before.items() if hashlib.sha256((ROOT/name).read_bytes()).hexdigest()!=h]
assert set(changed)<=allowed,changed
for name in ('analysis.py','engine.py','common.py','writer.py','report.py','preservation.py','cli.py'):
    key=str(Path('rangevalidator')/name)
    assert hashlib.sha256((ROOT/key).read_bytes()).hexdigest()==before[key],name
proof=dict(changed_preexisting_files=changed,parser_engine_common_writer_report_unchanged=True,v1_v2_accessed=False)
(ROOT/'docs/DRM_BATCH_SCOPE_CHECK.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(proof,ensure_ascii=False,indent=2))
