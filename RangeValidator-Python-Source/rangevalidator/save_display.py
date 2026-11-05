"""Presentation shared by the GUI and XLSX report."""
from collections import Counter
from pathlib import Path

from .manual_edits import AUTO_MISMATCH, AUTO_CREATE, MANUAL_REVIEW

SAVE_LABELS = {
    'SAVED_VERIFIED': '저장 및 검증 완료',
    'SAVED_UNVERIFIED': '저장됨 / 검증 보류',
    'SAVED_VERIFY_FAILED': '저장됨 / 값 검증 실패',
    'SAVE_FAILED': '저장 실패',
    'SKIPPED': '수정 제외',
}
CLEANUP_LABEL = 'Excel 종료·정리 경고'
WRITTEN_STATES = ('SAVED_VERIFIED', 'SAVED_UNVERIFIED', 'SAVED_VERIFY_FAILED')


def summarize_files(outcomes):
    groups = {}
    for outcome in outcomes.values():
        groups.setdefault(outcome['file'], []).append(outcome)
    for file, cells in groups.items():
        counts = Counter(cell['status'] for cell in cells)
        warnings = {f"{w.get('phase')}: {w.get('api') or w.get('stage')} ({w.get('code')})"
                    for cell in cells for w in cell.get('cleanup_warnings', ())}
        yield dict(file=file, state=cells[0]['file_status'], cells=cells, counts=counts,
                   written=sum(counts[state] for state in WRITTEN_STATES),
                   output=next((cell['output'] for cell in cells if cell.get('output')), ''),
                   warnings=sorted(warnings))


def gui_save_text(folder, outcomes):
    files = list(summarize_files(outcomes))
    states = Counter(item['state'] for item in files)
    counts = ' / '.join(f'{label} {states[state]}개' for state, label in SAVE_LABELS.items() if states[state])
    status = f'파일 저장 결과 — {counts} → {folder}'
    lines = [str(folder), status, '']
    for item in files:
        output = item['output']
        lines.append(f"{Path(item['file']).name} — {SAVE_LABELS[item['state']]}" + (f' → {output}' if output else ''))
        lines.append(f"  선택 후보 {len(item['cells'])}칸 / 저장 확인 {item['written']}칸")
        kinds = Counter(v.get('edit_kind', AUTO_MISMATCH) for v in item['cells'])
        if kinds[MANUAL_REVIEW] or kinds[AUTO_CREATE]:
            lines.append(f'  자동 MISMATCH {kinds[AUTO_MISMATCH]}칸 / CREATE 신규 입력 {kinds[AUTO_CREATE]}칸 / REVIEW 수동 {kinds[MANUAL_REVIEW]}칸')
        for state, label in SAVE_LABELS.items():
            if item['counts'][state]:
                lines.append(f"  {label} {item['counts'][state]}칸")
                lines.extend(f"    REVIEW 수동 {v.get('manual_value', '')} · {v['reason']}"
                             for v in item['cells'] if v['status'] == state
                             and v.get('edit_kind') == MANUAL_REVIEW)
                if state != 'SAVED_VERIFIED':
                    lines.extend(f"    {v.get('stage', '')} {v['reason']}"
                                 for v in item['cells'] if v['status'] == state and v.get('edit_kind') != MANUAL_REVIEW)
        lines.extend(f'  {CLEANUP_LABEL} — {warning}' for warning in item['warnings'])
    if getattr(outcomes, 'report_warning', ''):
        lines.append(outcomes.report_warning)
    return status, '\n'.join(lines)
