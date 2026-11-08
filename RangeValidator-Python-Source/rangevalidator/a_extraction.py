"""User A (legacy engine B): header-only normalization and table observations.

No variable spelling, range classification, or source binding is changed here.
"""
from collections import defaultdict
import re
from .common import rows, at, text

KEYS = ('input/output', 'property', 'name', 'datatype', 'scope')


def normalize_header(value):
    return ''.join(str(value or '').split()).casefold()


def _safe(cell):
    if cell.is_formula or cell.kind == 'error':
        return False
    if not cell.merged_range:
        return True
    match = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', cell.merged_range)
    return bool(match and match[2] == match[4] and cell.merge_anchor == cell.address)


def scan_a(sheet):
    headers, tables, candidates = {}, [], []
    active, section = None, ''
    for number, cells in rows(sheet).items():
        labels = defaultdict(list)
        for cell in cells:
            label = normalize_header(cell.value)
            labels['scope' if label == '범위' else label].append(cell)
        # Distinctive A columns prevent altering sectioned Name/Type/Range tables.
        is_header = len(set(KEYS) & labels.keys()) >= 3 and bool(
            {'input/output', 'property'} & labels.keys())
        if is_header:
            reason = ''
            if any(len(labels[k]) > 1 for k in KEYS):
                reason = 'DUPLICATE_HEADER'
            elif any(not labels[k] for k in KEYS):
                reason = 'RANGE_CELL_UNRESOLVED' if not labels['scope'] else 'HEADER_MISSING'
            elif any(not _safe(labels[k][0]) for k in KEYS):
                reason = 'UNSAFE_HEADER_MERGE_OR_FORMULA'
            active = dict(id=f'{sheet.name}!row{number}', header_row=number, section=section,
                          columns={k: v[0] for k,v in labels.items() if k in KEYS and len(v)==1},
                          reason=reason)
            tables.append(dict(table_id=active['id'], section=section, header_row=number,
                headers={k:[c.address for c in labels[k]] for k in KEYS},
                normalized_headers={c.address:normalize_header(c.value) for c in cells},
                state='HEADER_FOUND' if not reason else reason))
            if not reason:
                headers[number] = active['columns']
            continue
        populated = [c for c in cells if text(c).strip()]
        plain_name = bool(len(populated)==1 and active and 'name' in active['columns']
                          and populated[0].column==active['columns']['name'].column
                          and re.fullmatch(r'[A-Za-z_]\w*',text(populated[0]).strip()))
        if len(populated)==1 and not plain_name and not re.fullmatch(r'[A-Za-z_]\w*(?:\[\w+\])?(?:\.|->)\w+',text(populated[0]).strip()):
            section = text(populated[0]).strip()
            active = None
            continue
        # Other header structures/sections end A ownership of following rows.
        if {'name','type'} <= labels.keys() or (populated and normalize_header(text(populated[0])) in (
                '함수설명','상세설명','functiondescription','algorithm','알고리즘')):
            active = None
            continue
        if active is None or 'name' not in active['columns']:
            continue
        columns = active['columns']
        nc = at(sheet,number,columns['name'].column)
        name = text(nc).strip()
        if not name and nc.merge_anchor:
            name = text(sheet.cells.get(nc.merge_anchor,nc)).strip()
        if not name:
            continue
        rc = at(sheet,number,columns['scope'].column) if 'scope' in columns and not active['reason'] else None
        candidates.append(dict(sheet=sheet.name, function=sheet.name.split('.',1)[-1].strip(),
            format='A', table_id=active['id'], section=active['section'], name=name,
            name_cell=nc.address, range_cell=rc.address if rc else None,
            existing_range=rc.value if rc else None, range_formula=rc.formula if rc else None,
            input_output=text(at(sheet,number,columns['input/output'].column)) if 'input/output' in columns else '',
            property=text(at(sheet,number,columns['property'].column)) if 'property' in columns else '',
            name_merge=nc.merged_range, scope_merge=rc.merged_range if rc else None,
            state='VARIABLE_EXTRACTED' if rc else 'RANGE_CELL_UNRESOLVED',
            reason=active['reason'], row=number))
    return headers, tables, candidates
