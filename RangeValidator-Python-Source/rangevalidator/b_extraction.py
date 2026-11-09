"""Discover user B tables (legacy engine format A) and their same-row cells."""
from collections import defaultdict
import re

from .a_extraction import normalize_header
from .common import at, rows, safe_header, text, sheet_function_name, normalize_member_expression


HEADERS = ('name', 'type', 'range')
SECTIONS = {
    normalize_header(title): title
    for title in ('Parameter(Input)', 'Parameter(Output)', 'Global(Input)', 'Global(Output)',
                  'Parameter', 'Global', 'Local', 'Return')
}
BOUNDARIES = {
    'functioninformation', 'functiondescription', 'algorithm', 'calibration',
    'calibrationvariables', '함수정보', '함수설명', '상세설명', '알고리즘',
    '캘리브레이션', '보정',
}
DECLARATION_NAME = re.compile(
    r'[A-Za-z_]\w*(?:\s*\[[^\]\r\n]*\])*'
    r'(?:\s*(?:\.|->)\s*[A-Za-z_]\w*(?:\s*\[[^\]\r\n]*\])*)*',
    re.ASCII,
)


def scan_b(sheet, table_rows=None):
    """Return safe header columns and diagnostics for complete and incomplete tables."""
    table_rows = rows(sheet) if table_rows is None else table_rows
    headers, tables, candidates = {}, [], []
    section, section_cell = '', None
    active, columns = None, {}
    calibration = False

    def finish(end_row, reason):
        nonlocal active
        if active is not None:
            active.update(end_row=end_row, end_reason=reason)
            active = None

    for number, cells in table_rows.items():
        populated = [cell for cell in cells if text(cell).strip() or cell.is_formula]
        if not populated:
            continue
        first = normalize_header(populated[0].value)
        if len(populated) == 1 and first in SECTIONS and safe_header(populated[0]):
            finish(number - 1, 'NEXT_SECTION')
            section, section_cell = SECTIONS[first], populated[0].address
            calibration = False
            continue

        labels = defaultdict(list)
        for cell in populated:
            labels[normalize_header(cell.value)].append(cell)
        if first in BOUNDARIES or ({'name', 'type', 'value'} <= labels.keys() and 'range' not in labels):
            finish(number - 1, 'SECTION_BOUNDARY')
            section, section_cell = '', None
            calibration = first in ('calibration', 'calibrationvariables', '캘리브레이션', '보정')
            continue
        if calibration:
            continue

        # The Korean alias needs a standalone B section and the other B headers.
        if section_cell and {'name', 'type'} <= labels.keys():
            labels['range'].extend(labels.get('범위', ()))
        header_keys = {key for key in HEADERS if labels.get(key)}
        if len(header_keys) >= 2:
            finish(number - 1, 'NEXT_HEADER')
            columns = {key: labels[key][0] for key in HEADERS if len(labels.get(key, ())) == 1}
            if any(len(labels.get(key, ())) > 1 for key in HEADERS):
                reason = 'DUPLICATE_HEADER'
            elif 'range' not in columns:
                reason = 'RANGE_CELL_UNRESOLVED'
            elif len(columns) != len(HEADERS):
                reason = 'HEADER_MISSING'
            elif not all(safe_header(cell) for cell in columns.values()):
                reason = 'UNSAFE_HEADER_MERGE_OR_FORMULA'
            else:
                reason = ''
                headers[number] = columns
            active = dict(
                format='B', table_id=f'{sheet.name}!row{number}', section=section,
                section_cell=section_cell, header_row=number,
                headers={key: [cell.address for cell in labels.get(key, ())] for key in HEADERS},
                normalized_headers={cell.address: normalize_header(cell.value) for cell in populated},
                start_row=None, last_variable_row=None, state=reason or 'HEADER_FOUND',
            )
            tables.append(active)
            continue

        if active is None or 'name' not in columns:
            # Unrelated content between a title and header cannot authorize an alias.
            if active is None:
                section, section_cell = '', None
            continue
        name_cell = at(sheet, number, columns['name'].column)
        name = text(name_cell).strip()
        if not name and name_cell.merge_anchor:
            name = text(sheet.cells.get(name_cell.merge_anchor, name_cell)).strip()
        if name in ('-', '--', 'N/A'):
            continue
        valid_name = DECLARATION_NAME.fullmatch(normalize_member_expression(name)) is not None
        standalone_label = (len(populated) == 1 and populated[0].address == name_cell.address
                            and re.fullmatch(r'[A-Za-z_]\w*', name))
        if not valid_name and name:
            # A populated Type or Range cell identifies an in-table row even
            # when its Name is not a variable (for example, a call expression).
            data_columns = {columns[key].column for key in ('type', 'range') if key in columns}
            if any(cell.column in data_columns for cell in populated):
                continue
        if standalone_label or not valid_name:
            finish(number - 1, 'UNCERTAIN_BOUNDARY')
            section, section_cell = '', None
            continue

        if active['start_row'] is None:
            active['start_row'] = number
        active['last_variable_row'] = number
        range_cell = at(sheet, number, columns['range'].column) if active['state'] == 'HEADER_FOUND' else None
        candidates.append(dict(
            sheet=sheet.name, function=sheet_function_name(sheet.name), format='B',
            section=section, section_cell=section_cell, table_id=active['table_id'],
            name=name, name_cell=name_cell.address, range_cell=range_cell.address if range_cell else None,
            existing_range=range_cell.value if range_cell else None,
            range_formula=range_cell.formula if range_cell else None,
            name_merge=name_cell.merged_range, range_merge=range_cell.merged_range if range_cell else None,
            state='VARIABLE_EXTRACTED' if range_cell else 'RANGE_CELL_UNRESOLVED',
            reason='' if range_cell else active['state'], row=number,
        ))

    finish(max(table_rows, default=0), 'END_OF_SHEET')
    return headers, tables, candidates
