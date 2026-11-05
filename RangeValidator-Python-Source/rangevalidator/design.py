"""Conservative A/B header discovery and read-only interval comparison."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from .analysis import match_expression
from .models import Cell, address


@dataclass(frozen=True)
class DesignRow:
    file: str
    sheet: str
    row: int
    format: str
    name_cell: Cell
    type_cell: Cell
    range_cell: Cell
    reason: str = ''


@dataclass(frozen=True)
class Comparison:
    document: DesignRow
    status: str
    reason: str
    expected: tuple[int, int] | None = None
    evidence: object = None


def discover(workbook):
    rows = []
    for sheet in workbook.sheets:
        by_row = {}
        for c in sheet.cells.values():
            by_row.setdefault(c.row, []).append(c)
        header, fmt = None, None
        for rowno in sorted(by_row):
            cells = by_row[rowno]
            labels = {}
            for c in cells:
                if isinstance(c.value, str):
                    labels.setdefault(c.value.strip().casefold(), []).append(c)
            found = []
            for f, typ, rng in [('A', 'type', 'range'), ('B', 'data type', 'scope')]:
                keys = ('name', typ, rng)
                if all(k in labels for k in keys):
                    found.append((f, [labels[k] for k in keys]))
            if found:
                if len(found) != 1 or any(len(group) != 1 for group in found[0][1]):
                    header, fmt = None, None
                    # No arbitrary choice of duplicate header coordinates.
                    rows.append(DesignRow(workbook.path, sheet.name, rowno, '?', cells[0], cells[0], cells[0], 'Ambiguous headers'))
                else:
                    fmt, groups = found[0]
                    header = [group[0] for group in groups]
                continue
            # A partial/unrelated Name table terminates the preceding block.
            if 'name' in labels and any(k in labels for k in ('type', 'data type', 'value', 'file')):
                header, fmt = None, None
                continue
            if header is None:
                continue
            selected = [sheet.cells.get(address(rowno, h.column), Cell(rowno, h.column, address(rowno, h.column), None)) for h in header]
            if selected[0].value is None:
                continue
            unsafe = any(c.is_formula or c.merged_range or c.kind == 'error' for c in selected + header)
            rows.append(DesignRow(workbook.path, sheet.name, rowno, fmt, *selected,
                                  'Formula/merged/error data or header' if unsafe else ''))
    return rows


NUMBER = r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?'
PAIR = re.compile(rf'\s*({NUMBER})\s*(?:~|\.\.)\s*({NUMBER})\s*')
SCALAR = re.compile(rf'\s*{NUMBER}\s*')


def compare(model, row):
    if row.reason:
        return Comparison(row, 'REVIEW', row.reason)
    name = row.name_cell.value
    if name == '-':
        return Comparison(row, 'SKIP', 'Placeholder')
    if not isinstance(name, str):
        return Comparison(row, 'REVIEW', 'Non-text expression')
    matched = match_expression(model, name)
    if matched.status != 'TYPE_MATCH':
        return Comparison(row, matched.status, matched.reason)
    f = matched.field
    # Confirm document Type rather than ignoring a contradictory type cell.
    spellings = {f.declared_type}
    if f.declared_type in ('unsigned', 'unsigned int', 'uint'):
        spellings |= {'unsigned', 'unsigned int', 'uint'}
    if f.declared_type in ('signed', 'signed int'):
        spellings |= {'signed', 'signed int'}
    if not isinstance(row.type_cell.value, str) or ' '.join(row.type_cell.value.split()) not in spellings:
        return Comparison(row, 'REVIEW', 'Document type unconfirmed/contradictory', evidence=matched)
    expected = (f.minimum, f.maximum)
    value = row.range_cell.value
    text = '' if value is None else str(value)
    m = PAIR.fullmatch(text)
    try:
        if m and (Decimal(m[1]), Decimal(m[2])) == expected:
            return Comparison(row, 'MATCH', 'Numeric interval agrees', expected, matched)
    except InvalidOperation:
        pass
    if m or not text.strip() or SCALAR.fullmatch(text):
        return Comparison(row, 'AUTO_FIX_CANDIDATE', 'Certified range differs; read-only proposal', expected, matched)
    return Comparison(row, 'REVIEW', 'Unknown/free-text range', expected, matched)
