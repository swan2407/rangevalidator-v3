"""V1 Common Parameter table and declaration rules, independent of Reader backend."""
from dataclasses import dataclass, field
import posixpath
import re
from .analysis import identifier
from .models import Cell, address


def text(c):
    return '' if c.value is None else str(c.value)


def norm(c):
    return ' '.join(text(c).split()).casefold()


def normalize_member_expression(value):
    """Ignore spacing around member operators and subscripts, retaining identifier boundaries."""
    value = re.sub(r'-\s*>', '->', value.strip())
    return re.sub(r'\s*(->|[.\[\]])\s*', r'\1', value)


def sheet_function_name(name):
    """Remove the conventional numeric prefix when it precedes a function name."""
    return re.sub(r'^\s*\d+\.\s*', '', name, count=1).strip() or name.strip()


def key(s):
    return posixpath.normpath(str(s).strip().replace('\\', '/')).lower() if str(s).strip() else ''


def rows(sheet):
    out = {}
    for c in sheet.cells.values():
        if c.value is not None or c.is_formula:
            out.setdefault(c.row, []).append(c)
    return {r: sorted(cs, key=lambda c: c.column) for r, cs in sorted(out.items())}


def at(sheet, row, col):
    a = address(row, col)
    return sheet.cells.get(a, Cell(row, col, a, None))


def unsafe(c):
    return c.is_formula or c.merged_range is not None or c.kind == 'error'


def safe_header(c):
    if c.is_formula or c.kind == 'error':
        return False
    if not c.merged_range:
        return True
    match = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', c.merged_range)
    return bool(match and match[2] == match[4] and c.merge_anchor == c.address)


@dataclass
class Declaration:
    source_file: str
    raw_name: str
    raw_type: str
    value: str
    name: str = ''
    type_name: str = ''
    tag: bool = False
    pointer: bool = False
    array: bool = False
    extent: str = ''
    type_qualifiers: list[str] = field(default_factory=list)
    pointer_qualifiers: list[str] = field(default_factory=list)
    workbook: str = ''
    sheet: str = ''
    row: int = 0
    cells: dict = field(default_factory=dict)
    reviews: list[str] = field(default_factory=list)
    global_block: bool = False


@dataclass
class Common:
    workbook: object
    declarations: list[Declaration] = field(default_factory=list)
    reviews: list[str] = field(default_factory=list)


def parse_declaration(file, name, typ, value=''):
    d = Declaration(file, name, typ, value)
    m = re.fullmatch(r'\s*([A-Za-z_]\w*)\s*(?:\[\s*([1-9][0-9]*)\s*\])?\s*', name, re.ASCII)
    if not m or not identifier(m[1]):
        d.reviews.append('Unsupported Common Name')
        d.global_block = True
        return d
    d.name, d.array, d.extent = m[1], bool(m[2]), m[2] or ''
    if not file.strip() or any(c in file for c in '\t\r\n'):
        d.global_block = True
        d.reviews.append('Missing/invalid File context')
    spelling = typ.strip()
    arr = re.search(r'\[\s*([1-9][0-9]*)\s*\]$', spelling)
    if arr:
        if d.array:
            d.reviews.append('Repeated array extent')
            return d
        d.array, d.extent = True, arr[1]
        spelling = spelling[:arr.start()].strip()
    d.pointer = '*' in spelling
    if spelling.count('*') > 1 or (d.pointer and d.array):
        d.reviews.append('Multiple pointer/array of pointers unsupported')
        return d
    left, _, right = spelling.partition('*')
    pending = False
    for w in left.split():
        if w in ('const', 'volatile'):
            if pending or w in d.type_qualifiers:
                d.reviews.append('Invalid/repeated qualifier')
                return d
            d.type_qualifiers.append(w)
        elif w == 'struct' and not d.type_name and not pending:
            d.tag = pending = True
        elif identifier(w) and not d.type_name:
            d.type_name, pending = w, False
        else:
            d.reviews.append('Unsupported/ambiguous Type')
            return d
    if not d.type_name or pending:
        d.reviews.append('Missing declared struct type; Value never infers identity')
    for w in right.split():
        if w not in ('const', 'volatile') or w in d.pointer_qualifiers:
            d.reviews.append('Unsupported pointer qualifier/declarator')
            return d
        d.pointer_qualifiers.append(w)
    return d


def parse_common(book):
    out = Common(book)
    for sheet in book.sheets:
        header, found, unrelated, section = None, False, False, ''
        table = rows(sheet)
        for row, cs in table.items():
            labels = {}
            for c in cs:
                labels.setdefault(norm(c), []).append(c)
            populated = [c for c in cs if text(c).strip() or c.is_formula]
            if len(populated) == 1:
                title = norm(populated[0])
                sections = [('macro','Macro'), ('매크로','Macro'), ('calibration','Calibration'), ('보정','Calibration'),
                            ('global','Global'), ('전역','Global'), ('struct','Struct'), ('구조체','Struct'), ('enum','Enum'), ('열거','Enum')]
                selected = next((s for token,s in sections if token in title), '')
                if selected:
                    section, header, unrelated = selected, None, True
                    continue
            keys = ('file','name','type','value')
            complete = all(k in labels for k in keys)
            if section in ('Macro','Calibration','Struct','Enum') and not complete:
                continue
            if complete:
                h = [labels[k][0] for k in keys]
                if any(len(labels[k]) != 1 for k in keys) or any(unsafe(c) or not isinstance(c.value,str) for c in h):
                    out.reviews.append(f'{sheet.name}:{row}: ambiguous Common header')
                    header = None
                else:
                    header, found, unrelated = h, True, False
                continue
            if 'name' in labels and 'type' in labels:
                if 'file' in labels or 'value' in labels:
                    out.reviews.append(f'{sheet.name}:{row}: incomplete declaration header')
                    unrelated = False
                else:
                    unrelated = True
                header = None
                continue
            if header is None:
                if not unrelated and len(populated) > 1:
                    out.reviews.append(f'{sheet.name}:{row}: uninterpreted row before header')
                continue
            required = [at(sheet,row,c.column) for c in header]
            if not any(text(c).strip() or c.is_formula for c in required):
                continue
            d = parse_declaration(*(text(c) for c in required))
            d.workbook, d.sheet, d.row = book.path, sheet.name, row
            d.cells = {c.address:c for c in cs + required}
            for i,c in enumerate(required):
                if unsafe(c) or (c.value is not None and not isinstance(c.value,str)):
                    d.reviews.append(f'Unsafe declaration cell {c.address}')
                    if i in (0,1):
                        d.global_block = True
            out.declarations.append(d)
        if table and not found:
            out.reviews.append(f'{sheet.name}: no unambiguous Common declaration table')
    if not out.declarations:
        out.reviews.append('No Common declarations extracted')
    return out
