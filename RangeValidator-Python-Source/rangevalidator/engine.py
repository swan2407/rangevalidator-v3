"""Production Common/multi-header/multi-document pipeline; V1 semantic adapter."""
from dataclasses import dataclass, field, replace
from pathlib import Path
import re
from .analysis import analyze, identifier
from .bitfield_catalog import BitfieldCatalog, build_bitfield_catalog
from .common import key, rows, at, text, norm
from .models import Cell


@dataclass
class Catalog:
    types: list = field(default_factory=list)
    reviews: list[str] = field(default_factory=list)       # informational (shown in 진단), never a global block
    variables: list = field(default_factory=list)
    typedefs: dict = field(default_factory=dict)
    macros: dict = field(default_factory=dict)
    enums: set = field(default_factory=set)
    files: list = field(default_factory=list)
    bitfields: BitfieldCatalog | None = None
    bitfield_members: frozenset[str] = frozenset()


@dataclass(frozen=True)
class HeaderVariable:
    name: str
    pointer: bool
    struct: object              # Struct or None (resolved lazily by type_text)
    file: str
    line: int
    array: bool = False
    type_text: str = ''


@dataclass(frozen=True)
class Result:
    id: str
    file: str
    sheet: str
    cell: str
    expression: str
    status: str
    reason: str
    existing: object = None
    expected: str = ''
    struct: str = ''
    member: str = ''
    width: int = 0
    signed: bool | None = None
    source: str = ''
    source_line: int = 0
    common_sheet: str = ''
    common_row: int = 0
    common_file: str = ''
    format: str = ''
    backend: str = ''
    candidate: bool = False


def _signature(st):
    return tuple((f.member, f.declared_type, f.width) for f in st.fields) + tuple(sorted(st.members.items()))


def catalog_headers(headers, encoding='utf-8-sig', diagnostics=None):
    """Analyze every C header/source once and merge into one catalog.
    Same struct identity (typedef name or tag) defined twice: identical → kept once, different → marked conflict."""
    from .analysis import Struct
    headers = tuple(headers)
    bitfields = build_bitfield_catalog(headers, encoding)
    members = frozenset(f.declaration.member for st in bitfields.definitions.values() for f in st.fields)
    catalog = Catalog(bitfields=bitfields, bitfield_members=members)
    models = []
    for path in headers:
        try:
            if diagnostics: diagnostics.emit('C','source.read')
            raw = Path(path).read_bytes()
            text = None
            for enc in (encoding, 'utf-8-sig', 'cp949', 'latin-1'):
                try:
                    text = raw.decode(enc); break
                except UnicodeDecodeError:
                    continue
            m = analyze(text or raw.decode('latin-1'), str(Path(path).resolve()))
            models.append(m)
            catalog.files.append(str(Path(path).resolve()))
            catalog.reviews.extend(f'{r.code}: {Path(r.file).name}:{r.line}: {r.reason}' for r in m.reviews)
        except Exception as exc:
            if diagnostics: diagnostics.emit('C','source.read','FAILED',getattr(exc,'winerror',getattr(exc,'errno',None)))
            catalog.reviews.append(f'SOURCE_READ_ERROR: {Path(path).name}: {type(exc).__name__}')
    if not headers:
        catalog.reviews.append('No headers supplied')
    # merge types (dedupe identical duplicates, flag conflicts)
    by_key = {}
    index_map = []                                   # per model: local index -> catalog index
    for m in models:
        local = []
        for st in m.types:
            k = ('typedef', st.name) if st.typedef and not st.name.startswith('<') else ('tag', st.tag) if st.tag else ('anon', st.file, st.line)
            if k in by_key:
                ci = by_key[k]
                other = catalog.types[ci]
                if _signature(other) != _signature(st):
                    msg = f'{Path(other.file).name}:{other.line} / {Path(st.file).name}:{st.line}'
                    other.conflict = msg
                    st.conflict = msg
                    catalog.reviews.append(f'TYPE_CONFLICT: {st.name or st.tag}: {msg}')
                local.append(ci)
                continue
            catalog.types.append(st)
            by_key[k] = len(catalog.types) - 1
            local.append(len(catalog.types) - 1)
        index_map.append(local)
        for alias, spec in m.typedefs.items():
            if spec.startswith('<pointer-to:'):
                spec = '<pointer-to:%d>' % local[int(spec[len('<pointer-to:'):-1])]
            catalog.typedefs.setdefault(alias, spec)
        for k, v in m.macros.items():
            catalog.macros.setdefault(k, v)
        catalog.enums |= m.enums
    for m, local in zip(models, index_map):
        for v in m.variables:
            st = catalog.types[local[v.type_index]] if v.type_index >= 0 else None
            catalog.variables.append(HeaderVariable(v.name, v.pointer, st, v.file, v.line, v.array, v.type_text))
    # lazily resolve variables whose type lives in another file
    from .analysis import struct_index
    fixed = []
    for hv in catalog.variables:
        if hv.struct is None and hv.type_text:
            hit = struct_index(hv.type_text, catalog.types, catalog.typedefs)
            if hit is not None:
                hv = HeaderVariable(hv.name, hv.pointer or hit[1], catalog.types[hit[0]], hv.file, hv.line, hv.array, hv.type_text)
        fixed.append(hv)
    catalog.variables = fixed
    return catalog


def source_context(path):
    p = Path(path)
    stem = p.stem
    for suffix in ('_DetailDesign', '_Detail'):
        if stem.endswith(suffix) and len(stem) > len(suffix):
            stem = stem[:-len(suffix)]
            break
    m = re.search(r'\(([^()]+\.c)\)', p.name, re.IGNORECASE)      # 'BMS … Specification(canfd_comm.c).xlsx'
    if m:
        stem = m.group(1)
    if not stem.lower().endswith('.c'):
        stem += '.c'
    return key(str(p.parent / stem))


def safe_header(c):
    if c.is_formula or c.kind == 'error':
        return False
    if not c.merged_range:
        return True
    m = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', c.merged_range)
    return bool(m and m[2] == m[4] and c.merge_anchor == c.address)


EXPR_RE = re.compile(r'\s*([A-Za-z_]\w*)\s*((?:\[[^\]]*\]\s*)*)((?:\s*(?:->|\.)\s*[A-Za-z_]\w*\s*(?:\[[^\]]*\]\s*)*)+)\s*\Z', re.ASCII)


def parse_expression(expression):
    """'a.b', 'p->b', 'arr[i].b', 'a.sub.b' → (base, [(op, member), ...]) or None"""
    m = EXPR_RE.fullmatch(expression or '')
    if not m or not identifier(m[1]):
        return None
    path = re.findall(r'(->|\.)\s*([A-Za-z_]\w*)', m[3])
    if not path or any(not identifier(name) for _, name in path):
        return None
    return m[1], bool(m[2].strip()), path


def _base_struct(catalog, base):
    """→ (Struct, declaration-or-None, note) or (None, decl, reason)"""
    # C sources: variable declarations (extern or definition)
    hvs = [v for v in catalog.variables if v.name == base and v.struct is not None]
    structs = {id(v.struct): v.struct for v in hvs}
    if len(structs) == 1:
        v = hvs[0]
        return v.struct, None, f'C 선언 {Path(v.file).name}:{v.line} ({v.type_text})'
    if len(structs) > 1:
        names = sorted({f'{v.struct.name or v.struct.tag}@{Path(v.file).name}' for v in hvs})
        return None, None, f"변수 '{base}' 가 서로 다른 타입으로 선언됨: {', '.join(names)}"
    return None, None, f"'{base}' 의 구조체 타입을 찾지 못함 (C/H 선언 없음)"


def resolve(common, catalog, expression, context):
    """→ (status, reason, BitField|None, Declaration|None)"""
    from .analysis import struct_index
    parsed = parse_expression(expression)
    note = ''
    if not parsed:
        found = re.findall(r'[A-Za-z_]\w*(?:\[[^\]]*\])?(?:(?:->|\.)[A-Za-z_]\w*(?:\[[^\]]*\])?)+', expression or '')
        if len(found) == 1:
            parsed = parse_expression(found[0])
            note = f" (칸 글 '{expression.strip()[:40]}' 에서 '{found[0]}' 추출)"
        if not parsed:
            return 'REVIEW', '변수 표기 형식 아님 (obj.member / ptr->member / arr[i].member)', None, None
    base, _array, path = parsed
    st, d, how = _base_struct(catalog, base)
    if st is None:
        return 'REVIEW', how, None, d
    if st.conflict:
        return 'REVIEW', f"구조체 '{st.name or st.tag}' 정의가 파일마다 다름: {st.conflict}", None, d
    # walk nested members
    for op, member in path[:-1]:
        nxt = st.nested.get(member)
        if nxt is None and member in st.members:
            hit = struct_index(st.members[member].replace(' *', '').replace('[]', ''), catalog.types, catalog.typedefs)
            nxt = catalog.types[hit[0]] if hit is not None else None
        if nxt is None:
            return 'REVIEW', f"'{member}' 은 '{st.name or st.tag}' 의 구조체 멤버가 아님", None, d
        st = nxt
    last = path[-1][1]
    fs = [f for f in st.fields if f.member == last]
    if not fs:
        if last in st.ordinary or last in st.members:
            return 'SKIP', f"'{last}' 은 비트필드가 아닌 일반 멤버 ({st.members.get(last, '')}) → 범위 계산 대상 아님", None, d
        return 'REVIEW', f"'{last}' 멤버가 '{st.name or st.tag}' 에 없음", None, d
    if len(fs) > 1:
        return 'REVIEW', f"'{last}' 멤버가 중복 정의됨", None, d
    f = fs[0]
    if f.minimum is None:
        return 'REVIEW', f"비트폭 {f.width or '?'} ({f.declared_type}): {f.note or '범위 계산 불가'}", f, d
    return 'TYPE_MATCH', how + note, f, d


def is_blank_range(value):
    return value is None or (isinstance(value, str) and not value.strip())


def parse_interval(value):
    s = str(value).strip()
    p = re.fullmatch(r'\[?\s*([+-]?[0-9]+)\s*(?:~|\.\.|to|,|–|—)\s*([+-]?[0-9]+)\s*\]?', s, re.IGNORECASE)
    if p:
        return int(p[1]), int(p[2])
    return None


def interval_status(value, expected):
    """Existing range cell vs expected 'lo~hi'. Accepts 'lo~hi', 'lo ~ hi', 'lo..hi', 'lo to hi', '[lo, hi]', 'lo – hi'."""
    s = '' if value is None else str(value).strip()
    if not s:
        return 'MISMATCH', 'FILLED_EMPTY_RANGE'
    bounds = parse_interval(s)
    if bounds:
        actual = f'{bounds[0]}~{bounds[1]}'
        return ('MATCH', '기존 범위 = 계산 범위') if actual == expected else ('MISMATCH', 'REPLACED_INVALID_RANGE')
    if re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?', s):
        return 'MISMATCH', 'REPLACED_INVALID_RANGE'
    return 'REVIEW', f"기존 범위 '{s[:40]}' 가 정수 구간 표기(lo~hi)가 아님 → 직접 확인"


def how_text(f):
    return f"{f.declared_type} {f.member}:{f.width} ({Path(f.file).name}:{f.line})" if f else ''


MEMBER_ACCESS = re.compile(
    r'(?P<base>[A-Za-z_]\w*)\s*(?:\[\s*(?P<index>\w+)\s*\]\s*)?'
    r'(?P<operator>\.|->)\s*(?P<member>[A-Za-z_]\w*)', re.ASCII)


def _input_source(bitfields, hint):
    if bitfields is None or not hint:
        return None
    normalized = key(hint)
    files = bitfields.source_scopes
    for matches in (
        [file for file in files if key(file) == normalized],
        [file for file in files if key(file).endswith('/' + normalized)],
    ):
        if matches:
            return matches[0] if len(matches) == 1 else None
    matches = [file for file in files if Path(key(file)).name == Path(normalized).name]
    return matches[0] if len(matches) == 1 else None


def _catalog_match(catalog, access, source_file, function):
    bitfields = catalog.bitfields
    name, member = access['base'], access['member']
    indexed, arrow = access['index'] is not None, access['operator'] == '->'

    def supports_access(depth, dimensions):
        if indexed:
            if dimensions:
                dimensions -= 1
            elif depth:
                depth -= 1
            else:
                return False
        return dimensions == 0 and depth == int(arrow)

    if function.status not in ('MISSING', 'RESOLVED'):
        return 'REVIEW', f'함수 정의를 확정할 수 없음 ({function.status})', None
    if function.status == 'RESOLVED':
        source_file = source_file or function.file
    binding = function.variables.get(name) or bitfields.lookup_variable(source_file, name)
    if binding.status == 'MISSING':
        return 'REVIEW', f"'{name}' 의 구조체 타입을 확인할 C/H 선언이 없음", None
    if binding.status in ('INTEGER', 'SCALAR'):
        return 'SKIP', '', None
    if binding.status != 'STRUCT':
        return 'REVIEW', f"'{name}' 의 변수 선언 연결이 불확실함 ({binding.status})", None
    link = binding.type_link
    first = binding.declarations[0]
    reason = f'C 선언 {Path(first.origin.file).name}:{first.origin.line} ({first.spelling})'
    dimensions = len(first.arrays)
    decay = int(first.kind == 'PARAMETER' and dimensions == 1)
    compatible = supports_access(link.pointer_depth + decay, dimensions - decay)
    definition = bitfields.definitions[link.definition]
    if definition.status != 'RESOLVED':
        return 'REVIEW', f"'{definition.name}' 구조체 정의가 불확실함 ({definition.status})", None
    fields = [f for f in definition.fields if f.declaration.member == member]
    if not fields:
        return 'SKIP', '', None
    # An ordinary or absent member is excluded before validating the operator.
    if not compatible:
        return 'REVIEW', f"'{name}' 의 배열/포인터 선언과 접근 형태가 다름", None
    f = fields[0]
    if len(fields) != 1 or f.status != 'RESOLVED':
        return 'REVIEW', f"'{member}' 의 비트필드 범위를 확정할 수 없음 ({f.status})", f.declaration
    return 'TYPE_MATCH', reason, f.declaration


def is_function_sheet(name):
    return bool(re.fullmatch(r'\s*\d+\.\s*\S.*', name))


def compare_book(common, catalog, book, relative_path=None):
    """Compare with C/H evidence. The legacy common argument is ignored."""
    results, coverage = [], []
    filename = source_context(relative_path or Path(book.path).name)
    for sheet in book.sheets:
        if not is_function_sheet(sheet.name):
            continue
        table = rows(sheet)
        headers, issues = {}, []
        for row, cs in table.items():
            labels = {}
            for c in cs:
                labels.setdefault(norm(c), []).append(c)
            found = []
            for fmt, keys in [('A',('name','type','range')), ('B',('input/output','property','name','data type','scope'))]:
                if all(k in labels for k in keys):
                    found.append((fmt,keys))
            if found:
                if len(found) != 1:
                    issues.append('Conflicting A/B header')
                    continue
                fmt,keys = found[0]
                if any(len(labels[k]) != 1 or not safe_header(labels[k][0]) for k in keys):
                    issues.append('Duplicate/unsafe header')
                    continue
                names = ('name','type','range') if fmt == 'A' else ('name','data type','scope')
                headers[row] = (fmt,[labels[k][0] for k in names],[labels[k][0] for k in ('input/output','property')] if fmt == 'B' else [])
            else:
                potential = len({'name','type','range'} & set(labels)) >= 2 or len({'input/output','property','name','data type','scope'} & set(labels)) >= 4
                if potential and any(c.is_formula or c.kind == 'error' for c in cs):
                    issues.append('Partial design header contains formula/error')
        if len({v[0] for v in headers.values()}) > 1:
            issues.append('Mixed A/B tables within sheet')
        if issues or (not headers and table and sheet.name.casefold() not in ('cover','contents','table of contents','표지','목차')):
            coverage.append((sheet.name,'REVIEW','; '.join(issues) or 'Unknown worksheet format'))
            continue
        if not headers:
            coverage.append((sheet.name,'SKIP','Cover/contents/empty sheet'))
            continue
        coverage.append((sheet.name,'TARGET','Design headers recognized'))
        source_file = _input_source(catalog.bitfields, filename)
        function_name = sheet.name.split('.', 1)[1].strip()
        function = catalog.bitfields.lookup_function(source_file, function_name) if catalog.bitfields else None
        header, section = None,''
        for row,cs in table.items():
            if row in headers:
                header = headers[row]
                continue
            populated = [c for c in cs if text(c).strip()]
            if len(populated) == 1:
                title = norm(populated[0])
                if re.fullmatch(r'(parameter|global|local|static|return|calibration|input|output|전역|지역|매개|파라미터|반환|캘리브레이션|보정)\w*(\s*(variable|variables|변수|값))?(\s*\((input|output|입력|출력)\))?',title) \
                        or parse_expression(text(populated[0]).strip()) is None:
                    section = title                                   # 섹션 제목 행 (Global Variable, Local Variable, …)
                    if title.startswith(('calibration','캘리브레이션','보정')):
                        header = None
                    continue
            if header is None or section == 'calibration':
                continue
            fmt,h,classification = header
            nc,tc,rc = [at(sheet,row,c.column) for c in h]
            if fmt == 'A' and (norm(nc),norm(tc),norm(rc)) == ('name','type','value'):
                header = None
                continue
            if not any(text(c).strip() or c.is_formula for c in (nc,tc,rc)):
                continue
            # Retain anchor text for merged expression evidence, but never certify merges.
            expr = text(nc).strip()
            if not expr and nc.merge_anchor:
                expr = text(sheet.cells.get(nc.merge_anchor,nc)).strip()
            access = MEMBER_ACCESS.fullmatch(expr)
            if not access or access['member'] not in catalog.bitfield_members:
                continue
            if not identifier(access['base']) or (access['index'] and not (
                    access['index'].isdigit() or identifier(access['index']))):
                continue
            status,reason,f = _catalog_match(catalog, access, source_file, function)
            if status == 'SKIP':
                continue
            rid = f'{book.path}|{sheet.name}|{rc.address}'
            base = dict(id=rid,file=book.path,sheet=sheet.name,cell=rc.address,expression=expr,existing=rc.formula if rc.is_formula else rc.value,format=fmt,backend=book.backend)
            if f:
                base.update(struct=f.struct_name,member=f.member,width=f.width,signed=f.signed,source=f.file,source_line=f.line,
                            expected=f'{f.minimum}~{f.maximum}' if f.minimum is not None else '')
                design_type=text(tc).strip()
                if design_type and f.declared_type and design_type.replace(' ','').lower()!=f.declared_type.replace(' ','').lower():
                    reason=f"{reason} / 설계서 Type '{design_type}' ≠ 코드 '{f.declared_type}'"
            if status == 'TYPE_MATCH':
                if rc.is_formula or rc.kind == 'error':
                    status,reason = 'REVIEW','범위 칸에 수식/오류 값이 있어 자동 수정하지 않음'
                elif rc.merged_range and rc.merge_anchor != rc.address:
                    status,reason = 'REVIEW','범위 칸이 병합 영역의 안쪽 칸 → 직접 확인'
                elif is_blank_range(rc.value):
                    incomplete = book.partial and (not book.read_errors or any(
                        not error.get('sheet_name') or error['sheet_name'].casefold() == sheet.name.casefold()
                        for error in book.read_errors))
                    if incomplete or rc.kind != 'value':
                        status,reason = 'REVIEW','공란 범위 셀의 읽기 근거가 불완전함'
                    else:
                        status,reason = 'CREATE','확정 비트필드의 공란 범위 신규 입력'
                else:
                    status,reason = interval_status(rc.value,base['expected'])
                    reason = reason + ' / ' + how_text(f) if status in ('MATCH','MISMATCH','CREATE') else reason
            results.append(Result(**base,status=status,reason=reason,candidate=status in ('MISMATCH','CREATE')))
    # Duplicate target coordinates are never independent edits.
    counts = {}
    for r in results:
        k = (r.sheet,r.cell)
        counts[k] = counts.get(k,0) + 1
    return [replace(r,status='REVIEW',reason='Duplicate target cell',candidate=False) if counts[(r.sheet,r.cell)] > 1 else r for r in results], coverage
