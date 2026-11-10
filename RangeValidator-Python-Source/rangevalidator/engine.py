"""Production Common/multi-header/multi-document pipeline; V1 semantic adapter."""
from dataclasses import dataclass, field, replace
from pathlib import Path
import re
from .analysis import analyze, identifier, matching_names
from .bitfield_catalog import BitfieldCatalog, build_bitfield_catalog
from .common import key, rows, at, text, norm, normalize_member_expression, safe_header, sheet_function_name
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


def catalog_headers(headers, encoding='utf-8-sig', diagnostics=None, build_options=None):
    """Analyze every C header/source once and merge into one catalog.
    Same struct identity (typedef name or tag) defined twice: identical → kept once, different → marked conflict."""
    from .analysis import Struct
    headers = tuple(headers)
    build_context = None
    original_headers = headers
    if build_options is not None:
        from .build_options import BuildOptions, load_build_options, preprocess_sources
        if not isinstance(build_options, BuildOptions):
            build_options = load_build_options(build_options)
        build_context = preprocess_sources(headers, build_options, encoding)
        headers = tuple(build_context.sources)
        bitfields = build_bitfield_catalog(headers, encoding, sources=build_context.sources, isolated=True)
    else:
        bitfields = build_bitfield_catalog(headers, encoding)
    members = frozenset(f.declaration.member for st in bitfields.definitions.values() for f in st.fields)
    catalog = Catalog(bitfields=bitfields, bitfield_members=members)
    catalog.build_context = build_context
    models = []
    for path in headers:
        try:
            if diagnostics: diagnostics.emit('C','source.read')
            if build_context is not None:
                text = build_context.sources[path]
            else:
                raw = Path(path).read_bytes()
                text = None
                for enc in (encoding, 'utf-8-sig', 'cp949', 'latin-1'):
                    try:
                        text = raw.decode(enc); break
                    except UnicodeDecodeError:
                        continue
                text = text or raw.decode('latin-1')
            m = analyze(text, str(Path(path).resolve()))
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
    if build_context is not None:
        catalog.files = [str(Path(p).resolve()) for p in original_headers]
        catalog.reviews.extend('BUILD_OPTIONS_UNRESOLVED: ' + message for message in build_context.errors.values())
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
    if s == "'-":
        s = '-'  # Some producers store Excel's leading text marker literally.
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


def _missing_declaration_reason(bitfields, name, source_file, function_name):
    locations = [(d.origin, '') for spelling in matching_names(bitfields.variables, name)
                 for d in bitfields.variables[spelling]]
    locations.extend(item for spelling in matching_names(bitfields.function_declaration_locations, name)
                     for item in bitfields.function_declaration_locations[spelling])
    if not locations:
        return f"DECL_NOT_COLLECTED: 등록된 C/H에서 '{name}' 선언을 수집하지 못함"

    context = f'함수 {function_name}' if function_name else '현재 함수'
    if source_file:
        context += f' / {Path(source_file).name}'
    locations = list(dict.fromkeys(locations))
    shown = ', '.join(f'{Path(origin.file).name}:{origin.line}' +
                      (f' ({function})' if function else '')
                      for origin, function in locations[:3])
    if len(locations) > 3:
        shown += f' 외 {len(locations) - 3}곳'
    return (f"DECL_OUT_OF_SCOPE: '{name}' 선언은 수집했으나 {context}에서 연결되지 않음 "
            f"({shown})")


def _catalog_match(catalog, access, source_file, function, *, function_name=None):
    bitfields = catalog.bitfields
    name, member = access['base'], access['member']
    if function.alternatives:
        scopes = function.alternatives
        local_scopes = [scope for scope in scopes if matching_names(scope.variables, name)]
        if local_scopes and len(local_scopes) != len(scopes):
            shared = bitfields.lookup_variable(source_file or function.file, name)
            if shared.status == 'MISSING' and bitfields.lookup_variable(None, name).status == 'MISSING':
                scopes = local_scopes
        matches = [_catalog_match(catalog, access, source_file, scope, function_name=function_name)
                   for scope in scopes]
        return _agree_member_matches(matches, name, member)
    if function.status not in ('MISSING', 'RESOLVED', 'CONDITIONAL'):
        return 'REVIEW', f'함수 정의를 확정할 수 없음 ({function.status})', None
    source_file = source_file or function.file or None
    binding = bitfields.lookup_variable(source_file, name, function.variables)
    if binding.status == 'MISSING' and getattr(catalog, 'build_context', None) is not None:
        from .build_options import member_scope
        return _member_catalog_match(member_scope(bitfields, source_file), access,
            _missing_declaration_reason(bitfields, name, source_file, function_name))
    if binding.status == 'MISSING' and source_file is not None:
        binding = bitfields.lookup_variable(None, name)
    if binding.status == 'MISSING':
        return _member_catalog_match(bitfields, access,
            _missing_declaration_reason(bitfields, name, source_file, function_name))

    match = _match_binding(bitfields, access, binding)
    if binding.declarations and function.variables.get(binding.declarations[0].name) is binding:
        if all(d.conditional for d in binding.declarations):
            fallback = bitfields.lookup_variable(source_file, name)
            if fallback.status != 'MISSING':
                return _agree_member_matches([match, _match_binding(bitfields, access, fallback)], name, member)
    return match


def _agree_member_matches(matches, name, member):
    """Combine the results for the same Excel row across source declarations."""
    first = matches[0]
    if all(match == first for match in matches):
        return first
    if all(status == 'TYPE_MATCH' for status, _, _ in matches):
        ranges = {(field.minimum, field.maximum) for _, _, field in matches}
        if len(ranges) == 1:
            return first[0], first[1] + ' / 선언 후보의 범위 동일', first[2]
    if all(status == 'SKIP' and reason for status, reason, _ in matches):
        return first[0], first[1] + ' / 모든 선언 후보에서 일반 멤버', None
    details = []
    for status, reason, field in matches:
        if status == 'TYPE_MATCH':
            details.append(f'{field.minimum}~{field.maximum} ({Path(field.file).name}:{field.line})')
        else:
            details.append(reason or '구조체 멤버가 아님')
    shown = ' / '.join(list(dict.fromkeys(details))[:5])
    return 'REVIEW', f"'{name}.{member}'의 선언 후보별 비교 결과가 다름: {shown}", None


def _match_binding(bitfields, access, binding):
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

    link = binding.type_link
    definition = bitfields.definitions.get(link.definition) if link else None
    if binding.status in ('INTEGER', 'SCALAR') or (link and link.integer_type and not definition):
        return 'SKIP', '', None
    if binding.status not in ('STRUCT', 'CONDITIONAL'):
        return 'REVIEW', f"'{name}' 의 변수 선언 연결이 불확실함 ({binding.status})", None
    first = binding.declarations[0]
    reason = f'C 선언 {Path(first.origin.file).name}:{first.origin.line} ({first.spelling})'
    candidates = {definition.origin: definition} if definition else {}
    if not candidates:
        # Check definitions of the declared type, never unrelated member names.
        for declaration in binding.declarations:
            if declaration.target in bitfields.definitions:
                record = bitfields.definitions[declaration.target]
                candidates[record.origin] = record
                continue
            spelling = declaration.spelling.removeprefix('struct ')
            names = matching_names(bitfields.type_candidates, spelling)
            if not names:
                return 'REVIEW', f"'{name}'의 선언 타입 '{spelling}' 후보를 찾지 못함 / {reason}", None
            for spelling in names:
                candidates.update((d.origin, d) for d in bitfields.type_candidates[spelling])
    status, reason, field = _compare_member_candidates(list(candidates.values()), member, reason)
    # Ordinary-member classification does not require a bitfield access check.
    if status != 'TYPE_MATCH':
        return status, reason, field
    compatible = True
    for declaration in binding.declarations:
        dimensions = len(declaration.arrays)
        decay = int(declaration.kind == 'PARAMETER' and dimensions == 1)
        depth = link.pointer_depth if definition else declaration.pointer_depth
        compatible &= supports_access(depth + decay, dimensions - decay)
    if not compatible:
        return 'REVIEW', f"'{name}' 의 배열/포인터 선언과 접근 형태가 다름", None
    return status, reason, field


def _member_catalog_match(bitfields, access, missing_reason):
    """Infer the root's struct type first, then compare only members of that type."""
    base, member = access['base'], access['member']
    stems = [base]
    if access['operator'] == '->' and base.casefold().endswith('_p'):
        stems.insert(0, base[:-2])
    attempts = list(dict.fromkeys(name for stem in stems
                    for name in ((stem,) if stem.casefold().endswith('_t') else (stem + '_t', stem))))
    candidates = {}
    type_names = ()
    for spelling in attempts:
        type_names = matching_names(bitfields.type_candidates, spelling)
        if type_names:
            for name in type_names:
                candidates.update((d.origin, d) for d in bitfields.type_candidates[name])
            break
    if not candidates:
        return 'REVIEW', (missing_reason + ' / 변수명에 대응하는 구조체 타입을 찾지 못함 '
                          f"({', '.join(attempts)})"), None
    candidates = list(candidates.values())
    locations = [f'{d.name} ({Path(d.origin.file).name}:{d.origin.line})' for d in candidates]
    shown = ', '.join(locations[:5]) + (f' 외 {len(locations) - 5}개' if len(locations) > 5 else '')
    reason = (f"변수 선언 연결 없이 변수명으로 타입 추정: '{base}' → {', '.join(type_names)} / "
              f"해당 구조체의 '{member}' 비교 / {shown}")
    return _compare_member_candidates(candidates, member, reason)


def _compare_member_candidates(candidates, member, reason):
    """Compare source-declared ranges, independently of build-branch activation."""
    if not candidates:
        return 'REVIEW', f"'{member}'의 구조체 타입 후보가 없음 / {reason}", None
    if any(d.status not in ('RESOLVED', 'CONDITIONAL') for d in candidates):
        return 'REVIEW', f'구조체 정의 후보가 충돌함 / {reason}', None
    members = {}
    for definition in candidates:
        available = set(definition.ordinary_members) | {f.declaration.member for f in definition.fields}
        names = matching_names(available, member)
        if not names:
            return 'REVIEW', f"MEMBER_NOT_FOUND: '{definition.name}'에서 '{member}' 멤버를 찾지 못함 / {reason}", None
        if len(names) != 1:
            return 'REVIEW', f"대소문자만 다른 멤버 후보가 여러 개 있음: {', '.join(sorted(names))} / {reason}", None
        members[definition.origin] = names[0]
    fields = [f for d in candidates for f in d.fields if f.declaration.member == members[d.origin]]
    ordinary = [d for d in candidates if members[d.origin] in d.ordinary_members]
    if not fields and ordinary:
        spellings = ', '.join(sorted(set(members.values())))
        return 'SKIP', f"일반 멤버: '{spellings}' (비트필드 아님) / {reason}", None
    if ordinary:
        return 'REVIEW', f'후보에 일반 멤버와 비트필드가 함께 있음 / {reason}', None
    unresolved = [f for f in fields if f.status not in ('RESOLVED', 'CONDITIONAL')
                  or f.declaration.minimum is None or f.declaration.maximum is None]
    if unresolved:
        details = ', '.join(f'{f.declaration.declared_type} {f.declaration.member}:{f.width_expression} '
                            f'({f.status}, {Path(f.declaration.file).name}:{f.declaration.line})'
                            for f in unresolved[:5])
        return 'REVIEW', f'비트필드 후보의 타입·부호 또는 폭을 계산하지 못함: {details} / {reason}', (
            fields[0].declaration if len(fields) == 1 else None)
    ranges = {(f.declaration.minimum, f.declaration.maximum) for f in fields}
    if len(ranges) != 1:
        details = ', '.join(f'{f.declaration.struct_name}.{f.declaration.member} '
                            f'{f.declaration.minimum}~{f.declaration.maximum} '
                            f'({Path(f.declaration.file).name}:{f.declaration.line})' for f in fields[:5])
        return 'REVIEW', f'같은 타입·멤버의 후보별 범위가 다름: {details} / {reason}', None
    field = fields[0].declaration
    if len(fields) > 1:
        field = replace(field, struct_name=' / '.join(sorted({d.name for d in candidates})))
        reason += ' / 후보들의 범위 동일'
    assumptions = sorted({note for f in fields for note in f.type_link.assumptions})
    if assumptions:
        reason += ' / ' + ' / '.join(assumptions)
    return 'TYPE_MATCH', reason, field


def is_function_sheet(name):
    return re.search(r'\d', name) is not None


def compare_book(common, catalog, book, relative_path=None):
    """Compare with C/H evidence. The legacy common argument is ignored."""
    results, coverage = [], []
    filename = source_context(relative_path or Path(book.path).name)
    for sheet in book.sheets:
        if not is_function_sheet(sheet.name):
            continue
        from .a_extraction import scan_a
        from .b_extraction import scan_b, DECLARATION_NAME
        a_headers, _, a_candidates = scan_a(sheet)
        a_rows = {c['row'] for c in a_candidates if c['range_cell'] is not None}
        table = rows(sheet)
        b_headers, _, b_candidates = scan_b(sheet, table)
        b_rows = {c['row'] for c in b_candidates if c['range_cell'] is not None}
        for candidate in (*a_candidates, *b_candidates):
            expr = candidate['name']
            normalized = normalize_member_expression(expr)
            if candidate['range_cell'] is None and DECLARATION_NAME.fullmatch(normalized) and re.search(r'\.|->', normalized):
                results.append(Result(
                    id=f"{book.path}|{sheet.name}|{candidate['name_cell']}|unresolved-range",
                    file=book.path, sheet=sheet.name, cell='', expression=expr, status='REVIEW',
                    reason=f"RANGE_CELL_UNRESOLVED: Name {candidate['name_cell']}에서 변수는 읽었으나 "
                           f"Range/Scope 셀을 확정하지 못함 ({candidate['reason']})",
                    format='B' if candidate['format'] == 'A' else 'A', backend=book.backend))
        headers, issues = {}, []
        for row, cs in table.items():
            labels = {}
            for c in cs:
                labels.setdefault(norm(c), []).append(c)
            if row in a_headers:
                ah = a_headers[row]
                headers[row] = ('B',[ah[k] for k in ('name','datatype','scope')],
                                [ah[k] for k in ('input/output','property')])
                if row in b_headers or all(k in labels for k in ('name','type','range')):
                    issues.append('Conflicting A/B header')
                continue
            if row in b_headers:
                headers[row] = ('A',[b_headers[row][k] for k in ('name','type','range')],[])
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
        function_name = sheet_function_name(sheet.name)
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
                        or parse_expression(normalize_member_expression(text(populated[0]))) is None:
                    section = title                                   # 섹션 제목 행 (Global Variable, Local Variable, …)
                    if title.startswith(('calibration','캘리브레이션','보정')):
                        header = None
                    continue
            if header is None or section == 'calibration':
                continue
            fmt,h,classification = header
            if fmt == 'B' and row not in a_rows:
                continue
            if fmt == 'A' and row not in b_rows:
                continue
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
            normalized = normalize_member_expression(expr)
            access = MEMBER_ACCESS.fullmatch(normalized)
            if not DECLARATION_NAME.fullmatch(normalized) or not re.search(r'\.|->', normalized):
                continue
            if access and identifier(access['base']) and (not access['index'] or
                    access['index'].isdigit() or identifier(access['index'])):
                status,reason,f = _catalog_match(catalog, access, source_file, function,
                                                 function_name=function_name)
            else:
                status,reason,f = ('REVIEW',
                    'MEMBER_PATH_UNRESOLVED: Excel 변수 행을 읽었으나 중첩 멤버 또는 첨자 경로의 타입을 확정하지 못함',
                    None)
            if status == 'SKIP' and not reason:
                continue
            rid = f'{book.path}|{sheet.name}|{rc.address}'
            base = dict(id=rid,file=book.path,sheet=sheet.name,cell=rc.address,expression=expr,existing=rc.formula if rc.is_formula else rc.value,format=fmt,backend=book.backend)
            if f:
                base.update(struct=f.struct_name,member=f.member,width=f.width,signed=f.signed,source=f.file,source_line=f.line,
                            expected=f'{f.minimum}~{f.maximum}' if f.minimum is not None else '')
                design_type=text(tc).strip()
                if design_type and f.declared_type and ''.join(design_type.split()).casefold() != ''.join(f.declared_type.split()).casefold():
                    reason=f"{reason} / 설계서 Type '{design_type}' ≠ 코드 '{f.declared_type}'"
            if status == 'TYPE_MATCH':
                binding_reason = reason
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
                reason += ' / ' + binding_reason
            results.append(Result(**base,status=status,reason=reason,candidate=status in ('MISMATCH','CREATE')))
    # Duplicate target coordinates are never independent edits.
    counts = {}
    for r in results:
        k = (r.sheet,r.cell)
        counts[k] = counts.get(k,0) + 1
    results = [replace(r,status='REVIEW',reason='Duplicate target cell',candidate=False)
               if r.cell and counts[(r.sheet,r.cell)] > 1 else r for r in results]
    if getattr(catalog, 'build_context', None) is not None:
        from .build_options import annotate_results
        results = annotate_results(results, catalog, relative_path or Path(book.path).name)
    return results, coverage
