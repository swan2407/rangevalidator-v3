"""Eclipse CDT C settings and isolated, source-mapped C preprocessing.

No compiler is invoked. Each C translation unit starts with a fresh macro table.
Errors discard that translation unit rather than certifying partial output.
"""
from dataclasses import dataclass, field, replace
import hashlib
import io
from pathlib import Path
import re
import xml.etree.ElementTree as ET


class BuildOptionsError(ValueError):
    pass


@dataclass(frozen=True)
class BuildOptions:
    path: str
    sha256: str
    macros: tuple[tuple[str, str], ...]
    include_paths: tuple[str, ...] = ()

    def receipt(self):
        return dict(file=self.path, sha256=self.sha256, macros=dict(self.macros),
                    include_paths=list(self.include_paths), empty_values='defined, empty replacement')


def load_build_options(path):
    path = Path(path).resolve()
    try:
        data = path.read_bytes()
        if len(data) > 2_000_000 or b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
            raise BuildOptionsError('지원하지 않는 XML 선언 또는 너무 큰 설정 파일입니다.')
        root = ET.fromstring(data)
    except (OSError, ET.ParseError) as exc:
        raise BuildOptionsError(f'빌드옵션 XML을 읽지 못했습니다. {exc}') from exc
    if root.tag != 'cdtprojectproperties':
        raise BuildOptionsError('Eclipse CDT 설정 내보내기 XML이 아닙니다.')
    macros, includes, found = {}, [], False
    for section in root.findall('section'):
        kind = section.get('name', '').rsplit('.', 1)[-1]
        for language in section.findall('language'):
            if language.get('name') != 'C Source File':
                continue
            if kind not in ('Macros', 'IncludePaths'):
                continue
            found = True
            if kind == 'Macros':
                for macro in language.findall('macro'):
                    name = (macro.findtext('name') or '').strip()
                    value = (macro.findtext('value') or '').strip()
                    if not re.fullmatch(r'[A-Za-z_]\w*', name, re.ASCII) or '\n' in value or '\r' in value:
                        raise BuildOptionsError(f'지원하지 않는 매크로 이름 또는 값입니다. {name!r}')
                    if name in macros and macros[name] != value:
                        raise BuildOptionsError(f'매크로 값이 중복되어 충돌합니다. {name}')
                    macros[name] = value
            else:
                for entry in language.findall('includepath'):
                    text = (entry.text or '').strip()
                    if not text:
                        continue
                    if entry.get('workspace_path', 'false').lower() == 'true' or '${' in text:
                        raise BuildOptionsError(f'실제 폴더 경로가 필요한 include 설정입니다. {text}')
                    folder = Path(text)
                    if not folder.is_absolute():
                        folder = path.parent / folder
                    includes.append(str(folder.resolve()))
    if not found:
        raise BuildOptionsError('XML에 C Source File용 매크로/include 설정이 없습니다.')
    return BuildOptions(str(path), hashlib.sha256(data).hexdigest(), tuple(macros.items()),
                        tuple(dict.fromkeys(includes)))


def _read(path, encoding):
    raw = Path(path).read_bytes()
    for codec in dict.fromkeys((encoding, 'utf-8-sig', 'cp949', 'latin-1')):
        try:
            return raw.decode(codec).lstrip('\ufeff')
        except UnicodeDecodeError:
            pass


@dataclass
class PreprocessedSources:
    options: BuildOptions
    sources: dict = field(default_factory=dict)
    locations: dict = field(default_factory=dict)
    errors: dict = field(default_factory=dict)
    units: dict = field(default_factory=dict)

    def receipt(self):
        return dict(**self.options.receipt(), translation_units=self.units, errors=self.errors)

    def original(self, file, line):
        return self.locations.get((file, line), (file, line))

    def reason(self, text):
        # Both full paths and short names occur in existing declaration evidence.
        names = {}
        for root in self.sources:
            names.setdefault(Path(root).name, []).append(root)
        alternatives = {root: root for root in self.sources}
        alternatives.update({name: roots[0] for name, roots in names.items() if len(roots) == 1})
        if not alternatives:
            return text
        pattern = r'(' + '|'.join(re.escape(p) for p in sorted(alternatives, key=len, reverse=True)) + r'):(\d+)'
        def location(match):
            root = alternatives[match[1]]
            file, line = self.original(root, int(match[2]))
            return f'{file}:{line}'
        return re.sub(pattern, location, text)


def preprocess_sources(paths, options, encoding='utf-8-sig'):
    from pcpp import Preprocessor

    paths = sorted({str(Path(p).resolve()) for p in paths})
    bundle = PreprocessedSources(options)
    cache, visited = {}, set()

    def read(path):
        if path not in cache:
            cache[path] = _read(path, encoding)
        return cache[path]

    class Processor(Preprocessor):
        def __init__(self):
            self.opened, self.conditions, self.warnings = set(), [], []
            super().__init__()
            self.rewrite_paths = []
            self.auto_pragma_once_enabled = False
            self.undef('__PCPP__')
            for folder in options.include_paths:
                self.add_path(folder)
            for name, value in options.macros:
                # Empty CDT values mean an empty replacement, not an invented 1.
                self.define(name + ' ' + value)

        def on_error(self, file, line, msg):
            raise BuildOptionsError(f'{file}:{line} — {msg}')

        def on_file_open(self, is_system_include, includepath):
            if self.include_depth > 100:
                raise BuildOptionsError('include 중첩이 100단계를 초과했습니다.')
            path = str(Path(includepath).resolve())
            value = read(path)
            self.opened.add(path)
            return io.StringIO(value)

        def on_include_not_found(self, is_malformed, is_system_include, curdir, includepath):
            suffix = '/' + includepath.replace('\\', '/').casefold()
            matches = [p for p in paths if Path(p).as_posix().casefold().endswith(suffix)]
            if not is_malformed and len(matches) == 1:
                # pcpp expects a search directory, including subdirectory suffixes.
                target = Path(matches[0])
                for _ in Path(includepath).parts:
                    target = target.parent
                return str(target)
            detail = '중복 헤더 경로' if len(matches) > 1 else '헤더를 찾을 수 없음'
            self.on_error(self.lastdirective.source, self.lastdirective.lineno,
                          f'{detail} — {includepath}. C/H 파일 또는 XML include 경로를 확인하세요.')

        def compiler_unknown(self, name):
            if re.match(r'__(?:GNUC|clang|ARM|ICC|TASKING|TI_COMPILER|BYTE_ORDER|ORDER_|SIZEOF|STDC)|_MSC_', name):
                self.on_error(self.lastdirective.source, self.lastdirective.lineno,
                              f'컴파일러 매크로 값이 필요합니다 — {name}')

        def on_unknown_macro_in_defined_expr(self, tok):
            self.compiler_unknown(tok.value)
            return False

        def on_unknown_macro_in_expr(self, ident):
            self.compiler_unknown(ident)
            return 0

        def on_unknown_macro_function_in_expr(self, ident):
            self.on_error(self.lastdirective.source, self.lastdirective.lineno,
                          f'조건식의 함수형 매크로를 해석할 수 없습니다 — {ident}')

        def evalexpr(self, tokens):
            self.evaluating_condition = True
            try:
                return super().evalexpr(tokens)
            finally:
                self.evaluating_condition = False

        def expand_macros(self, tokens, expanding_from=None):
            depth = getattr(self, 'expansion_depth', 0)
            self.expansion_depth = depth + 1
            try:
                expanded = super().expand_macros(tokens, expanding_from or [])
                if (depth == 0 and getattr(self, 'evaluating_condition', False)
                        and not any(t.type not in self.t_WS for t in expanded)):
                    self.on_error(self.lastdirective.source, self.lastdirective.lineno,
                                  '조건식이 비어 있습니다. 값 없는 매크로의 숫자 값을 확인하세요.')
                return expanded
            finally:
                self.expansion_depth = depth

        def on_directive_handle(self, directive, toks, ifpassthru, precedingtoks):
            self.lastdirective = directive
            if directive.value in ('if', 'ifdef', 'ifndef', 'elif'):
                self.conditions.append(dict(file=directive.source, line=directive.lineno,
                                            directive=directive.value, expression=''.join(t.value for t in toks)))
            return True

        def on_directive_unknown(self, directive, toks, ifpassthru, precedingtoks):
            text = ''.join(t.value for t in toks)
            if directive.value == 'warning':
                self.warnings.append(f'{directive.source}:{directive.lineno} — {text}')
                return True
            if directive.value == 'pragma' and re.match(r'\s*(pack|section)\b', text):
                return True  # storage layout is outside this tool's range calculation
            self.on_error(directive.source, directive.lineno,
                          f'#{directive.value} {text}')

    def process(root):
        pp = Processor()
        pp.opened.add(root)
        lines, location, current, words = [], [], None, []

        def flush():
            if words:
                lines.append(' '.join(words))
                location.append(current)
                words.clear()

        try:
            pp.parse(read(root), source=root)
            count = 0
            while True:
                tok = pp.token()
                if tok is None:
                    break
                count += 1
                if count > 5_000_000:
                    raise BuildOptionsError('전처리 토큰 수가 한 C 파일당 5백만 개를 초과했습니다.')
                if tok.type in pp.t_WS or not tok.value.strip():
                    continue
                at = (str(Path(tok.source or root).resolve()), tok.lineno)
                if at != current:
                    flush()
                    current = at
                words.append(tok.value)
            flush()
            bundle.sources[root] = '\n'.join(lines)
            bundle.locations.update({(root, i): original for i, original in enumerate(location, 1)})
        except (BuildOptionsError, OSError, RecursionError) as exc:
            bundle.sources[root] = ''
            bundle.errors[root] = str(exc)
        visited.update(pp.opened)
        bundle.units[root] = dict(files=sorted(pp.opened), conditions=pp.conditions,
                                  warnings=pp.warnings, status='REVIEW' if root in bundle.errors else 'APPLIED')

    for root in paths:
        if Path(root).suffix.casefold() == '.c':
            process(root)
    # Headers already consumed by a C file must not be analyzed again with a
    # different, context-free macro state. Unreferenced headers remain usable.
    for root in paths:
        if root not in visited:
            process(root)
    return bundle


def annotate_results(results, catalog, relative_path):
    """Expose physical source locations and block failed build contexts."""
    bundle = getattr(catalog, 'build_context', None)
    if bundle is None:
        return results
    from .engine import _input_source, source_context
    from .common import sheet_function_name
    source = _input_source(catalog.bitfields, source_context(relative_path))
    output = []
    for row in results:
        root = source
        scope = catalog.bitfields.lookup_function(source, sheet_function_name(row.sheet))
        if root not in bundle.errors and scope.status == 'RESOLVED' and (root is not None or not bundle.errors):
            root = scope.file
        errors = ([bundle.errors[root]] if root in bundle.errors else
                  list(bundle.errors.values()) if root is None else [])
        reason = bundle.reason(row.reason)
        note = f'빌드옵션 {Path(bundle.options.path).name} [{bundle.options.sha256[:12]}]'
        if root:
            note += f' / C 문맥 {Path(root).name}'
        if errors:
            output.append(replace(row, status='REVIEW', candidate=False, expected='', width=0,
                                  signed=None, source='', source_line=0,
                                  reason=note + ' / BUILD_OPTIONS_UNRESOLVED — ' + ' / '.join(errors[:3])))
        else:
            file, line = bundle.original(row.source, row.source_line)
            output.append(replace(row, source=file, source_line=line, reason=reason + ' / ' + note))
    return output


def member_scope(bitfields, root):
    """Type-name inference must not import an unrelated C file's active branch."""
    if root is None:
        return bitfields
    views = getattr(bitfields, '_build_scope_views', None)
    if views is None:
        bitfields._build_scope_views = views = {}
    if root not in views:
        view = replace(bitfields, definitions={key: value for key, value in bitfields.definitions.items() if key.file == root},
                       links={key: value for key, value in bitfields.links.items() if key[0] == root})
        if hasattr(bitfields, 'member_types'):
            view.member_types = bitfields.member_types
        views[root] = view
    return views[root]
