"""Conservative C/H bit-field inventory for use independently of Excel."""
import ast
from collections import Counter, defaultdict
from dataclasses import dataclass, field, replace
from pathlib import Path
import operator
import re

from .analysis import (BASE_PROFILES, QUALIFIERS, BitField, Diagnostic, SourceModel, Token,
                       _CMT_STR, _Parser, identifier, tokenize)


@dataclass(frozen=True, order=True)
class Origin:
    file: str
    line: int
    ordinal: int = 0


@dataclass(frozen=True)
class TypeLink:
    status: str
    definition: Origin | None = None
    integer_type: str = ''
    evidence: tuple[Origin, ...] = ()
    pointer_depth: int = 0
    qualifiers: tuple[tuple[str, ...], ...] = ()


@dataclass(frozen=True)
class VariableDeclaration:
    name: str
    spelling: str
    origin: Origin
    arrays: tuple[str, ...]
    qualifiers: tuple[tuple[str, ...], ...]
    storage: str
    kind: str
    conditional: bool
    target: Origin | None = None

    @property
    def pointer_depth(self):
        return len(self.qualifiers) - 1


@dataclass(frozen=True)
class VariableBinding:
    status: str
    declarations: tuple[VariableDeclaration, ...] = ()
    type_link: TypeLink | None = None
    nested_only: bool = False


@dataclass(frozen=True)
class FunctionScope:
    file: str
    status: str
    variables: dict[str, VariableBinding] = field(default_factory=dict)


@dataclass(frozen=True)
class _Function:
    name: str
    origin: Origin
    internal: bool
    conditional: bool
    parameters: tuple[VariableDeclaration | None, ...]
    locals: tuple[tuple[VariableDeclaration, bool], ...]
    definition: bool
    return_type: tuple[str, ...]
    prototyped: bool


@dataclass(frozen=True)
class CatalogField:
    declaration: BitField
    width_expression: str
    status: str
    type_link: TypeLink


@dataclass(frozen=True)
class CatalogStruct:
    origin: Origin
    name: str
    tag: str
    aliases: tuple[str, ...]
    status: str
    fields: tuple[CatalogField, ...]


@dataclass
class BitfieldCatalog:
    definitions: dict[Origin, CatalogStruct]
    links: dict[tuple[str, str], TypeLink]
    diagnostics: tuple[Diagnostic, ...]
    source_scopes: dict[str, dict[str, bool]] = field(default_factory=dict)
    conditional_declarations: frozenset[Origin] = frozenset()
    variables: dict[str, tuple[VariableDeclaration, ...]] = field(default_factory=dict)
    variable_bindings: dict[tuple[str | None, str], VariableBinding] = field(default_factory=dict)
    functions: dict[str, tuple[FunctionScope, ...]] = field(default_factory=dict)

    def lookup_function(self, file, name):
        candidates = self.functions.get(name, ())
        if file is not None:
            file = str(Path(file).resolve())
            visible = self.source_scopes.get(file, {})
            candidates = (tuple(scope for scope in candidates if scope.file == file)
                          or tuple(scope for scope in candidates if scope.file in visible))
        if len(candidates) == 1:
            scope = candidates[0]
            if file and self.source_scopes.get(file, {}).get(scope.file, False):
                return replace(scope, status='CONDITIONAL')
            return scope
        if candidates or name in self.functions:
            return FunctionScope(file or '', 'AMBIGUOUS')
        return FunctionScope(file or '', 'MISSING')

    def lookup_variable(self, file, name):
        scope = str(Path(file).resolve()) if file is not None else None
        return self.variable_bindings.get((scope, name), VariableBinding('MISSING'))

    def lookup(self, file, spelling):
        """Resolve a type in an input file's scope, preserving uncertainty."""
        return self.links.get((str(Path(file).resolve()), spelling), TypeLink('UNRESOLVED'))


@dataclass(frozen=True)
class _Declaration:
    target: str | Origin
    origin: Origin
    conditional: bool = False
    qualifiers: tuple[tuple[str, ...], ...] = ()


@dataclass
class _Unit:
    types: dict = field(default_factory=lambda: defaultdict(list))
    macros: dict = field(default_factory=lambda: defaultdict(list))
    includes: dict = field(default_factory=dict)
    definitions: dict = field(default_factory=dict)
    reviews: list = field(default_factory=list)
    conditional_declarations: set[Origin] = field(default_factory=set)
    variables: list[VariableDeclaration] = field(default_factory=list)
    functions: list[_Function] = field(default_factory=list)


class _Collector(_Parser):
    collect_bitfields = True

    def parse_bitfield(self, st, member, declared, width_expr, at):
        # The temporary field stores the raw width until the linkage pass.
        return BitField(st.name, st.tag, member, declared, 0, self.file, at, note=width_expr)


_DIRECTIVE = re.compile(r'^\s*#\s*(\w+)\b\s*(.*)$')
_INCLUDE = re.compile(r'"([^"]+)"$')
_INTEGER = re.compile(r'\b(0[xX][0-9a-fA-F]+|0[bB][01]+|0[0-7]*|[1-9][0-9]*)([uUlL]*)\b')
_OPERATORS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.FloorDiv: operator.floordiv, ast.Mod: operator.mod, ast.LShift: operator.lshift,
    ast.RShift: operator.rshift, ast.BitAnd: operator.and_, ast.BitOr: operator.or_,
    ast.BitXor: operator.xor,
}


def _guard(directives, lines):
    """Recognize a conventional whole-file include guard."""
    if len(directives) < 3:
        return None
    (start, opening, name), (_, define, body), (end, closing, _) = (
        directives[0], directives[1], directives[-1])
    if (opening != 'ifndef' or define != 'define' or body != name or closing != 'endif'
            or not identifier(name) or any(line.strip() for line in lines[:start - 1] + lines[end:])):
        return None
    depth = 0
    for _, command, _ in directives[:-1]:
        if command in ('if', 'ifdef', 'ifndef'):
            depth += 1
        elif command == 'endif':
            depth -= 1
        if depth == 0 or (depth == 1 and command in ('elif', 'else')):
            return None
    return start if depth == 1 else None


def _collect(source, file, files):
    unit = _Unit()
    lines = source.lstrip('\ufeff').splitlines()
    # Join splices before stripping comments, retaining physical line numbers.
    for i in range(len(lines) - 2, -1, -1):
        if lines[i].endswith('\\'):
            lines[i] = lines[i][:-1] + lines[i + 1]
            lines[i + 1] = ''
    source = _CMT_STR.sub(lambda m: m[0] if m[0][0] in "\"'" else ' ' + '\n' * m[0].count('\n'),
                          '\n'.join(lines))
    lines = source.split('\n')
    directives = [(i, *match.groups()) for i, line in enumerate(lines, 1)
                  if (match := _DIRECTIVE.fullmatch(line))]
    guard = _guard(directives, lines)
    conditions, conditional = [], set()
    by_line = {line: (command, body.strip()) for line, command, body in directives}
    for i, line in enumerate(lines, 1):
        if i not in by_line:
            if any(conditions):
                conditional.add(i)
            continue
        command, body = by_line[i]
        lines[i - 1] = ''
        if command in ('if', 'ifdef', 'ifndef'):
            conditions.append(i != guard)
        elif command == 'endif' and conditions:
            conditions.pop()
        elif command in ('define', 'undef'):
            match = re.match(r'([A-Za-z_]\w*)(.*)', body)
            if match:
                name, expr = match.groups()
                target = expr.strip() if command == 'define' and not expr.startswith('(') else ''
                unit.macros[name].append(_Declaration(target, Origin(file, i), any(conditions)))
        elif command == 'include':
            match = _INCLUDE.fullmatch(body)
            if match:
                local = str((Path(file).parent / match[1]).resolve())
                if local in files:
                    unit.includes[local] = unit.includes.get(local, True) and any(conditions)
    parser = _Collector(tokenize('\n'.join(lines), file, unit.reviews), file, SourceModel())
    def function(head, body):
        declaration = _collect_function(parser, head, body, conditional)
        if declaration is not None:
            unit.functions.append(declaration)

    for tokens in parser.declarations(function):
        words = [token.text for token in tokens]
        origin = Origin(file, tokens[0].line)
        uncertain = any(token.line in conditional for token in tokens)
        if uncertain:
            unit.conditional_declarations.add(origin)
        typedef = words[0] == 'typedef'
        offset = next((i for i in range(int(typedef), len(words))
                       if words[i] not in QUALIFIERS), len(words))
        try:
            kind = words[offset]
            after_kind = offset + 1
            if after_kind < len(words) and identifier(words[after_kind]):
                after_kind += 1
            if kind == 'struct' and words[after_kind:after_kind + 1] == ['{']:
                nested_start = len(parser.m.types)
                st, end, tag = parser.parse_aggregate(tokens, offset, typedef, origin.line)
                if st is None:
                    continue
                # Column-independent identity also distinguishes definitions on the same line.
                definition = _definition_id(Origin(file, st.line), unit.definitions)
                names = []
                if typedef:
                    names = _typedefs(parser, tokens[end:], definition, origin, uncertain, unit)
                if names:
                    parser.rename(st, names[0])
                unit.definitions[definition] = (st, uncertain)
                for nested in parser.m.types[nested_start:]:
                    nested_origin = _definition_id(Origin(file, nested.line), unit.definitions)
                    unit.definitions[nested_origin] = (nested, uncertain)
                if tag:
                    unit.types['struct ' + tag].append(_Declaration(definition, origin, uncertain))
                if not typedef:
                    _collect_variables(parser, words[:offset], tokens[end:], origin, uncertain, unit,
                                       'struct ' + tag if tag else st.name, definition)
            elif kind == 'struct' and len(words) == 2 and identifier(words[1]):
                spelling = 'struct ' + words[1]
                unit.types[spelling].append(_Declaration(spelling, origin, uncertain))
            elif typedef and '{' in words:
                unit.reviews.append(Diagnostic(file, origin.line, 'Unsupported aggregate typedef'))
            elif typedef:
                _collect_typedef(parser, tokens[1:], origin, uncertain, unit)
            else:
                prototype = _collect_function(parser, tokens, None, conditional)
                if prototype is not None:
                    unit.functions.append(prototype)
                else:
                    _collect_variables(parser, None, tokens, origin, uncertain, unit)
        except (IndexError, ValueError, RecursionError) as exc:
            unit.reviews.append(Diagnostic(file, origin.line, f'Declaration skipped: {exc}'))
    return unit


def _definition_id(origin, definitions):
    ordinal = 0
    while Origin(origin.file, origin.line, ordinal) in definitions:
        ordinal += 1
    return Origin(origin.file, origin.line, ordinal)


_CV = {'const', 'volatile', 'restrict', '__restrict', '__restrict__'}
_STORAGE = {'extern', 'static', '_Thread_local', 'register', 'auto'}
_NON_STRUCT_SCALARS = {'void', 'float', 'double', 'long double'}


def _declarator(words):
    """Split a simple object declarator into name, pointer qualifiers and dimensions."""
    cursor, qualifiers = 0, []
    while cursor < len(words) and words[cursor] == '*':
        cursor += 1
        start = cursor
        while cursor < len(words) and words[cursor] in _CV:
            cursor += 1
        qualifiers.append(tuple(sorted(set(words[start:cursor]))))
    if cursor == len(words) or not identifier(words[cursor]):
        return None
    name = words[cursor]
    cursor += 1
    dimensions = []
    while cursor < len(words) and words[cursor] == '[':
        try:
            end = words.index(']', cursor + 1)
        except ValueError:
            return None
        extent = words[cursor + 1:end]
        if '[' in extent:
            return None
        dimensions.append(' '.join(extent))
        cursor = end + 1
    if cursor != len(words) and words[cursor] != '=':
        return None
    return name, tuple(qualifiers), tuple(dimensions), cursor < len(words)


def _specifier_end(words):
    for i, word in enumerate(words):
        if word in ('*', '('):
            return i - int(word == '(' and i > 0 and identifier(words[i - 1]))
        if word in ('[', '='):
            return i - 1
    return len(words) - 1


def _collect_variables(parser, spec, tokens, origin, uncertain, unit, spelling='', target=None):
    unit.variables.extend(_variable_declarations(parser, spec, tokens, origin, uncertain, spelling, target))


def _variable_declarations(parser, spec, tokens, origin, uncertain, spelling='', target=None):
    parts = parser.split(tokens)
    if not parts:
        return
    if spec is None:
        first = [token.text for token in parts[0]]
        end = _specifier_end(first)
        if end < 1:
            return
        spec, parts[0] = first[:end], parts[0][end:]
        spelling = ' '.join(word for word in spec if word not in _CV | _STORAGE)
    storage = 'static' if 'static' in spec else 'extern' if 'extern' in spec else ''
    if not spelling or ('(' in spec or '{' in spec) or any(word in spec for word in ('register', 'auto')):
        return
    base_cv = tuple(sorted(set(spec) & _CV))
    for part in parts:
        parsed = _declarator([token.text for token in part])
        if parsed is None:
            continue
        name, pointers, arrays, initialized = parsed
        kind = 'DEFINITION' if initialized else 'DECLARATION' if storage == 'extern' else 'TENTATIVE'
        yield VariableDeclaration(name, spelling, origin, arrays,
                                  (base_cv, *pointers), storage, kind, uncertain, target)


def _unsupported_pointer(tokens, origin, uncertain):
    words = [t.text for t in tokens]
    if '(' not in words or '=' in words[:words.index('(')]:
        return None
    opening = words.index('(')
    base = ' '.join(word for word in words[:opening] if word not in _CV | _STORAGE | {'*'})
    if not _type_spelling(base):
        return None
    match = re.match(r"\( (?:\* |const |volatile |restrict )+([A-Za-z_]\w*) (?:\)|\[)",
                     ' '.join(words[opening:]))
    if not match:
        return None
    return VariableDeclaration(match[1], '<complex declarator>', origin, (), ((),), '',
                               'UNSUPPORTED', uncertain)


def _local_declarations(parser, body, conditional):
    """Scan statement boundaries and retain whether a declaration is in a nested scope."""
    pending, nesting = [], 0
    declarations, shadowed_types = [], set()
    parentheses = brackets = initializer = 0
    for token in body:
        word = token.text
        if word == '{' and not parentheses and not brackets:
            head = [t.text for t in pending]
            aggregate = bool(head and head[0] in ('struct', 'union', 'enum', 'typedef'))
            if initializer or '=' in head or aggregate:
                initializer += 1
            else:
                pending.clear()
                nesting += 1
                continue
        elif word == '}' and not parentheses and not brackets:
            if initializer:
                initializer -= 1
            else:
                pending.clear()
                nesting -= 1
                continue
        elif word == '(':
            parentheses += 1
        elif word == ')':
            parentheses -= 1
        elif word == '[':
            brackets += 1
        elif word == ']':
            brackets -= 1
        if word == ';' and parentheses == 1 and pending and pending[0].text == 'for':
            # The initializer's declaration is scoped to the loop, not the whole function.
            if len(pending) > 2 and not any(t.text == ';' for t in pending):
                part = pending[2:]
                origin = Origin(parser.file, part[0].line)
                declarations.extend((d, True) for d in _variable_declarations(
                    parser, None, part, origin, any(t.line in conditional for t in part))
                    if _type_spelling(d.spelling))
        if word != ';' or parentheses or brackets or initializer:
            pending.append(token)
            continue
        if not pending:
            continue
        origin = Origin(parser.file, pending[0].line)
        uncertain = any(t.line in conditional for t in pending)
        words = [t.text for t in pending]
        if words[0] == 'typedef':
            scratch = _Unit()
            if '{' in words and '}' in words:
                tail = len(words) - words[::-1].index('}')
                _typedefs(parser, pending[tail:], '', origin, uncertain, scratch)
            else:
                _collect_typedef(parser, pending[1:], origin, uncertain, scratch)
            shadowed_types.update(scratch.types)
        elif words[0] in ('struct', 'union', 'enum') and (
                len(words) == 2 or (len(words) > 2 and words[2] == '{') or words[1:2] == ['{']):
            if len(words) > 1 and identifier(words[1]):
                shadowed_types.add(' '.join(words[:2]))
            if '}' in words:
                tail = len(words) - words[::-1].index('}')
                declarations.extend((d, nesting > 0) for d in _variable_declarations(
                    parser, [], pending[tail:], origin, uncertain, '<local aggregate>'))
        else:
            parsed = [d for d in _variable_declarations(parser, None, pending, origin, uncertain)
                      if _type_spelling(d.spelling)]
            if not parsed:
                unsupported = _unsupported_pointer(pending, origin, uncertain)
                if unsupported:
                    parsed.append(unsupported)
            declarations.extend((d, nesting > 0) for d in parsed)
        pending.clear()
    return tuple((replace(d, spelling='<local type>') if d.spelling in shadowed_types else d, nested)
                 for d, nested in declarations)


def _type_spelling(spelling):
    words = spelling.split()
    return (spelling in BASE_PROFILES or spelling == 'char' or spelling in _NON_STRUCT_SCALARS
            or (len(words) == 1 and identifier(words[0]))
            or (len(words) == 2 and words[0] in ('struct', 'union', 'enum') and identifier(words[1])))


def _collect_function(parser, head, body, conditional):
    words = [t.text for t in head]
    if not words or words[0] == 'typedef' or '(' not in words:
        return None
    opening = words.index('(')
    if (opening < 2 or not identifier(words[opening - 1]) or words[-1] != ')'
            or '=' in words[:opening] or '{' in words[:opening]):
        return None
    return_type = tuple(word for word in words[:opening - 1] if word not in _STORAGE | {'inline'})
    if not _type_spelling(' '.join(word for word in return_type if word != '*' and word not in _CV)) and return_type != ('void',):
        return None
    parameters = []
    if words[opening + 1:-1] != ['void']:
        for part in parser.split(head[opening + 1:-1]):
            origin = Origin(parser.file, part[0].line)
            uncertain = any(t.line in conditional for t in part)
            declarations = tuple(_variable_declarations(parser, None, part, origin, uncertain))
            parameter = next((d for d in declarations if _type_spelling(d.spelling)), None)
            if parameter is None and body is None:
                # Unnamed parameters in prototypes still carry type information.
                unnamed = next(iter(_variable_declarations(
                    parser, None, [*part, Token('_rv_parameter', part[-1].line)], origin, uncertain)), None)
                if unnamed and _type_spelling(unnamed.spelling):
                    parameter = unnamed
            parameter = parameter or _unsupported_pointer(part, origin, uncertain)
            parameters.append(replace(parameter, kind='PARAMETER') if parameter else None)
    return _Function(words[opening - 1], Origin(parser.file, head[0].line),
                     'static' in words[:opening], any(t.line in conditional for t in head),
                     tuple(parameters), tuple(_local_declarations(parser, body or (), conditional)),
                     body is not None, return_type, opening + 1 < len(words) - 1)


def _typedefs(parser, tail, target, origin, uncertain, unit, base_cv=()):
    names = []
    for part in parser.split(tail):
        words = [token.text for token in part]
        parsed = _declarator(words)
        if parsed is not None:
            name, pointers, arrays, initialized = parsed
            if not arrays and not initialized:
                if not pointers:
                    names.append(name)
                unit.types[name].append(_Declaration(target, origin, uncertain, (base_cv, *pointers)))
                continue
        name = next((word for word in words if identifier(word)), '')
        if name:
            unit.types[name].append(_Declaration('', origin, uncertain))
        unit.reviews.append(Diagnostic(origin.file, origin.line, 'Unsupported typedef declarator'))
    return names


def _collect_typedef(parser, tokens, origin, uncertain, unit):
    if not tokens:
        return
    first = [token.text for token in parser.split(tokens)[0]]
    end = _specifier_end(first)
    if end < 1:
        return
    spec = first[:end]
    target = ' '.join(word for word in spec if word not in QUALIFIERS)
    _typedefs(parser, tokens[end:], target, origin, uncertain, unit, tuple(sorted(set(spec) & _CV)))
    if target.startswith('struct ') and len(target.split()) == 2:
        unit.types[target].append(_Declaration(target, origin, uncertain))


def _qualify(link, qualifiers):
    if not qualifiers or link.status not in ('STRUCT', 'INTEGER', 'SCALAR'):
        return link
    base = link.qualifiers or ((),)
    combined = (*base[:-1], tuple(sorted(set(base[-1]) | set(qualifiers[0]))), *qualifiers[1:])
    return replace(link, pointer_depth=len(combined) - 1, qualifiers=combined)


def _visible(file, units):
    visible, pending = {}, [(file, False)]
    while pending:
        item, conditional = pending.pop()
        if item not in visible or (visible[item] and not conditional):
            visible[item] = conditional
            pending.extend((path, conditional or guarded)
                           for path, guarded in units[item].includes.items())
    return visible


class _Registry:
    def __init__(self, units):
        self.units = units
        self.scopes = {file: _visible(file, units) for file in units}
        self.definitions = {key: value for unit in units.values() for key, value in unit.definitions.items()}

    def resolve(self, file, name, root=None, seen=frozenset()):
        root = root or file
        if (name in BASE_PROFILES and name != 'bool') or name == 'char':
            return TypeLink('INTEGER', integer_type=name)
        if name in _NON_STRUCT_SCALARS:
            return TypeLink('SCALAR', integer_type=name)
        key = (file, name)
        if key in seen or len(seen) >= 128:
            return TypeLink('UNRESOLVED')
        scope = self.scopes[file]
        is_tag = name.startswith('struct ')
        if is_tag:
            scope = self.scopes[root]
        declarations = [d for path in sorted(scope) for d in self.units[path].types.get(name, ())]
        if is_tag and root == file and Path(file).suffix.lower() == '.h' and not any(
                isinstance(d.target, Origin) for d in declarations) and declarations:
            # An including C file can complete an explicit header forward declaration.
            declarations.extend(d for path, visible in sorted(self.scopes.items())
                                if path != file and visible.get(file) is False
                                for d in self.units[path].types.get(name, ())
                                if isinstance(d.target, Origin))
        evidence = tuple(sorted({d.origin for d in declarations}))
        if any(d.conditional or scope.get(d.origin.file, False) for d in declarations):
            return TypeLink('CONDITIONAL', evidence=evidence)
        if is_tag:
            declarations = [d for d in declarations if d.target != name]
            if not declarations and evidence:
                return TypeLink('INCOMPLETE', evidence=evidence)
        if not declarations:
            return TypeLink('UNRESOLVED', evidence=evidence)
        targets = dict.fromkeys((d.origin.file, d.target, d.qualifiers) for d in declarations)
        links = [_qualify(TypeLink('STRUCT', target) if isinstance(target, Origin)
                          else self.resolve(path, target, root, seen | {key}), qualifiers)
                 for path, target, qualifiers in targets]
        first = links[0]
        outcomes = {(link.status, link.definition, link.integer_type, link.qualifiers) for link in links}
        evidence = tuple(sorted(set(evidence).union(*(link.evidence for link in links))))
        if len(outcomes) > 1 or (first.status in ('UNRESOLVED', 'INCOMPLETE')
                                and len({d.target for d in declarations}) > 1):
            return TypeLink('AMBIGUOUS', evidence=evidence)
        return replace(first, evidence=evidence)

    def width(self, expression, file):
        scope = self.scopes[file]

        def expand(text, stack):
            def substitute(match):
                name = match[0]
                if name in stack or len(stack) >= 32:
                    raise ValueError('recursive macro')
                declarations = [d for path in scope for d in self.units[path].macros.get(name, ())]
                bodies = {d.target for d in declarations}
                if len(bodies) != 1 or any(d.conditional or scope[d.origin.file] for d in declarations):
                    raise ValueError('missing, conflicting or conditional macro')
                body = bodies.pop()
                if not body:
                    raise ValueError('unsupported macro')
                return expand(body, stack | {name})

            if len(text) > 512:
                raise ValueError('oversized expression')
            # Textual expansion preserves C operator precedence for unparenthesized macros.
            parts, size, end = [], 0, 0
            for match in re.finditer(r"\b[A-Za-z_]\w*\b", text):
                part = text[end:match.start()] + substitute(match)
                size += len(part)
                if size > 512:
                    raise ValueError('oversized expansion')
                parts.append(part)
                end = match.end()
            return ''.join(parts) + text[end:]

        def evaluate(node):
            if isinstance(node, ast.Constant) and type(node.value) is int:
                value = node.value
            elif isinstance(node, ast.BinOp) and type(node.op) in _OPERATORS:
                left, right = evaluate(node.left), evaluate(node.right)
                if isinstance(node.op, (ast.LShift, ast.RShift)) and right > 63:
                    raise ValueError('oversized shift')
                value = _OPERATORS[type(node.op)](left, right)
            else:
                raise ValueError('unsupported width expression')
            if not 0 <= value <= 32767:
                raise ValueError('constant outside supported width-expression domain')
            return value

        def literal(match):
            digits = match[1]
            base = 0 if digits.lower().startswith(('0x', '0b')) else 8 if digits.startswith('0') else 10
            return str(int(digits, base))

        if not expression or len(expression) > 512 or '//' in expression:
            return None
        try:
            expanded = expand(expression, frozenset())
            if '//' in expanded:
                return None
            normalized = _INTEGER.sub(literal, expanded).replace('/', '//')
            if not re.fullmatch(r'[0-9+*/%()<>|&^ \t-]+', normalized):
                return None
            return evaluate(ast.parse(normalized, mode='eval').body)
        except (SyntaxError, ValueError, ZeroDivisionError, RecursionError):
            return None


def _field(raw, registry, blocked):
    width = registry.width(raw.note, raw.file)
    link = registry.resolve(raw.file, raw.declared_type)
    signed = None
    minimum = maximum = None
    integer = link.integer_type
    if blocked:
        status = blocked
    elif width is None or not 1 <= width <= 1024 or (integer == '_Bool' and width != 1):
        status = 'WIDTH_UNRESOLVED'
    elif link.status != 'INTEGER' or link.pointer_depth or integer in ('char', 'int'):
        status = link.status if link.status not in ('INTEGER', 'STRUCT') else 'SIGNEDNESS_UNRESOLVED'
    else:
        signed = BASE_PROFILES[integer][0]
        minimum = -(1 << (width - 1)) if signed else 0
        maximum = (1 << (width - int(signed))) - 1
        status = 'RESOLVED'
    declaration = BitField(raw.struct_name, raw.struct_tag, raw.member, raw.declared_type,
                           width or 0, raw.file, raw.line, signed, minimum, maximum,
                           '' if status == 'RESOLVED' else status)
    return CatalogField(declaration, raw.note, status, link)


def _bind_declarations(registry, declarations, root):
    links = []
    for declaration in declarations:
        file = declaration.origin.file
        context = root if file in registry.scopes.get(root, ()) else file
        if declaration.conditional or registry.scopes[context].get(file, False):
            return VariableBinding('CONDITIONAL', declarations)
        link = (TypeLink('STRUCT', declaration.target) if declaration.target else
                registry.resolve(file, declaration.spelling, context))
        links.append(_qualify(link, declaration.qualifiers))
    identities = {(link.status, link.definition, link.integer_type, link.qualifiers, len(d.arrays))
                  for d, link in zip(declarations, links)}
    if len(identities) != 1:
        return VariableBinding('AMBIGUOUS', declarations)
    if any(len({d.arrays[i] for d in declarations if d.arrays[i]}) > 1
           for i in range(len(declarations[0].arrays))):
        return VariableBinding('AMBIGUOUS', declarations)
    link = links[0]
    evidence = tuple(sorted(set().union(*(item.evidence for item in links),
                                        {d.origin for d in declarations})))
    return VariableBinding(link.status, declarations, replace(link, evidence=evidence))

def _bind_variables(registry):
    groups = defaultdict(list)
    for unit in registry.units.values():
        for declaration in unit.variables:
            groups[declaration.name].append(declaration)

    includers = defaultdict(set)
    for file, scope in registry.scopes.items():
        for visible in scope:
            includers[visible].add(file)

    bindings = {}
    for name, items in groups.items():
        declarations = tuple(items)
        contexts = []
        visible_from = set().union(*(includers[d.origin.file] for d in declarations))
        for file in sorted(visible_from):
            scope = registry.scopes[file]
            local = tuple(d for d in declarations if d.origin.file in scope)
            local_files = {d.origin.file for d in local}
            static_files = {d.origin.file for d in local if d.storage == 'static'}
            if static_files:
                # Internal linkage belongs to a single translation unit.
                binding = (_bind_declarations(registry, local, file) if len(local_files) == 1
                           else VariableBinding('AMBIGUOUS', local))
            else:
                shared = {d.origin.file for d in local if d.storage == 'extern'
                          and d.origin.file != file and not scope[d.origin.file]}
                connected = tuple(d for d in declarations if d.origin.file not in scope
                                  and d.storage != 'static'
                                  and any(registry.scopes[d.origin.file].get(h) is False for h in shared))
                binding = _bind_declarations(registry, local + connected, file)
            bindings[file, name] = binding
            contexts.append(binding)

        if len({d.origin.file for d in declarations}) == 1 and not any(d.storage == 'static' for d in declarations):
            file = declarations[0].origin.file
            binding = bindings[file, name]
            # An incomplete header can have different completions in its includers.
            resolved = {b.type_link.definition for b in contexts if b.status == 'STRUCT'}
            if len(resolved) > 1 or any(b.status == 'AMBIGUOUS' for b in contexts):
                binding = VariableBinding('AMBIGUOUS', declarations)
        elif any(b.declarations and set(b.declarations) == set(declarations)
                 and b.status in ('STRUCT', 'INTEGER', 'SCALAR') for b in contexts) and not any(
                     d.storage == 'static' for d in declarations):
            binding = _bind_declarations(registry, declarations, None)
        else:
            binding = VariableBinding('AMBIGUOUS' if len(declarations) > 1 else 'UNRESOLVED', declarations)
        bindings[None, name] = binding
    return {name: tuple(items) for name, items in groups.items()}, bindings


def _bind_functions(registry, globals_by_scope):
    groups = defaultdict(list)
    for unit in registry.units.values():
        for function in unit.functions:
            groups[function.name].append(function)

    def signature(function, root):
        types = []
        for parameter in function.parameters:
            if parameter is None or parameter.conditional:
                return None
            link = _bind_declarations(registry, (parameter,), root).type_link
            if not link or link.status not in ('STRUCT', 'INTEGER', 'SCALAR'):
                return None
            qualifiers = link.qualifiers or ((),)
            if parameter.arrays:
                qualifiers = (*qualifiers, ())
            # Top-level parameter qualifiers do not affect function compatibility.
            types.append((link.definition, link.integer_type, qualifiers[:-1], len(parameter.arrays[1:])))
        return tuple(types)

    functions = {}
    for name, declarations in groups.items():
        scopes = []
        for definition in (item for item in declarations if item.definition):
            file = definition.origin.file
            visible = registry.scopes[file]
            prototypes = [item for item in declarations if not item.definition
                          and item.origin.file in visible
                          and (not item.internal or item.origin.file == file)]
            status = 'CONDITIONAL' if definition.conditional else 'RESOLVED'
            if any(p.conditional or visible[p.origin.file] for p in prototypes):
                status = 'CONDITIONAL'
            elif prototypes and any(p.return_type != definition.return_type
                                    or (p.prototyped and (signature(p, file) is None
                                        or signature(p, file) != signature(definition, file)))
                                    for p in prototypes):
                status = 'AMBIGUOUS'

            names = defaultdict(list)
            for parameter in definition.parameters:
                if parameter:
                    names[parameter.name].append((parameter, False))
            for declaration, nested in definition.locals:
                names[declaration.name].append((declaration, nested))
            variables = {}
            for variable, candidates in names.items():
                items = tuple(item for item, _ in candidates)
                nested = candidates[0][1]
                previous = globals_by_scope.get((file, variable))
                if len(items) != 1 or (nested and previous is not None):
                    binding = VariableBinding('AMBIGUOUS', items)
                elif items[0].storage == 'extern':
                    if previous and previous.status not in ('STRUCT', 'INTEGER', 'SCALAR'):
                        binding = VariableBinding(previous.status, items)
                    else:
                        binding = _bind_declarations(registry, items + (previous.declarations if previous else ()), file)
                else:
                    binding = _bind_declarations(registry, items, file)
                variables[variable] = replace(binding, nested_only=nested)
            scopes.append(FunctionScope(file, status, variables))
        if scopes:
            functions[name] = tuple(scopes)
    return functions


def build_bitfield_catalog(paths, encoding='utf-8-sig'):
    """Collect input files first, then link declarations by source scope.

    Only include paths that resolve to supplied files participate. Numeric ranges describe
    the declared width and signedness, without asserting target storage capacity or layout.
    """
    files = sorted({str(Path(path).resolve()) for path in paths})
    units, diagnostics = {}, []
    for file in files:
        try:
            raw = Path(file).read_bytes()
        except OSError as exc:
            diagnostics.append(Diagnostic(file, 1, str(exc), 'SOURCE_READ_ERROR'))
            units[file] = _Unit()
            continue
        for codec in dict.fromkeys((encoding, 'utf-8-sig', 'cp949', 'latin-1')):
            try:
                source = raw.decode(codec)
                break
            except UnicodeDecodeError:
                continue
        units[file] = _collect(source, file, files)
        diagnostics.extend(units[file].reviews)
    registry = _Registry(units)
    links, aliases = {}, defaultdict(set)
    for file, scope in registry.scopes.items():
        names = {name for path in scope for name in units[path].types}
        for name in sorted(names):
            link = registry.resolve(file, name)
            links[file, name] = link
            if link.status == 'STRUCT' and not link.pointer_depth and not name.startswith('struct '):
                aliases[link.definition].add(name)
    definitions = {}
    for origin, (st, conditional) in registry.definitions.items():
        status = 'CONDITIONAL' if conditional else 'RESOLVED'
        for spelling in ('struct ' + st.tag, st.name) if st.tag else (st.name,):
            local = units[origin.file].types.get(spelling, ())
            if status == 'RESOLVED' and any(d.target == origin for d in local):
                link = links[origin.file, spelling]
                if link.status != 'STRUCT':
                    status = link.status
        counts = Counter(raw.member for raw in st.fields)
        fields = tuple(_field(raw, registry, status if status != 'RESOLVED' else
                              'AMBIGUOUS' if counts[raw.member] > 1 else '') for raw in st.fields)
        definitions[origin] = CatalogStruct(origin, st.name, st.tag,
                                             tuple(sorted(aliases[origin])), status, fields)
    conditional = frozenset(origin for unit in units.values() for origin in unit.conditional_declarations)
    variables, bindings = _bind_variables(registry)
    return BitfieldCatalog(definitions, links, tuple(diagnostics), registry.scopes, conditional,
                           variables, bindings, _bind_functions(registry, bindings))

