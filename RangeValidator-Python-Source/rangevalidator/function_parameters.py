"""Object parameter types, including pointers to arrays, for function linkage.

Layers run from the declared object towards its element type. Function declarators
are deliberately unsupported; an unknown signature must not certify linkage.
"""
from dataclasses import dataclass, replace


@dataclass(frozen=True)
class ParameterType:
    link: object
    qualifiers: tuple = ()
    layers: tuple = ()


def _declarator(words):
    from .bitfield_catalog import _CV
    from .analysis import identifier
    pos = 0

    def parse():
        nonlocal pos
        pointers = []
        while pos < len(words) and words[pos] == '*':
            pos += 1
            cv = set()
            while pos < len(words) and words[pos] in _CV:
                cv.add(words[pos]); pos += 1
            pointers.append(('pointer', tuple(sorted(cv))))
        name, layers = '', []
        if pos < len(words) and words[pos] == '(':
            pos += 1
            name, layers = parse()
            if pos >= len(words) or words[pos] != ')':
                raise ValueError('unclosed declarator')
            pos += 1
        elif pos < len(words) and identifier(words[pos]):
            name = words[pos]; pos += 1
        while pos < len(words) and words[pos] == '[':
            pos += 1
            start = pos
            while pos < len(words) and words[pos] != ']':
                pos += 1
            if pos == len(words):
                raise ValueError('unclosed array')
            layers.append(('array', tuple(words[start:pos])))
            pos += 1
        return name, layers + list(reversed(pointers))

    name, layers = parse()
    if pos != len(words):
        raise ValueError('unsupported declarator')
    return name, tuple(layers)


def _qualify(value, cv):
    layers = list(value.layers)
    for i, (kind, detail) in enumerate(layers):
        if kind == 'pointer':
            layers[i] = (kind, tuple(sorted(set(detail) | set(cv))))
            return replace(value, layers=tuple(layers))
    return replace(value, qualifiers=tuple(sorted(set(value.qualifiers) | set(cv))))


def _layers(registry, file, layers):
    result = []
    for kind, detail in layers:
        if kind == 'array':
            expression = ' '.join(detail)
            extent = registry.width(expression, file) if expression else None
            if expression and (extent is None or extent <= 0):
                raise ValueError('array extent is unresolved')
            detail = extent
        result.append((kind, detail))
    return tuple(result)


def _identity(value):
    return value.link.definition, value.link.integer_type, value.qualifiers, value.layers


def _resolve(registry, file, name, root, seen=frozenset()):
    from .bitfield_catalog import Origin, TypeLink
    from .analysis import matching_names
    key = (file, name)
    if key in seen or len(seen) >= 64:
        raise ValueError('recursive parameter typedef')
    scope = registry.scopes[file]
    names = matching_names({n for path in scope for n in registry.units[path].parameter_typedefs}, name)
    if not names:
        scope = registry.global_scope
        names = matching_names({n for unit in registry.units.values() for n in unit.parameter_typedefs}, name)
    aliases = [a for path in scope for n in names for a in registry.units[path].parameter_typedefs[n]]
    if len(names) > 1:
        raise ValueError('ambiguous parameter typedef')
    if aliases:
        values = []
        for target, cv, words, origin, conditional in aliases:
            _, layers = _declarator(words)
            base = (ParameterType(TypeLink('STRUCT', target)) if isinstance(target, Origin)
                    else _resolve(registry, origin.file, target, root, seen | {key}))
            base = _qualify(base, cv)
            values.append(replace(base, layers=_layers(registry, origin.file, layers) + base.layers))
        if len({_identity(v) for v in values}) != 1:
            raise ValueError('conflicting parameter typedefs')
        return values[0]
    link = registry.resolve(file, name, root)
    if link.status not in ('STRUCT', 'INTEGER', 'SCALAR', 'CONDITIONAL') or not (link.definition or link.integer_type):
        raise ValueError('parameter type is unresolved')
    cv = link.qualifiers or ((),)
    return ParameterType(link, cv[0], tuple(('pointer', q) for q in reversed(cv[1:])))


def parameter_type(registry, file, words, root):
    from .bitfield_catalog import _CV, _STORAGE, _type_spelling
    for end in range(len(words), 0, -1):
        spec = words[:end]
        spelling = ' '.join(w for w in spec if w not in _CV | _STORAGE)
        if not _type_spelling(spelling):
            continue
        try:
            name, layers = _declarator(words[end:])
            base = _qualify(_resolve(registry, file, spelling, root), set(spec) & _CV)
            # The outermost array bound is not part of a function parameter type.
            if layers and layers[0][0] == 'array':
                layers = (('pointer', ()), *layers[1:])
            value = replace(base, layers=_layers(registry, file, layers) + base.layers)
            if value.layers and value.layers[0][0] == 'array':
                value = replace(value, layers=(('pointer', ()), *value.layers[1:]))
            return name, value
        except (ValueError, RecursionError):
            continue
    raise ValueError('unsupported parameter declaration')


def signature(registry, function, root):
    result = []
    try:
        for words in function.raw_parameters:
            _, value = parameter_type(registry, function.origin.file, words, root)
            if value.layers and value.layers[0][0] == 'pointer':
                value = replace(value, layers=(('pointer', ()), *value.layers[1:]))
            elif not value.layers:
                value = replace(value, qualifiers=())
            result.append(_identity(value))
    except (ValueError, RecursionError):
        return None
    return tuple(result)


def array_parameter_binding(registry, function, index, parameter, root):
    """Adapt an array-pointer parameter to the existing indexed-member resolver."""
    from .bitfield_catalog import VariableDeclaration, VariableBinding
    try:
        name, value = parameter_type(registry, function.origin.file, function.raw_parameters[index], root)
    except (ValueError, RecursionError, IndexError):
        return None
    layers = list(value.layers)
    if not name or not any(kind == 'array' for kind, _ in layers):
        return None
    if not layers or layers.pop(0)[0] != 'pointer':
        return None
    arrays = ['']
    while layers and layers[0][0] == 'array':
        _, extent = layers.pop(0)
        arrays.append('' if extent is None else str(extent))
    if any(kind != 'pointer' for kind, _ in layers):
        return None
    qualifiers = (value.qualifiers, *(cv for _, cv in reversed(layers)))
    link = replace(value.link, pointer_depth=len(layers), qualifiers=qualifiers)
    declaration = VariableDeclaration(name, parameter.spelling if parameter else '<array parameter>',
                                     parameter.origin if parameter else function.origin,
                                     tuple(arrays), qualifiers, '', 'PARAMETER', function.conditional,
                                     link.definition)
    return VariableBinding(link.status, (declaration,), link)
