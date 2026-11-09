"""Isolated libclang experiment. No comparison engine or Writer integration."""
import ctypes
from importlib import metadata
import json
from pathlib import Path
import re


SIGNED = {'CHAR_S', 'SCHAR', 'SHORT', 'INT', 'LONG', 'LONGLONG', 'INT128'}
UNSIGNED = {'CHAR_U', 'UCHAR', 'USHORT', 'UINT', 'ULONG', 'ULONGLONG', 'UINT128', 'BOOL'}
ARRAYS = {'CONSTANTARRAY', 'INCOMPLETEARRAY', 'VARIABLEARRAY', 'DEPENDENTSIZEDARRAY'}
RECORDS = {'STRUCT_DECL', 'UNION_DECL'}


def load_clang(dll=None):
    from clang import cindex
    if dll is not None:
        path = str(Path(dll).resolve(strict=True))
        if cindex.Config.loaded:
            if Path(cindex.conf.get_filename()).resolve() != Path(path):
                raise RuntimeError('Use a separate process to switch the loaded libclang DLL')
        else:
            cindex.Config.set_library_file(path)
    return cindex


def library_info(dll=None):
    clang = load_clang(dll)
    try:
        clang.Index.create()  # cindex compatibility check stays enabled.
    except clang.LibclangError as exc:
        raise RuntimeError(f'libclang load or binding compatibility check failed: {exc}') from exc
    library = clang.conf.lib
    path = Path(clang.conf.get_filename()).resolve()
    if not path.is_file():
        raise RuntimeError('Pass the full libclang library path so the loaded file can be identified')

    class CXString(ctypes.Structure):
        _fields_ = [('data', ctypes.c_void_p), ('private_flags', ctypes.c_uint)]

    lib = ctypes.CDLL(str(path), handle=library._handle)
    lib.clang_getClangVersion.restype = CXString
    lib.clang_getCString.argtypes, lib.clang_getCString.restype = [CXString], ctypes.c_char_p
    lib.clang_disposeString.argtypes, lib.clang_disposeString.restype = [CXString], None
    version = lib.clang_getClangVersion()
    try:
        text = lib.clang_getCString(version).decode('utf-8', 'replace')
    finally:
        lib.clang_disposeString(version)
    bindings = {name: metadata.version(name) for name in metadata.packages_distributions().get('clang', [])}
    return dict(version=text, binding_distributions=bindings, binding_file=clang.__file__,
                library=str(path), library_bytes=path.stat().st_size, compatibility_check=True)


def compilation_jobs(database):
    """Read argv-form C compilation entries. Refuse shell strings and duplicate variants."""
    path = Path(database).resolve()
    entries = json.loads(path.read_text(encoding='utf-8'))
    jobs = {}
    for entry in entries:
        directory = Path(entry['directory']).resolve()
        source = (directory / entry['file']).resolve()
        if source in jobs:
            raise ValueError(f'Multiple compile variants for {source.name}; select one explicitly')
        args = entry.get('arguments')
        if not isinstance(args, list) or not args or not all(isinstance(v, str) for v in args):
            raise ValueError('compile_commands requires an arguments array in this experiment')
        flags = [f'-working-directory={directory}']
        index = 1
        while index < len(args):
            arg = args[index]
            index += 1
            if arg == '-o':
                if index >= len(args):
                    raise ValueError('Missing output argument')
                index += 1
            elif arg != '-c' and (arg.startswith('-') or (directory / arg).resolve() != source):
                flags.append(arg)
        jobs[source] = flags
    return jobs


def unsupported_options(args):
    separate = {'-I', '-D', '-U', '-isystem', '-target', '-resource-dir', '-x', '--sysroot'}
    prefixes = ('-I', '-D', '-U', '-std=', '--target=', '-target=', '-working-directory=',
                '-resource-dir=', '--sysroot=')
    simple = {'-Wall', '-Wextra', '-pedantic', '-fms-extensions', '-fshort-enums',
              '-fsigned-char', '-funsigned-char', '-nostdinc'}
    bad = []
    index = 0
    while index < len(args):
        arg = args[index]
        index += 1
        if arg in separate:
            if index >= len(args):
                bad.append(arg)
            else:
                operand = args[index]
                index += 1
                if operand.startswith(('-', '@')) or (arg == '-x' and operand != 'c'):
                    bad.extend((arg, operand))
        elif arg not in simple and not arg.startswith(prefixes):
            bad.append(arg)
    return bad


class _Inventory:
    def __init__(self, root, trusted, directory):
        self.root, self.trusted = root, trusted
        self.directory = directory
        self.records = {}
        self.aliases = []
        self.variables = []
        self.accesses = []
        self.sources = {}

    def location(self, loc):
        if loc.file is None:
            return None
        file = Path(loc.file.name)
        try:
            relative = (self.directory / file).resolve().relative_to(self.root).as_posix()
        except (ValueError, OSError):
            return None
        return dict(file=relative, line=loc.line, column=loc.column, offset=loc.offset)

    def record_id(self, cursor):
        if cursor is None or cursor.kind.name not in RECORDS:
            return None
        definition = cursor.get_definition() or cursor
        loc = self.location(definition.location)
        return f'{loc["file"]}:{loc["line"]}:{loc["column"]}' if loc else None

    def type_info(self, typ):
        current = typ.get_canonical()
        canonical = current.spelling
        pointers, arrays, layers = 0, [], []
        element = None
        while True:
            kind = current.kind.name
            cv = [name for name, present in (
                ('const', current.is_const_qualified()), ('volatile', current.is_volatile_qualified()),
                ('restrict', current.is_restrict_qualified())) if present]
            layers.append(dict(kind=kind, qualifiers=cv))
            if kind == 'POINTER':
                pointers += 1
                current = current.get_pointee()
            elif kind in ARRAYS:
                size = current.get_array_size()
                arrays.append(size if size >= 0 else None)
                current = current.get_array_element_type()
                if element is None:
                    element = current.spelling
            else:
                break
        decl = current.get_declaration()
        identity = self.record_id(decl)
        resolved = bool(self.trusted and identity and decl.get_definition() is not None)
        return dict(original=typ.spelling, canonical=canonical, base_kind=current.kind.name,
                    record_id=identity, pointer_depth=pointers, array_dimensions=arrays,
                    array_element_type=element, layers=layers, record_link_trusted=resolved)

    def field(self, cursor):
        bit = cursor.is_bitfield()
        width = cursor.get_bitfield_width() if bit else None
        width = width if width is None or width >= 0 else None
        typ = cursor.type.get_canonical()
        if typ.kind.name == 'ENUM':
            typ = typ.get_declaration().enum_type.get_canonical()
        integer = typ.kind.name
        signed = True if integer in SIGNED else False if integer in UNSIGNED else None
        size = typ.get_size()
        valid = self.trusted and bit and width is not None and width > 0 and signed is not None and size > 0 and width <= size * 8
        bounds = ([-(1 << (width - 1)) if signed else 0, (1 << (width - int(signed))) - 1]
                  if valid else None)
        reason = ('ZERO_WIDTH_PADDING' if width == 0 else 'UNCONFIRMED_TRANSLATION_UNIT'
                  if not self.trusted else 'WIDTH_OR_SIGNEDNESS_UNRESOLVED')
        return dict(name=cursor.spelling, location=self.location(cursor.location), type=self.type_info(cursor.type),
                    bitfield=bit, width=width, signed=signed if valid else None, range=bounds,
                    status='RESOLVED' if valid else 'BITFIELD_UNRESOLVED' if bit else 'NOT_BITFIELD',
                    reason=reason if bit and not valid else '')

    def access(self, cursor, function, loc):
        field = cursor.referenced
        if field is None or field.kind.name != 'FIELD_DECL':
            return
        base = cursor
        while base.kind.name != 'DECL_REF_EXPR':
            if base.kind.name not in ('MEMBER_REF_EXPR', 'ARRAY_SUBSCRIPT_EXPR', 'UNEXPOSED_EXPR',
                                      'PAREN_EXPR', 'UNARY_OPERATOR'):
                return
            base = next(iter(base.get_children()), None)
            if base is None:
                return
        variable = base.referenced
        if variable is None or variable.kind.name not in ('VAR_DECL', 'PARM_DECL'):
            return
        start = self.location(cursor.extent.start)
        end = self.location(cursor.extent.end)
        if not start or not end or start['file'] != end['file']:
            return
        path = start['file']
        if path not in self.sources:
            self.sources[path] = (self.root / path).read_bytes()
        expr = self.sources[path][start['offset']:end['offset']].decode('utf-8', 'replace')
        member = self.field(field)
        binding = self.type_info(variable.type)
        self.accesses.append(dict(expression=expr, function=function, location=loc,
            variable=variable.spelling, variable_location=self.location(variable.location),
            variable_type=binding, member=field.spelling, member_record_id=self.record_id(field.semantic_parent),
            member_location=member['location'], bitfield=member['bitfield'],
            width=member['width'], range=member['range'],
            trusted=binding['record_link_trusted'] and bool(self.record_id(field.semantic_parent))))

    def visit(self, cursor, function=None, scope=None, parent_record=None):
        loc = self.location(cursor.location)
        if loc is None:
            return  # Included system declarations remain outside the project inventory.
        kind = cursor.kind.name
        if kind == 'FUNCTION_DECL':
            function = dict(name=cursor.spelling, declaration=loc)
        if kind == 'COMPOUND_STMT':
            scope = [self.location(cursor.extent.start), self.location(cursor.extent.end)]
        if kind in RECORDS:
            identity = self.record_id(cursor)
            row = self.records.setdefault(identity, dict(
                id=identity, tag=cursor.spelling if not cursor.is_anonymous() else '',
                kind=kind, declarations=[], definition=None, fields=[], aliases=[],
                parent_record_id=parent_record))
            row['declarations'].append(loc)
            if cursor.is_definition():
                row['definition'] = loc
                row['fields'] = [self.field(child) for child in cursor.get_children()
                                 if child.kind.name == 'FIELD_DECL']
            parent_record = identity
        elif kind == 'TYPEDEF_DECL':
            typ = self.type_info(cursor.underlying_typedef_type)
            self.aliases.append(dict(name=cursor.spelling, location=loc, type=typ))
        elif kind in ('VAR_DECL', 'PARM_DECL'):
            self.variables.append(dict(name=cursor.spelling, location=loc, function=function,
                scope=scope, storage=cursor.storage_class.name,
                kind='PARAMETER' if kind == 'PARM_DECL' else 'LOCAL' if function else 'GLOBAL',
                type=self.type_info(cursor.type)))
        elif kind == 'MEMBER_REF_EXPR':
            self.access(cursor, function, loc)
        for child in cursor.get_children():
            self.visit(child, function, scope, parent_record)

    def finish(self):
        for alias in self.aliases:
            typ = alias['type']
            if typ['record_id'] in self.records and not typ['pointer_depth'] and not typ['array_dimensions']:
                self.records[typ['record_id']]['aliases'].append(alias['name'])
        for record in self.records.values():
            record['status'] = 'RESOLVED' if self.trusted and record['definition'] else 'TYPE_UNRESOLVED'
        return dict(records=list(self.records.values()), typedefs=self.aliases, variables=self.variables,
                    member_accesses=self.accesses)


def analyze(source, *, root, args=(), config_confirmed=False, dll=None):
    """Inventory one C translation unit under explicitly declared build settings."""
    source, root = Path(source).resolve(), Path(root).resolve()
    relative = source.relative_to(root).as_posix()
    flags = list(args)
    directory = next((Path(arg.partition('=')[2]).resolve() for arg in reversed(flags)
                      if arg.startswith('-working-directory=')), Path.cwd())
    result = dict(source=relative, args=flags, ast_created=False, trusted=False, diagnostics=[],
                  records=[], typedefs=[], variables=[], member_accesses=[],
                  missing_headers=[], unresolved_identifiers=[], incomplete_reasons=[])
    bad = unsupported_options(flags)
    if bad:
        return dict(result, status='PARSE_INCOMPLETE', unsupported_options=bad,
                    incomplete_reasons=['UNSUPPORTED_OPTIONS'])
    result['unsupported_options'] = []
    clang = load_clang(dll)
    try:
        tu = clang.Index.create().parse(str(source), args=['-x', 'c', *flags])
    except clang.TranslationUnitLoadError as exc:
        return dict(result, status='PARSE_FAILED', diagnostics=[dict(severity=4, message=str(exc))],
                    incomplete_reasons=['TRANSLATION_UNIT_LOAD_ERROR'])
    result['ast_created'] = True
    location = _Inventory(root, False, directory).location
    diags = [dict(severity=d.severity, message=d.spelling, location=location(d.location),
                  option=d.option) for d in tu.diagnostics]
    reasons = result['incomplete_reasons']
    if not config_confirmed or not any(a.startswith(('--target=', '-target=')) or a == '-target' for a in flags):
        reasons.append('BUILD_CONFIGURATION_UNCONFIRMED')
    if any(d['severity'] >= 3 for d in diags):
        reasons.append('CLANG_ERRORS')
    if any(re.search(r'unsupported option|argument unused|unknown argument|ignored', d['message'], re.I)
           for d in diags):
        reasons.append('BUILD_OPTION_UNSUPPORTED_OR_IGNORED')
    result.update(diagnostics=diags, trusted=not reasons,
                  status='PARSE_INCOMPLETE' if reasons else 'PARSE_WARNING' if any(d['severity'] == 2 for d in diags) else 'PARSE_OK',
                  missing_headers=[m.group(1) for d in diags if (m := re.search(r"['<]([^'>]+)['>] file not found", d['message']))],
                  unresolved_identifiers=[m.group(1) for d in diags if (m := re.search(r"undeclared identifier '([^']+)'", d['message']))])
    inventory = _Inventory(root, result['trusted'], directory)
    for cursor in tu.cursor.get_children():
        inventory.visit(cursor)
    result.update(inventory.finish())
    return result
