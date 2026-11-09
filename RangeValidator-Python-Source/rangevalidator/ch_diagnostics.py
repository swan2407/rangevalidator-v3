"""Opt-in, read-only evidence for the existing C/H comparison pipeline."""
from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path
import re

from .analysis import analyze, identifier, matching_names
from .common import normalize_member_expression
from .bitfield_catalog import _DIRECTIVE, _collect_units
from .engine import MEMBER_ACCESS, _catalog_match, _input_source, catalog_headers, compare_book, source_context
from .extraction_diagnostics import diagnose_workbook


def _location(origin):
    return asdict(origin) if origin else None


def _bounds(field):
    return [field.minimum, field.maximum] if field.minimum is not None else None


def _link(link):
    return asdict(link) if link else None


class CHDiagnostics:
    def __init__(self, paths, *, encoding='utf-8-sig', build_options=None):
        self.paths = sorted({str(Path(path).resolve()) for path in paths})
        self.catalog = catalog_headers(self.paths, encoding)
        self.bitfields = self.catalog.bitfields
        self.options = {str(Path(path).resolve()): list(args)
                        for path, args in (build_options or {}).items()}
        sources, self.directives, self.syntax_issues = {}, {}, {}
        self.files, self.records, self.typedefs, self.variables, self.failures = [], [], [], [], []
        for path in self.paths:
            try:
                raw = Path(path).read_bytes()
            except OSError as exc:
                self.failures.append(dict(file=path, line=None, code='SOURCE_READ_ERROR', reason=str(exc)))
                continue
            for codec in dict.fromkeys((encoding, 'utf-8-sig', 'cp949', 'latin-1')):
                try:
                    source = raw.decode(codec)
                    break
                except UnicodeDecodeError:
                    continue
            sources[path] = source
            # The catalog collector does not return its parser-level reviews.
            # Recover corroborating syntax evidence from the existing legacy parser.
            self.syntax_issues[path] = [d for d in analyze(source, path).reviews
                                        if d.code == 'UNSUPPORTED_SYNTAX']
            self.directives[path] = [dict(line=number, directive=m[1], argument=m[2])
                                     for number, line in enumerate(source.splitlines(), 1)
                                     if (m := _DIRECTIVE.fullmatch(line))]
        self.units = _collect_units(sources, set(self.paths))
        self.raw_records = {origin: raw for unit in self.units.values()
                            for origin, (raw, _) in unit.definitions.items()}
        self._inventory()

    def _inventory(self):
        typedef_names = defaultdict(set)
        tag_counts = Counter(raw.tag for raw in self.raw_records.values() if raw.tag)
        for path, unit in self.units.items():
            for name in unit.types:
                typedef_names[name].add(path)
        for path in self.paths:
            unit = self.units.get(path)
            if unit is None:
                self.files.append(dict(file=path, read=False))
                continue
            records, typedefs, variables = [], [], []
            for origin, (raw, _) in unit.definitions.items():
                final = self.bitfields.definitions.get(origin)
                fields = [] if final is None else [
                    dict(member=f.declaration.member, line=f.declaration.line,
                         declared_type=f.declaration.declared_type, width=f.declaration.width or None,
                         signed=f.declaration.signed, range=_bounds(f.declaration),
                         width_expression=f.width_expression, status=f.status)
                    for f in final.fields]
                ordinary = {name: raw.members[name] for name in dict.fromkeys(raw.ordinary)}
                nested = {name: _location(link.definition) for name, spelling in ordinary.items()
                          if (link := self.bitfields.lookup(path, spelling)).status == 'STRUCT'}
                records.append(dict(origin=_location(origin), tag=raw.tag, name=raw.name, defined=True,
                                    aliases=list(final.aliases) if final else [],
                                    status=final.status if final else 'CATALOG_OMISSION',
                                    duplicate_tag=tag_counts[raw.tag] > 1,
                                    ordinary_members=ordinary, typed_nested_members=nested,
                                    inline_nested_members=list(raw.nested), bitfields=fields))
            for name, declarations in unit.types.items():
                link = self.bitfields.lookup(path, name)
                if name.startswith('struct '):
                    if link.status == 'INCOMPLETE':
                        records.append(dict(origin=_location(declarations[0].origin), tag=name[7:],
                                            defined=False, aliases=[], status=link.status,
                                            ordinary_members={}, inline_nested_members=[], bitfields=[],
                                            duplicate_tag=len(typedef_names[name]) > 1))
                    continue
                for decl in declarations:
                    typedefs.append(dict(name=name, origin=_location(decl.origin),
                                         target=_location(decl.target) if not isinstance(decl.target, str) else decl.target,
                                         conditional=decl.conditional, link=_link(link),
                                         duplicate_name=len(typedef_names[name]) > 1 or len(declarations) > 1))
            for decl in unit.variables:
                variables.append(self._variable(decl, 'GLOBAL', None,
                                                self.bitfields.lookup_variable(path, decl.name)))
            for fn in unit.functions:
                function = self.bitfields.lookup_function(path, fn.name)
                for scope, declarations in (
                    ('PARAMETER' if fn.definition else 'PROTOTYPE_PARAMETER', fn.parameters),
                    ('LOCAL', [decl for decl, _ in fn.locals]),
                ):
                    for decl in declarations:
                        if decl is None:
                            self.failures.append(dict(file=path, line=fn.origin.line, function=fn.name,
                                                      code='UNSUPPORTED_PARAMETER_DECLARATION'))
                            continue
                        binding = function.variables.get(decl.name) if fn.definition else None
                        variables.append(self._variable(decl, scope, fn.name, binding, function.status))
            self.records.extend(records)
            self.typedefs.extend(typedefs)
            self.variables.extend(variables)
            failures = [dict(asdict(d), evidence='CATALOG_COLLECTOR') for d in unit.reviews]
            failures.extend(dict(asdict(d), evidence='LEGACY_SYNTAX_PASS')
                            for d in self.syntax_issues[path] if d not in unit.reviews)
            self.failures.extend(failures)
            unmatched_includes = []
            for item in self.directives[path]:
                if item['directive'] == 'include':
                    quoted = re.fullmatch(r'"([^"]+)"', item['argument'].strip())
                    if not quoted or str((Path(path).parent / quoted[1]).resolve()) not in unit.includes:
                        unmatched_includes.append(dict(item, code='INCLUDE_NOT_LINKED'))
            self.files.append(dict(file=path, read=True, structures=sum(r['defined'] for r in records),
                                   typedefs=len(typedefs), bitfields=sum(len(r['bitfields']) for r in records),
                                   ordinary_members=sum(len(r['ordinary_members']) for r in records),
                                   variables=len(variables), parser_diagnostics=len(failures),
                                   incomplete_declarations=len(
                                       {d['line'] for d in failures}
                                       | {v['origin']['line'] for v in variables
                                          if v['binding_status'] not in ('STRUCT', 'INTEGER', 'SCALAR',
                                                                         'PROTOTYPE_ONLY')}
                                       | {t['origin']['line'] for t in typedefs
                                          if t['link']['status'] not in ('STRUCT', 'INTEGER', 'SCALAR')}),
                                   conditional_declarations=[_location(d) for d in sorted(unit.conditional_declarations)],
                                   include_links=unit.includes, directives=self.directives[path],
                                   unmatched_include_observations=unmatched_includes,
                                   build_options=self._option_status(path)))

    @staticmethod
    def _variable(decl, scope, function, binding, function_status=None):
        return dict(**asdict(decl), scope=scope, function=function, function_status=function_status,
                    pointer_depth=decl.pointer_depth,
                    binding_status=binding.status if binding else
                    'PROTOTYPE_ONLY' if scope == 'PROTOTYPE_PARAMETER' else 'MISSING',
                    type_link=_link(binding.type_link) if binding else None)

    def _option_status(self, source):
        return ('BUILD_OPTIONS_PROVIDED_NOT_APPLIED' if source in self.options
                else 'BUILD_OPTIONS_NOT_PROVIDED')

    def _binding_code(self, source, binding):
        if binding.status == 'CONDITIONAL':
            return 'PREPROCESSOR_UNRESOLVED'
        if binding.status == 'AMBIGUOUS':
            return 'AMBIGUOUS_TYPE'
        if binding.status == 'INCOMPLETE':
            return 'STRUCT_NOT_FOUND'
        if binding.status == 'MISSING':
            return 'VARIABLE_DECL_NOT_FOUND'
        if any(d.kind == 'UNSUPPORTED' or d.spelling.startswith('<') for d in binding.declarations):
            return 'UNSUPPORTED_DECLARATION'
        for decl in binding.declarations:
            if decl.spelling.startswith('struct '):
                return 'STRUCT_NOT_FOUND'
            scope = self.bitfields.source_scopes.get(source or decl.origin.file, {})
            if any(decl.spelling in self.units[path].types for path in scope if path in self.units):
                return 'TYPEDEF_UNRESOLVED'
        if any(d.arrays for d in binding.declarations):
            return 'ARRAY_ELEMENT_TYPE_UNRESOLVED'
        if any(d.pointer_depth for d in binding.declarations):
            return 'POINTER_TYPE_UNRESOLVED'
        return 'VARIABLE_TYPE_UNRESOLVED'

    def _width_code(self, field):
        if field.status == 'CONDITIONAL':
            return 'PREPROCESSOR_UNRESOLVED'
        if field.status == 'AMBIGUOUS':
            return 'AMBIGUOUS_MEMBER'
        if field.status == 'WIDTH_UNRESOLVED':
            visible = self.bitfields.source_scopes.get(field.declaration.file, {})
            names = re.findall(r'\b[A-Za-z_]\w*\b', field.width_expression)
            macros = [d for path in visible if path in self.units
                      for name in names for d in self.units[path].macros.get(name, ())]
            if any(d.conditional or visible[d.origin.file] for d in macros):
                return 'PREPROCESSOR_UNRESOLVED'
            return 'BITFIELD_WIDTH_UNRESOLVED'
        return 'BITFIELD_TYPE_UNRESOLVED'

    def trace_expression(self, expression, source, function_name):
        """Explain production gates, retaining observed members when a gate excludes them."""
        source = str(Path(source).resolve()) if source is not None else None
        expression = normalize_member_expression(expression)
        out = dict(source=source, root=None, member_path=[], candidates=[], other_declarations=[],
                   type_link=None, bitfield=None, range=None, production_match=None)

        def finish(code):
            return dict(out, code=code, build_options=self._option_status(source)
                        if code == 'PREPROCESSOR_UNRESOLVED' else None)

        parsed = re.fullmatch(r'\s*([A-Za-z_]\w*)((?:\s*\[[^\]\r\n]*\])*)'
                              r'((?:\s*(?:\.|->)\s*[A-Za-z_]\w*'
                              r'(?:\s*\[[^\]\r\n]*\])*)+)\s*', expression, re.ASCII)
        if not parsed or not identifier(parsed[1]):
            return finish('EXPRESSION_UNSUPPORTED')
        name, indexes, tail = parsed.groups()
        parts = re.findall(r'(->|\.)\s*([A-Za-z_]\w*)((?:\s*\[[^\]\r\n]*\])*)', tail)
        access = MEMBER_ACCESS.fullmatch(expression)
        out['production_gate'] = ('SINGLE_MEMBER_SYNTAX_REQUIRED' if not access else
                                  'MEMBER_NAME_NOT_IN_BITFIELD_INDEX'
                                  if not matching_names(self.catalog.bitfield_members, access['member']) else 'MEMBER_INDEX_HIT')
        out.update(root=name, root_indexes=indexes, member_path=[
            dict(operator=op, member=member, indexes=subscripts) for op, member, subscripts in parts])
        function = self.bitfields.lookup_function(source, function_name)
        out['function_status'] = function.status
        if source is None and function.status in ('RESOLVED', 'CONDITIONAL'):
            source = function.file
            out['source'] = source
        binding = self.bitfields.lookup_variable(source, name, function.variables)
        if binding.status == 'MISSING' and source is not None:
            binding = self.bitfields.lookup_variable(None, name)
        out['candidates'] = [asdict(d) for d in binding.declarations]
        if binding.status != 'STRUCT':
            out['other_declarations'] = [asdict(d) for d in self.bitfields.variables.get(name, ())
                                         if d not in binding.declarations]
        if function.status not in ('RESOLVED', 'MISSING', 'CONDITIONAL', 'MULTIPLE'):
            return finish('AMBIGUOUS_FUNCTION')
        out.update(binding_status=binding.status, type_link=_link(binding.type_link))
        if access:
            status, reason, matched = _catalog_match(self.catalog, access, source, function)
            out['production_match'] = dict(status=status, reason=reason)
            if status == 'TYPE_MATCH' and matched is not None:
                out.update(bitfield=True, field=asdict(matched), range=_bounds(matched))
                return finish('TYPE_MATCH')
            if status == 'SKIP' and reason:
                out['bitfield'] = False
                return finish('NOT_BITFIELD')
            if function.alternatives and status == 'REVIEW':
                return finish('AMBIGUOUS_FUNCTION' if all(s.status == 'AMBIGUOUS' for s in function.alternatives)
                              else 'AMBIGUOUS_MEMBER')
        if binding.status in ('INTEGER', 'SCALAR'):
            return finish('NON_STRUCT_VARIABLE')
        if binding.status != 'STRUCT':
            return finish(self._binding_code(source, binding))
        origin = binding.type_link.definition
        record = self.bitfields.definitions.get(origin)
        raw = self.raw_records.get(origin)
        if record is None or raw is None:
            return finish('STRUCT_NOT_FOUND')
        if record.status != 'RESOLVED':
            return finish('PREPROCESSOR_UNRESOLVED' if record.status == 'CONDITIONAL'
                          else 'AMBIGUOUS_TYPE')
        if not access:
            out['first_member_type'] = raw.members.get(parts[0][1])
            return finish('MEMBER_PATH_UNRESOLVED')
        member = access['member']
        names = matching_names(set(record.ordinary_members) | {f.declaration.member for f in record.fields}, member)
        if len(names) > 1:
            return finish('AMBIGUOUS_MEMBER')
        if names:
            member = names[0]
        fields = [f for f in record.fields if f.declaration.member == member]
        if len(fields) > 1:
            return finish('AMBIGUOUS_MEMBER')
        if not fields:
            issues = [*self.units[origin.file].reviews, *self.syntax_issues[origin.file]]
            if member in raw.ordinary and not any(d.line == origin.line for d in issues):
                out['bitfield'] = False
                return finish('NOT_BITFIELD')
            # The existing collector can drop declarations without producing a warning.
            return finish('PARSE_INCOMPLETE' if issues else 'MEMBER_NOT_FOUND')
        field = fields[0]
        out.update(bitfield=True, field=asdict(field.declaration), range=_bounds(field.declaration))
        status, reason, matched = _catalog_match(self.catalog, access, source, function)
        out['production_match'] = dict(status=status, reason=reason)
        if status == 'TYPE_MATCH':
            return finish('TYPE_MATCH')
        if matched is not None and field.status != 'RESOLVED':
            return finish(self._width_code(field))
        return finish('ARRAY_ELEMENT_TYPE_UNRESOLVED' if indexes or binding.declarations[0].arrays
                      else 'POINTER_TYPE_UNRESOLVED')

    def workbook(self, book, relative_path=None):
        results, coverage = compare_book(None, self.catalog, book, relative_path)
        source = _input_source(self.bitfields, source_context(relative_path or Path(book.path).name))
        rows = []
        for sheet in diagnose_workbook(book, results)['sheets']:
            for row in sheet['variables']:
                trace = self.trace_expression(row['name'], source, row['function'])
                code = trace['code']
                if code == 'TYPE_MATCH':
                    if row['comparison_status'] == 'REVIEW':
                        code = 'EXCEL_RANGE_REVIEW'
                    elif row['comparison_status'] is None:
                        code = 'EXCEL_ROW_EXCLUDED'
                rows.append(dict(row, trace=trace, code=code))
        return dict(file=book.path, backend=book.backend, coverage=coverage, rows=rows,
                    production_counts=dict(Counter(r.status for r in results)),
                    review_reasons=dict(Counter(row['code'] for row in rows
                                                if row['comparison_status'] == 'REVIEW')),
                    diagnostic_reasons=dict(Counter(row['code'] for row in rows)))

    def report(self, workbooks=()):
        bound = [v for v in self.variables if v['scope'] != 'PROTOTYPE_PARAMETER']
        rows = [row for book in workbooks for row in book['rows']]
        variable_counts = Counter(v['scope'].lower() for v in bound)
        variable_counts.update(type_links_success=sum(v['binding_status'] == 'STRUCT' for v in bound),
                               type_links_failure=sum(v['binding_status'] not in ('STRUCT', 'INTEGER', 'SCALAR')
                                                      for v in bound),
                               non_struct=sum(v['binding_status'] in ('INTEGER', 'SCALAR') for v in bound))
        return dict(files=self.files, structures=self.records, typedefs=self.typedefs,
                    variables=self.variables, parse_failures=self.failures, workbooks=workbooks,
                    counts=dict(files=len(self.files), structures=sum(r['defined'] for r in self.records),
                                typedefs=len(self.typedefs), bitfields=sum(len(r['bitfields']) for r in self.records),
                                variables=dict(variable_counts),
                                excel_expressions_matched=sum(r['trace']['code'] == 'TYPE_MATCH' for r in rows),
                                review_reasons=dict(Counter(r['code'] for r in rows
                                                           if r['comparison_status'] == 'REVIEW')),
                                rows_without_production_result=sum(r['comparison_status'] is None for r in rows)),
                    build_options_supplied=self.options,
                    catalog_omission_suspicions=[r for r in rows if r['code'] in
                        ('STRUCT_NOT_FOUND', 'MEMBER_NOT_FOUND', 'PARSE_INCOMPLETE', 'VARIABLE_DECL_NOT_FOUND')],
                    build_option_suspicions=[r for r in rows if r['trace']['build_options']])
