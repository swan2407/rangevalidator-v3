"""Resolve a worksheet's function before selecting any row's variable."""
from pathlib import Path

from .analysis import identifier, matching_names
from .bitfield_catalog import FunctionScope


def parameter_linkage_issue(scope, variable):
    names = matching_names(scope.parameter_issues, variable)
    if names:
        return ' / '.join(scope.parameter_issues[name] for name in names)
    if scope.uncollected_parameters and not matching_names(scope.variables, variable):
        return (f"'{variable}'의 함수 내 연결을 확정하지 못함: 이름을 수집하지 못한 매개변수가 있음 "
                '/ 같은 이름의 전역 변수로 대신 연결하지 않음')
    return ''


def sheet_function_scope(catalog, source, name):
    scope = catalog.lookup_function(source, name)
    if scope.status != 'MISSING' or not name or len(name) < 24 or not identifier(name):
        return scope, name, ''
    # Excel titles are limited to 31 characters, including numbering. Search all
    # registered sources: a design document need not share its C file's name.
    names = [full for full in catalog.functions
             if len(full) > len(name) and full.casefold().startswith(name.casefold())]
    exact = [full for full in names if full.startswith(name)]
    names = exact or names
    if source and len(names) > 1:
        local = [full for full in names if any(s.file == source for s in catalog.functions[full])]
        names = local or names
    if not names:
        return scope, name, ''
    if len(names) != 1:
        shown = ', '.join(sorted(names)[:5])
        return (FunctionScope(source or '', 'AMBIGUOUS'), name,
                f"시트 함수명 '{name}'에 대응하는 함수 후보가 여러 개임: {shown}")
    full = names[0]
    resolved = catalog.lookup_function(source, full)
    # Core lookup may reject a unique function because of an unrelated file hint.
    if resolved.status in ('MISSING', 'AMBIGUOUS') and source:
        retry = catalog.lookup_function(None, full)
        if retry.status not in ('MISSING', 'AMBIGUOUS'):
            resolved = retry
    return resolved, full, f"시트 함수명 '{name}' → '{full}'"


def missing_function_local_reason(catalog, source, function_name, variable):
    """Do not substitute a global for an unresolved function's possible local."""
    if not function_name:
        return ''  # Direct catalog diagnostics may explicitly request no function.
    locations = catalog.function_declaration_locations
    candidates = [(origin, function) for spelling in matching_names(locations, variable)
                  for origin, function in locations[spelling]
                  if not source or origin.file == source]
    if not candidates:
        return ''
    shown = ', '.join(f'{function} ({Path(origin.file).name}:{origin.line})'
                      for origin, function in sorted(set(candidates))[:5])
    return (f"함수 '{function_name}'를 연결하지 못함 / '{variable}'의 매개변수·지역 선언 후보: "
            f"{shown} / 같은 이름의 전역 변수로 대신 연결하지 않음")
