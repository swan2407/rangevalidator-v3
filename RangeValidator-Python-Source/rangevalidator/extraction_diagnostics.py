"""Observations of sheet/table extraction, separate from comparison decisions.

These observations never feed the candidate, REVIEW, or Writer paths. Format names
follow the supplied company samples: A = Input/Output/Property/Name/Data Type/Scope;
B = sectioned Name/Type/Range. The production engine uses the opposite labels.
"""
from collections import Counter
from time import perf_counter

from .a_extraction import scan_a
from .b_extraction import scan_b
from .engine import is_function_sheet


def diagnose_workbook(book, results=()):
    """Report observed cells and whether the unchanged engine produced each result."""
    started = perf_counter()
    by_result = {(r.sheet, r.cell, r.expression): r for r in results if r.file == book.path}
    errors = book.read_errors
    failed_names = {item.get("sheet_name") for item in errors if item.get("sheet_name")}
    sheets = []
    for sheet in book.sheets:
        included = is_function_sheet(sheet.name)
        partial = book.partial and (not errors or any(
            not item.get("sheet_name") or item.get("sheet_name") == sheet.name for item in errors))
        entry = dict(sheet=sheet.name, function_sheet=included,
                     states=["SHEET_INCLUDED" if included else "SHEET_EXCLUDED"],
                     format="UNKNOWN", tables=[], variables=[],
                     read_partial=bool(partial), headers=[],
                     variable_name_cells=0, extracted_variables=0,
                     resolved_range_cells=0, uncertain_variables=0)
        if partial:
            entry["states"].append("SHEET_READ_PARTIAL")
        if not included:
            sheets.append(entry)
            continue
        formats = set()
        for fmt, scan, target in (('A', scan_a, 'scope'), ('B', scan_b, 'range')):
            _, tables, candidates = scan(sheet)
            if not tables:
                continue
            formats.add(fmt)
            entry['tables'].extend(tables)
            for table in tables:
                entry['states'].append(table['state'])
                entry['headers'].append(dict(name=table['headers']['name'],
                    range_or_scope=table['headers'][target], format=fmt, section=table['section']))
            for candidate in candidates:
                actual = by_result.get((sheet.name, candidate['range_cell'], candidate['name']))
                entry['variables'].append(dict(candidate, file=book.path, extracted=actual is not None,
                    comparison_status=actual.status if actual else None,
                    comparison_reason=actual.reason if actual else None))
        entry["format"] = next(iter(formats)) if len(formats) == 1 else "MIXED" if formats else "UNKNOWN"
        if not entry["tables"]:
            entry["states"].append("HEADER_MISSING")
        entry["states"].append("FORMAT_" + entry["format"] + "_CANDIDATE"
                               if entry["format"] in ("A", "B") else "FORMAT_UNKNOWN")
        entry["variable_name_cells"] = len(entry["variables"])
        entry["extracted_variables"] = sum(v["extracted"] for v in entry["variables"])
        entry["resolved_range_cells"] = sum(v["range_cell"] is not None for v in entry["variables"])
        entry["uncertain_variables"] = sum(not v["extracted"] for v in entry["variables"])
        sheets.append(entry)
    for name in sorted(failed_names - {s["sheet"] for s in sheets}):
        sheets.append(dict(sheet=name, function_sheet=is_function_sheet(name),
                           states=["SHEET_READ_FAILED"], format="UNKNOWN", tables=[], variables=[],
                           read_partial=True, headers=[], variable_name_cells=0,
                           extracted_variables=0, resolved_range_cells=0, uncertain_variables=0))
    observed = {item['sheet']: item for item in book.sheet_observations}
    for observation in book.sheet_observations:
        if not observation['included']:
            sheets.append(dict(sheet=observation['sheet'], function_sheet=False,
                states=list(observation['states']), format='UNKNOWN', tables=[], variables=[],
                read_partial=False, headers=[], variable_name_cells=0, extracted_variables=0,
                resolved_range_cells=0, uncertain_variables=0))
    counts = Counter()
    for sheet in sheets:
        if sheet['sheet'] in observed:
            sheet['states'] = list(dict.fromkeys(sheet['states'] + observed[sheet['sheet']]['states']))
        counts["function_candidates" if sheet["function_sheet"] else "excluded_sheets"] += 1
        counts["body_read_success"] += ("SHEET_BODY_READ_COMPLETED" in sheet['states'] if observed else
                                       "SHEET_READ_FAILED" not in sheet["states"])
        counts["body_read_failed"] += any(state in sheet["states"] for state in
                                           ("SHEET_READ_FAILED", "SHEET_BODY_READ_FAILED"))
    counts["sheet_name_failures"] = sum(not item.get("sheet_name") for item in errors)
    timings = dict(book.read_timings)
    timings["diagnostic_scan_seconds"] = perf_counter() - started
    return dict(file=book.path, backend=book.backend,
                total_sheets=book.source_sheet_count if book.source_sheet_count is not None else len(sheets),
                counts=dict(counts), timings=timings, sheets=sheets,
                read_errors=errors)


def diagnose_session(session):
    reports = []
    for path in session.details:
        book = session.books.get(path)
        if book is None:
            reports.append(dict(file=path, backend=session.readers.get(path, ""),
                                total_sheets=None, counts={}, timings={}, sheets=[],
                                read_errors=[e for e in session.errors if e.get("file") == path],
                                status="FILE_READ_FAILED"))
        else:
            reports.append(diagnose_workbook(book, session.results))
    return reports
