"""Generate independent, disposable A/B sample-shaped XLSX and ground truth."""
from collections import Counter, defaultdict
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def _cell(row, col):
    return f"{get_column_letter(col)}{row}"


def _table(ws, truth, fmt, count, *, header_row, col, section, blank=False,
           merged=False, missing_range=False, spaced=False, duplicate=False,
           formula=False, korean_range=False):
    ws.cell(header_row - 2, col, section)
    if fmt == "A":
        offsets = (0, 1, 2, 4 if merged else 3, 6 if merged else 5)
        labels = ("Input/Output", "Property", "Name", "Data Type", "Scope")
        name_col, range_col = col + offsets[2], col + offsets[4]
    else:
        offsets = (0, 2 if merged else 1, 3 if merged else 2)
        labels = ("Name", "Type", "Range")
        name_col, range_col = col, col + offsets[2]
    for offset, label in zip(offsets, labels):
        if missing_range and label in ("Range", "Scope"):
            continue
        if korean_range and label in ("Range", "Scope"):
            label = "범위"
        if spaced and label in ("Name", "Range", "Scope"):
            label = "  " + label[:2] + "\n" + label[2:] + "  "
        ws.cell(header_row, col + offset, label)
    if merged:
        ws.merge_cells(start_row=header_row, start_column=name_col,
                       end_row=header_row, end_column=name_col + 1)
    for index in range(count):
        row = header_row + 1 + index + (index // 4 if blank else 0)
        field = 1 if duplicate and index == count - 1 else index % 200 + 1
        name = f"obj.v{field:03d}"
        value = "0~7" if index % 3 == 0 else "0~3"
        if index % 7 == 6:
            value = None
        if formula and index == 0:
            value = '=IF(1=0,"x","")'
        ws.cell(row, name_col, name)
        ws.cell(row, range_col, value)
        if fmt == "A":
            ws.cell(row, col, "Input")
            ws.cell(row, col + offsets[1], section)
            ws.cell(row, col + offsets[3], "unsigned int")
        else:
            ws.cell(row, col + offsets[1], "unsigned int")
        truth.append(dict(file=ws.parent._excel1_filename, sheet=ws.title, format=fmt,
                          section=section, name=name, name_cell=_cell(row, name_col),
                          range_cell=_cell(row, range_col), original_range=value))
    return row


def _workbook(root, filename, specifications):
    book = Workbook()
    book._excel1_filename = filename
    book.active.title = "표지"
    book.active["A1"] = "Name is a word in this description, not a variable header."
    book.create_sheet("commonparameter")["A1"] = "Common Parameter"
    truth = []
    for idx, spec in enumerate(specifications, 1):
        sheet = book.create_sheet(f"{idx}. {spec.get('title', 'Function')}" if idx % 2 else
                                  f"{idx}.{spec.get('title', 'Function')}")
        sheet["A1"] = "Name appears in a description sentence."
        base = dict(spec)
        base.pop("title", None)
        fmt = base.pop("fmt")
        groups = base.pop("groups", None)
        if groups:
            row = base.pop("header_row", 20)
            for section, count in groups:
                end = _table(sheet, truth, fmt, count, header_row=row,
                             section=section, **base)
                row = end + 5
        else:
            count = base.pop("count")
            section = base.pop("section", "함수 변수 정보" if fmt == "A" else "Global(Input)")
            _table(sheet, truth, fmt, count, section=section, **base)
    path = root / filename
    book.save(path)
    book.close()
    return path, truth


def build_suite(root):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    source = root / "excel1_types.h"
    source.write_text("typedef struct { " + " ".join(
        f"unsigned int v{i:03d}:3;" for i in range(1, 201)) +
        " } Excel1_t; extern Excel1_t obj;\n", encoding="utf-8")
    cases = {
        "a_10.xlsx": [dict(fmt="A", count=10, header_row=24, col=2, title="charger_init")],
        "b_15.xlsx": [dict(fmt="B", groups=[("Parameter(Input)", 3), ("Global(Input)", 7),
                                                ("Local", 5)], header_row=20, col=2,
                              title="calc_chg_acc")],
        "mixed_abab.xlsx": [dict(fmt=f, count=n, header_row=15 + i * 3, col=2 + i,
                                 title=f"mixed_{i}") for i, (f, n) in
                            enumerate(zip("ABAB", (3, 7, 5, 9)))],
        "mixed_baba.xlsx": [dict(fmt=f, count=n, header_row=15 + i * 2, col=3 + i,
                                 title=f"mixed_{i}") for i, (f, n) in
                            enumerate(zip("BABA", (8, 2, 6, 4)))],
        "shifted.xlsx": [dict(fmt="A", count=9, header_row=42, col=5, title="shifted_a"),
                         dict(fmt="B", count=11, header_row=53, col=6, title="shifted_b")],
        "gaps_merged.xlsx": [dict(fmt="A", count=8, header_row=24, col=2, blank=True,
                                  merged=True, title="merged_a"),
                             dict(fmt="B", groups=[("Global(Input)", 5), ("Global(Output)", 4)],
                                  header_row=28, col=2, blank=True, merged=True,
                                  duplicate=True, title="merged_b")],
        "missing_header.xlsx": [dict(fmt="B", count=5, header_row=20, col=2,
                                     missing_range=True, title="missing_range")],
        "formula_blank.xlsx": [dict(fmt="A", count=6, header_row=24, col=2,
                                    formula=True, title="formula_blank")],
        "labels_spaced.xlsx": [dict(fmt="A", count=6, header_row=24, col=2,
                                    spaced=True, title="spaced_a"),
                               dict(fmt="B", count=6, header_row=28, col=3,
                                    spaced=True, title="spaced_b")],
        "korean_mixed_header.xlsx": [dict(fmt="B", count=4, header_row=25, col=2,
                                           korean_range=True, title="korean_range")],
        "large_120.xlsx": [dict(fmt="A" if i % 2 == 0 else "B", count=i + 1,
                                header_row=18 + i % 7, col=2 + i % 4,
                                merged=i % 8 == 0,
                                title=f"bulk_{i:03d}") for i in range(120)],
        "large_200.xlsx": [dict(fmt="A", count=200, header_row=31, col=4,
                                blank=True, title="large_a"),
                           dict(fmt="B", count=200, header_row=44, col=5,
                                blank=True, title="large_b")],
    }
    paths, truth = [], []
    for name, specs in cases.items():
        path, expected = _workbook(root, name, specs)
        paths.append(path)
        truth.extend(expected)
    ground_truth = root / "ground_truth.json"
    ground_truth.write_text(json.dumps(truth, ensure_ascii=False, indent=2), encoding="utf-8")
    return source, paths, ground_truth


def assess(expected, reports):
    """Compare independent generator coordinates with observed engine results."""
    actual = [v for report in reports for sheet in report.get("sheets", [])
              for v in sheet["variables"] if v["extracted"]]
    observed = [v for report in reports for sheet in report.get("sheets", [])
                for v in sheet["variables"]]
    actual_keys = Counter((Path(v["file"]).name, v["sheet"], v["name"], v["name_cell"], v["range_cell"])
                          for v in actual)
    observed_by_name = defaultdict(list)
    for v in observed:
        observed_by_name[(Path(v["file"]).name, v["sheet"], v["name"], v["name_cell"])].append(v)
    missing = []
    for item in expected:
        key = (item["file"], item["sheet"], item["name"], item["name_cell"], item["range_cell"])
        if actual_keys[key] > 0:
            actual_keys[key] -= 1
            continue
        candidates = observed_by_name.get(key[:4], [])
        reason = (candidates[0]["reason"] or "NOT_EXTRACTED") if candidates else "NOT_OBSERVED"
        wrong = next((v for v in candidates if v["range_cell"] != item["range_cell"]), None)
        missing.append(dict(**item, reason=reason,
                            observed_range_cell=wrong["range_cell"] if wrong else None))
    duplicates_or_extra = [dict(file=k[0], sheet=k[1], name=k[2], name_cell=k[3],
                                range_cell=k[4], count=n)
                           for k, n in actual_keys.items() if n]
    formats = {(Path(r["file"]).name, s["sheet"]): s["format"]
               for r in reports for s in r.get("sheets", [])}
    format_errors = list({(v["file"], v["sheet"], v["format"],
                           formats.get((v["file"], v["sheet"]))) for v in expected
                          if formats.get((v["file"], v["sheet"])) != v["format"]})
    format_errors = [dict(file=f, sheet=s, expected=e, observed=o)
                     for f, s, e, o in sorted(format_errors)]
    return dict(ground_truth_variables=len(expected), actual_extracted_variables=len(actual),
                correct_extractions=len(expected)-len(missing), missing=len(missing),
                duplicate_or_extra=sum(v["count"] for v in duplicates_or_extra),
                wrong_cell_links=sum(v["observed_range_cell"] is not None for v in missing),
                format_errors=len(format_errors), missing_rows=missing,
                duplicate_or_extra_rows=duplicates_or_extra, format_error_rows=format_errors)
