"""Excel-1 evidence checks; no production extraction or decision changes."""
import json
import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from rangevalidator.extraction_diagnostics import diagnose_session
from rangevalidator.workflow import run_comparison
from tools.excel1_synthetic import assess, build_suite


class ExtractionDiagnosticsTests(unittest.TestCase):
    def test_synthetic_ground_truth_and_unchanged_results(self):
        with TemporaryDirectory() as directory:
            source, paths, truth_path = build_suite(directory)
            expected = json.loads(truth_path.read_text(encoding="utf-8"))
            reports = []
            for path in paths:
                session = run_comparison(None, [source], [path], backend="openpyxl")
                original_results = tuple(session.results)
                signature = session.signature()
                report = diagnose_session(session)[0]
                self.assertEqual(tuple(session.results), original_results)
                self.assertEqual(session.signature(), signature)
                reports.append(report)
            outcome = assess(expected, reports)
            self.assertEqual(outcome["ground_truth_variables"], 7793)
            self.assertEqual(outcome["correct_extractions"], 7788)
            self.assertEqual(outcome["wrong_cell_links"], 0)
            self.assertEqual(outcome["duplicate_or_extra"], 0)
            self.assertEqual(outcome["missing"], 5)
            self.assertEqual({item["file"] for item in outcome["missing_rows"]},
                             {"missing_header.xlsx"})
            bulk = next(report for report in reports if Path(report["file"]).name == "large_120.xlsx")
            self.assertEqual(bulk["counts"]["function_candidates"], 120)
            self.assertEqual(sum(v["extracted"] for sheet in bulk["sheets"]
                                 for v in sheet["variables"]), 7260)
            self.assertEqual(bulk["counts"]["excluded_sheets"], 2)
            self.assertGreaterEqual(bulk["timings"]["body_read_seconds"], 0)

    def test_formula_blank_and_excluded_body_are_observed(self):
        with TemporaryDirectory() as directory:
            source, paths, _ = build_suite(directory)
            path = next(p for p in paths if p.name == "formula_blank.xlsx")
            session = run_comparison(None, [source], [path], backend="openpyxl")
            report = diagnose_session(session)[0]
            function = next(s for s in report["sheets"] if s["function_sheet"])
            first = function["variables"][0]
            self.assertTrue(first["extracted"])
            self.assertEqual(first["comparison_status"], "REVIEW")
            self.assertEqual(first["range_formula"], '=IF(1=0,"x","")')
            excluded = next(s for s in report["sheets"] if s["sheet"] == "commonparameter")
            self.assertIn("SHEET_EXCLUDED", excluded["states"])
            self.assertEqual(report["counts"]["body_read_success"], report["total_sheets"])

    def test_source_binding_review_is_distinct_from_excel_omission(self):
        with TemporaryDirectory() as directory:
            source, paths, _ = build_suite(directory)
            source.write_text(source.read_text(encoding="utf-8").replace("extern Excel1_t obj;", ""),
                              encoding="utf-8")
            path = next(p for p in paths if p.name == "a_10.xlsx")
            session = run_comparison(None, [source], [path], backend="openpyxl")
            report = diagnose_session(session)[0]
            variables = next(s for s in report["sheets"] if s["function_sheet"])["variables"]
            self.assertEqual(len(variables), 10)
            self.assertTrue(all(v["extracted"] and v["comparison_status"] == "REVIEW"
                                and v["comparison_reason"] for v in variables))

    def test_assessor_detects_wrong_range_connection(self):
        with TemporaryDirectory() as directory:
            source, paths, ground_truth = build_suite(directory)
            path = next(p for p in paths if p.name == "a_10.xlsx")
            session = run_comparison(None, [source], [path], backend="openpyxl")
            report = diagnose_session(session)[0]
            broken = copy.deepcopy(report)
            variable = next(s for s in broken["sheets"] if s["function_sheet"])["variables"][0]
            variable["range_cell"] = "X99"
            truth = [v for v in json.loads(ground_truth.read_text(encoding="utf-8"))
                     if v["file"] == path.name]
            result = assess(truth, [broken])
            self.assertEqual(result["wrong_cell_links"], 1)


if __name__ == "__main__":
    unittest.main()
