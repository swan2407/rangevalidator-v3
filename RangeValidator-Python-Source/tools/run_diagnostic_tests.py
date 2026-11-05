"""Explicit allowlist. Never discover/run the full regression, Writer or report tests."""
from pathlib import Path
import hashlib,io,json,os,sys,time,unittest
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT));sys.path.insert(0,str(ROOT/'tests'))
os.environ['RV_TEST_EXCEL']='1';os.environ['RV_TEST_GUI']='1'
names=['test_diagnostics.DiagnosticLogTests','test_diagnostics.DiagnosticGuiTests',
 'test_readers.ComContractTests','test_readers.ComCleanupExtraTests',
 'test_readers.ExcelIntegrationTests','test_readers.ExcelLifecycleTests.test_missing_corrupt_paths',
 'test_readers.ExcelLifecycleTests.test_mta_refused','test_gui.GuiTests.test_com_failure_auto_log_and_masking']
suite=unittest.TestSuite(unittest.defaultTestLoader.loadTestsFromName(n) for n in names)
stream=io.StringIO();start=time.perf_counter();result=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
out=ROOT/'docs';(out/'DIAGNOSTIC_TEST_RUN.log').write_text(stream.getvalue(),encoding='utf-8')
record=dict(tests=result.testsRun,failures=len(result.failures),errors=len(result.errors),skipped=len(result.skipped),seconds=round(time.perf_counter()-start,3),successful=result.wasSuccessful(),test_allowlist=names,full_suite_run=False,writer_tests_run=False,report_tests_run=False,
 source_sha256={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'rangevalidator').glob('*.py')})
(out/'DIAGNOSTIC_TEST_RUN.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(record,ensure_ascii=False,indent=2))
if not result.wasSuccessful():print(stream.getvalue())
raise SystemExit(int(not result.wasSuccessful() or bool(result.skipped)))
