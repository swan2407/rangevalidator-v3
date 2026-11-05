"""Explicit related-test allowlist. Full/Writer/report suites are not executed."""
import io,json,os,sys,time,unittest,hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT));sys.path.insert(0,str(ROOT/'tests'))
os.environ['RV_TEST_EXCEL']='1';os.environ['RV_TEST_GUI']='1'
names=['test_drm_batch.InputStateTests','test_drm_batch.BatchWorkflowTests','test_drm_batch.ComReadTests','test_drm_batch.BatchGuiTests',
       'test_readers.ComContractTests','test_readers.ComCleanupExtraTests','test_readers.ExcelIntegrationTests',
       'test_readers.ExcelLifecycleTests.test_missing_corrupt_paths']
suite=unittest.TestSuite(unittest.defaultTestLoader.loadTestsFromName(n) for n in names)
stream=io.StringIO();start=time.perf_counter();r=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
(root:=ROOT/'docs').joinpath('DRM_BATCH_TEST_RUN.log').write_text(stream.getvalue(),encoding='utf-8')
record=dict(tests=r.testsRun,failures=len(r.failures),errors=len(r.errors),skipped=len(r.skipped),seconds=round(time.perf_counter()-start,3),successful=r.wasSuccessful(),allowlist=names,full_suite_run=False,writer_report_tests_run=False,
 source_sha256={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'rangevalidator').glob('*.py')})
(root/'DRM_BATCH_TEST_RUN.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(record,ensure_ascii=False,indent=2))
if not r.wasSuccessful():print(stream.getvalue())
raise SystemExit(int(not r.wasSuccessful() or bool(r.skipped)))
