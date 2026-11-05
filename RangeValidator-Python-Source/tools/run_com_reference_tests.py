"""Explicit related tests only. No discovery, full suite, or Writer/report tests."""
import io,json,os,sys,time,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
os.environ['RV_TEST_EXCEL']='1'
names=['test_readers.ComContractTests','test_readers.ComCleanupExtraTests','test_com_reference','test_design.DesignTests','test_drm_batch.InputStateTests','test_drm_batch.BatchWorkflowTests',
 'test_pipeline.PipelineTests.test_common_declaration_metadata','test_pipeline.PipelineTests.test_literal_decisions_both_formats',
 'test_readers.ExcelIntegrationTests','test_readers.ExcelLifecycleTests']
suite=unittest.TestSuite(unittest.defaultTestLoader.loadTestsFromName(n) for n in names)
stream=io.StringIO();start=time.perf_counter()
r=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
(ROOT/'docs/COM_REFERENCE_TEST_RUN.log').write_text(stream.getvalue(),encoding='utf-8')
record=dict(tests=r.testsRun,failures=len(r.failures),errors=len(r.errors),skipped=len(r.skipped),seconds=round(time.perf_counter()-start,3),allowlist=names,full_suite_run=False,actual_company_drm_verified=False)
(ROOT/'docs/COM_REFERENCE_TEST_RUN.json').write_text(json.dumps(record,indent=2),encoding='utf-8')
print(json.dumps(record));print(stream.getvalue())
raise SystemExit(int(not r.wasSuccessful() or bool(r.skipped)))
