import os
from pathlib import Path
import tempfile
import unittest
from dataclasses import replace
from unittest.mock import patch
from openpyxl import load_workbook
from rangevalidator.workflow import run_comparison
from rangevalidator.readers import read_openpyxl,read_com,read_auto
from rangevalidator.common import parse_common,parse_declaration
from rangevalidator.engine import catalog_headers,resolve,source_context
from rangevalidator.writer import save_selected
from rangevalidator.report import create_report
from rangevalidator.models import ReadFailure
from v3_fixtures import fixtures


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name)
        self.common,self.headers,self.details=fixtures(self.root)
    def session(self,details=None,backend='openpyxl'):
        return run_comparison(self.common,self.headers,details or self.details,backend)
    def test_literal_decisions_both_formats(self):
        s=self.session()
        for p in self.details:
            actual=[r.status for r in s.results if r.file==str(p) and r.sheet.startswith('1.Design')]
            self.assertEqual(actual,['MISMATCH','MATCH','MISMATCH','REVIEW','REVIEW','REVIEW',
                                     'MISMATCH','CREATE','MISMATCH','REVIEW','MATCH','MATCH'])
        self.assertEqual(sum(r.candidate for r in s.results),10)
    def test_common_declaration_metadata(self):
        c=parse_common(read_openpyxl(self.common))
        self.assertFalse(c.reviews)
        self.assertEqual(len(c.declarations),4)
        d=c.declarations[0]
        self.assertEqual((d.name,d.type_name,d.pointer,d.type_qualifiers,d.pointer_qualifiers,d.row),('p','Device_t',True,['const'],['const'],3))
        self.assertEqual(d.cells['E3'].value,'retained')
    def test_qualifier_array_negative(self):
        good=[('obj','Device_t'),('p','Device_t volatile const * volatile const'),('tag','struct DeviceTag *'),('arr[10]','Device_t'),('arr','Device_t[10]')]
        for n,t in good:self.assertFalse(parse_declaration('motor.c',n,t).reviews,(n,t))
        for t in ('Device_t **','Device_t * restrict','Device_t [N]','Device_t *[10]','Device_t const const','const struct','Device_t extra'):
            self.assertTrue(parse_declaration('motor.c','x',t).reviews,t)
        self.assertTrue(parse_declaration('motor.c','a[10]','Device_t[10]').reviews)
    def test_unrelated_header_type_is_isolated(self):
        other=self.root/'other.h';other.write_text('typedef struct { unsigned int active:5; } Sensor_t;')
        s=run_comparison(self.common,[*self.headers,other],self.details,'openpyxl')
        self.assertEqual(next(r.status for r in s.results if r.expression=='q->active'),'MISMATCH')
        self.assertEqual(next(r.status for r in s.results if r.expression=='p->active'),'MISMATCH')
    def test_preprocessor_does_not_block(self):
        other=self.root/'conditional.h';other.write_text('#if MODE\nstruct A { unsigned int a:1; };\n#endif')
        s=run_comparison(self.common,[*self.headers,other],self.details,'openpyxl')
        self.assertTrue(any(r.candidate for r in s.results))
    def test_sections_exclude_macro_struct_calibration(self):
        w=load_workbook(self.common);ws=w.active
        for title in ('Macro','Struct','Calibration','Enum'):
            ws.append([title]);ws.append(['Name','Type']);ws.append(['bad','Bogus *'])
        w.save(self.common);w.close()
        c=parse_common(read_openpyxl(self.common));self.assertFalse(c.reviews);self.assertEqual(len(c.declarations),4)
    def test_filename_rules(self):
        for name in ('motor.xlsx','motor.c.xlsx','motor_Detail.xlsx','motor_DetailDesign.xlsx'):
            self.assertEqual(source_context(name),'motor.c')
        self.assertEqual(source_context('sub/motor_Detail.xlsx'),'sub/motor.c')
    def test_explicit_context_is_advisory(self):
        w=load_workbook(self.details[0]);w.active['A1']='other.c';w.save(self.details[0]);w.close()
        self.assertTrue(any(r.candidate for r in self.session().results if r.file==str(self.details[0])))
    def test_partial_detail_failure_isolated(self):
        bad=self.root/'bad.xlsx';bad.write_bytes(b'corrupt')
        s=self.session([self.details[0],bad])
        self.assertEqual(len(s.errors),1);self.assertTrue(any(r.candidate for r in s.results if r.file==str(self.details[0])))
    def test_unknown_sheet_is_reported_not_blocking(self):
        w=load_workbook(self.details[0]);w.create_sheet('2.Unknown')['A1']='Unknown layout';w.save(self.details[0]);w.close()
        s=self.session()
        self.assertTrue(any(r.candidate for r in s.results if r.file==str(self.details[0])))
        self.assertTrue(any(c[1]=='2.Unknown' and c[2]=='REVIEW' for c in s.coverage))
    def test_invalid_and_foreign_selection_blocked(self):
        s=self.session();r=next(r for r in s.results if r.candidate)
        for ids in ([next(r.id for r in s.results if r.status=='REVIEW')],[r.id,r.id],['foreign']):
            with self.assertRaises(ValueError):save_selected(s,ids,self.root/'out')
        self.assertFalse((self.root/'out').exists())
    def test_session_tampering_blocked(self):
        s=self.session();s.results[0]=replace(s.results[0],expected='0~999')
        with self.assertRaises(RuntimeError):s.assert_unchanged()
    def test_report_literal_and_all_sheets(self):
        s=self.session();out=self.root/'report.xlsx';create_report(s,{},out)
        w=load_workbook(out);self.assertEqual(w.sheetnames,['Summary','Mismatch','Review','Errors','Results'])
        self.assertEqual(w['Results'].max_row,len(s.results)+1)
        formula = next(cell for cell in w['Results']['J'] if cell.value == '=1+1')
        self.assertEqual(formula.data_type,'s')
        w.close()
        with self.assertRaises(FileExistsError):create_report(s,{},out)
    def test_native_error_fallback_contract(self):
        with patch('rangevalidator.readers.read_openpyxl',side_effect=ReadFailure('openpyxl.open',ValueError('bad format'))),patch('rangevalidator.readers.read_com',return_value='ok') as com:
            self.assertEqual(read_auto('x.xlsx'),'ok');com.assert_called_once()
        for exc in (ReadFailure('path',ValueError('bad')),ReadFailure('openpyxl.open',FileNotFoundError(2,'missing')),ReadFailure('openpyxl.extract',ValueError('budget'))):
            with patch('rangevalidator.readers.read_openpyxl',side_effect=exc),patch('rangevalidator.readers.read_com') as com,self.assertRaises(ReadFailure):read_auto('x.xlsx')
            com.assert_not_called()

    def test_cp949_header_encoding_auto(self):
        original=self.headers[0].read_text(encoding='utf-8')
        self.headers[0].write_bytes((original+'// 한글 합성 주석\n').encode('cp949'))
        for enc in ('utf-8-sig','cp949'):
            s=run_comparison(self.common,self.headers,self.details,'openpyxl',enc)
            self.assertTrue(any(r.candidate for r in s.results),enc)


@unittest.skipUnless(os.environ.get('RV_TEST_EXCEL')=='1','Actual Excel opt-in required')
class WriterIntegrationTests(unittest.TestCase):
    setUp=PipelineTests.setUp
    session=PipelineTests.session
    def test_com_saveas_selected_preservation(self):
        common,headers,details=fixtures(self.root/'images',images=True)
        for detail in details:
            w=load_workbook(detail);w.active.cell(17,3 if w.active.title=='1.Design A' else 5).value='0~99';w.save(detail);w.close()
        s=run_comparison(common,headers,details,'openpyxl')
        before={p:Path(p).read_bytes() for p in [str(common),*map(str,headers),*map(str,details)]}
        selected=[r.id for r in s.results if r.candidate]
        folder,outcomes=save_selected(s,selected,self.root/'out')
        self.assertEqual(len(outcomes),12)
        self.assertTrue(all(o['status']=='SAVED_VERIFIED' for o in outcomes.values()),outcomes)
        self.assertEqual(before,{p:Path(p).read_bytes() for p in before})
        self.assertEqual({p.name for p in folder.iterdir()},{*(d.stem + '_fixed' + d.suffix for d in details), 'RangeValidator_Report.xlsx'})
        saved=run_comparison(common,headers,[folder/(d.stem + '_fixed' + d.suffix) for d in details],'openpyxl')
        positions={(Path(outcomes[r.id]['output']).name,r.sheet,r.cell) for r in s.results if r.id in selected}
        self.assertTrue(all(r.status=='MATCH' for r in saved.results if (Path(r.file).name,r.sheet,r.cell) in positions))
    def test_reader_semantic_decisions_equal(self):
        a=self.session(backend='openpyxl');b=self.session(backend='com')
        self.assertEqual([(r.sheet,r.cell,r.status,r.expected,r.candidate) for r in a.results],[(r.sheet,r.cell,r.status,r.expected,r.candidate) for r in b.results])
    def test_per_file_failure_continues(self):
        import rangevalidator.writer as writer
        s=self.session();real=writer._apply_file
        def apply(source,*args):
            if Path(source)==self.details[1]:raise RuntimeError('synthetic second-file failure')
            return real(source,*args)
        chosen=[r.id for r in s.results if r.candidate and r.cell.endswith('5')]
        with patch.object(writer,'_apply_file',side_effect=apply):folder,outcomes=save_selected(s,chosen,self.root/'out')
        self.assertEqual(sorted(o['status'] for o in outcomes.values()),['SAVED_VERIFIED','SAVE_FAILED'])
        self.assertTrue((folder/'RangeValidator_Report.xlsx').is_file())

    def test_com_only_evidence_never_reports_protection_verified(self):
        import zipfile
        import rangevalidator.writer as writer
        s=self.session()
        selected=[next(r.id for r in s.results if r.candidate)]
        with patch.object(writer,'manifest',side_effect=zipfile.BadZipFile('synthetic native-reader unavailable')) as manifest:
            folder,outcomes=save_selected(s,selected,self.root/'out')
        manifest.assert_not_called()
        self.assertEqual(next(iter(outcomes.values()))['status'],'SAVED_VERIFIED')
        self.assertNotIn('drm', next(iter(outcomes.values())))
