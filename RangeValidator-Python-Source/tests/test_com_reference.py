"""Sanitized reference-flow, failure, limits, and V3 integration tests only."""
import importlib.util
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import patch
from rangevalidator.com_values import _last_cell, _copy_sheet, excel_session, close_excel
from rangevalidator.readers import read_com
from rangevalidator.models import ReadFailure, ReaderLimits
from rangevalidator.diagnostics import Diagnostics, format_event
from test_readers import FakeRuntime, FakeCell, ComError, snapshot


class Cells:
    def __init__(self, ws): self.ws=ws; self.finds=[]
    def Item(self, r, c):
        ws=self.ws
        cell=FakeCell(r,c,ws.data.get((r,c)), ws.formulas.get((r,c)), (r,c) in ws.merges)
        cell.Value=cell.Value2
        def end(direction):
            candidates=[(rr,cc) for rr,cc in ws.data if (cc==c if direction==-4162 else rr==r)]
            rr,cc=max(candidates,key=lambda x:x[0] if direction==-4162 else x[1]) if candidates else (1,c) if direction==-4162 else (r,1)
            return self.Item(rr,cc)
        cell.End=end
        return cell
    __call__=Item
    def Find(self, **kw):
        self.finds.append(kw)
        if kw['SearchDirection']==1: raise AssertionError('first_column must not be mandatory')
        if kw['SearchOrder'] in self.ws.fail_find: raise ComError('synthetic Find denied')
        if not self.ws.data: return None
        r,c=max(self.ws.data,key=lambda x:x[0] if kw['SearchOrder']==1 else x[1])
        return self.Item(r,c)


class Sheet:
    Name='Sanitized'; Visible=-1
    Rows=NS(Count=1048576); Columns=NS(Count=16384)
    def __init__(self,data=None,fail_find=(), formulas=None,merges=()):
        self.data={(1,1):'Name',(1,2):'Type',(1,3):'Range',(3,3):2} if data is None else data
        self.formulas=formulas or {};self.merges=set(merges); self.fail_find=fail_find
        self.Cells=Cells(self);self.range_calls=0
        self.UsedRange=NS(Row=1,Column=1,Rows=NS(Count=max((r for r,c in self.data),default=1)),Columns=NS(Count=max((c for r,c in self.data),default=1)))
    def Range(self, first,last):
        self.range_calls+=1
        values=tuple(tuple(self.data.get((r,c)) for c in range(first.Column,last.Column+1)) for r in range(first.Row,last.Row+1))
        formulas=tuple(tuple(self.formulas.get((r,c),self.data.get((r,c))) for c in range(first.Column,last.Column+1)) for r in range(first.Row,last.Row+1))
        return NS(Value2=values,Formula=formulas,MergeCells=None if self.merges else False)


def runtime(sheet=None, fail=None):
    rt=FakeRuntime(fail);rt.book.Worksheets.ws=sheet or Sheet();return rt


class BoundsTests(unittest.TestCase):
    """V3.1 경계 규칙: Find('?*') 가 1차, 실패한 쪽은 UsedRange 를 상한으로만, 블록 읽기가 실제 끝을 정한다."""
    def test_find_both_ok(self):
        ws=Sheet();self.assertEqual(_last_cell(ws),(3,3))
        self.assertTrue(all(f['What']=='?*' and f['SearchDirection']==2 for f in ws.Cells.finds))
    def test_column_find_failure_uses_usedrange_cap(self):
        ws=Sheet({(3,300):'x'},fail_find=(2,))
        self.assertEqual(_last_cell(ws),(3,256))            # 열은 256 으로 상한
    def test_row_find_failure_uses_usedrange_cap(self):
        ws=Sheet({(400,3):'x'},fail_find=(1,))
        self.assertEqual(_last_cell(ws),(4096,3))
    def test_both_fail_empty(self):
        for ws in (Sheet({}),Sheet({},fail_find=(1,2))):
            self.assertEqual(_last_cell(ws),(4096,64))             # Find 실패 → 최소 탐색 범위 (블록 읽기가 빈 시트임을 확인)
            sheet,*_=_copy_sheet(ws,ReaderLimits(),10**9);self.assertIsNone(sheet.bounds)   # 읽어 보면 빈 시트
    def test_phantom_empty_string_cells_do_not_extend_bounds(self):
        ws=Sheet({(1,1):'a',(2,1):'b',(150000,1):''})            # DRM 유령 칸("") 이 멀리 있어도 실제 끝은 2행
        sheet,cells,text,errors=_copy_sheet(ws,ReaderLimits(),10**9)
        self.assertEqual(sheet.bounds,(1,1,2,1));self.assertLess(ws.range_calls,8)
    def test_values_formulas_merges(self):
        ws=Sheet(formulas={(3,3):'=1+1'},merges={(2,1),(2,2)})
        sheet,*_=_copy_sheet(ws,ReaderLimits(),10**9)
        self.assertEqual((sheet.cells['C3'].value,sheet.cells['C3'].formula,sheet.cells['B2'].merge_anchor),(2,'=1+1','A2'))


class ReaderTests(unittest.TestCase):
    def tearDown(self):close_excel()
    def test_bulk_values_formula_and_hidden(self):
        ws=Sheet(formulas={(3,3):'=1+1'});ws.Visible=2
        b=read_com('synthetic.xlsx',_runtime=runtime(ws));s=b.sheets[0]
        self.assertEqual((s.state,s.cells['C3'].value,s.cells['C3'].formula),('veryHidden',2,'=1+1'))
        self.assertLessEqual(ws.range_calls,3)                  # 블록 1회 + CountA + 병합 질의
    def test_merges_restored(self):
        ws=Sheet(merges={(2,1),(2,2)})
        b=read_com('synthetic.xlsx',_runtime=runtime(ws))
        self.assertEqual(b.sheets[0].cells['B2'].merge_anchor,'A2')
    def test_large_bulk_read(self):
        ws=Sheet({(20000,3):'last'});b=read_com('synthetic.xlsx',_runtime=runtime(ws))
        self.assertEqual(b.sheets[0].cells['C20000'].value,'last');self.assertEqual(b.sheets[0].bounds,(1,1,20000,3))
    def test_session_reuses_and_closes(self):
        rt=runtime()
        with excel_session():
            for i in range(3):read_com('synthetic.xlsx',_runtime=rt)
            self.assertNotIn(('quit',),rt.events)
        self.assertEqual(rt.events.count(('dispatch',)),1)
        self.assertEqual(rt.events.count(('quit',)),1)
        self.assertEqual(sum(e[0]=='close' for e in rt.events),3)
    def test_open_none_exact_fullpath(self):
        rt=runtime();rt.book.FullName=str(Path('synthetic.xlsx').resolve());rt.book.Name='synthetic.xlsx'
        rt.app.Workbooks.Open=lambda **kw:None
        rt.app.Workbooks.Count=1;rt.app.Workbooks.Item=lambda i:rt.book
        self.assertTrue(read_com('synthetic.xlsx',_runtime=rt).sheets)
    def test_open_none_same_basename_other_path_rejected(self):
        rt=runtime();rt.book.FullName=str(Path('other/synthetic.xlsx').resolve())
        rt.app.Workbooks.Open=lambda **kw:None;rt.app.Workbooks.Count=1;rt.app.Workbooks.Item=lambda i:rt.book
        with self.assertRaises(ReadFailure):read_com('synthetic.xlsx',_runtime=rt)
    def test_security_open_failure_cleanup(self):
        for fail in ('sta','dispatch','open'):
            rt=runtime(fail=fail)
            with self.subTest(fail=fail),self.assertRaises(ReadFailure):read_com('synthetic.xlsx',_runtime=rt)
            if fail in ('security','open'):self.assertIn(('quit',),rt.events)
    def test_cleanup_failures_retained(self):
        rt=runtime(fail='cleanup')
        with self.assertRaises(ReadFailure) as ctx:read_com('synthetic.xlsx',_runtime=rt)
        self.assertIn('com.Quit',[e['stage'] for e in ctx.exception.cleanup_errors]);self.assertIn(('uninit',),rt.events)
    def test_find_hresult_sheet_name_and_fallback_bounds(self):
        d=Diagnostics();d.set_context(file_index=7,input_role='detail')
        book = read_com('private-company.xlsx',_runtime=runtime(Sheet(fail_find=(1,2))),_trace=d)
        log=''.join(format_event(e) for e in d.events)
        self.assertNotIn('private-company',log);self.assertIn('Sheet: Sanitized',log)
        self.assertTrue(any(e.hresult=='0x800A03EC' for e in d.events))
        self.assertTrue(any(e.stage=='com.bounds.final' and e.row==4096 and e.column==64 and e.file_index==7 for e in d.events))
        self.assertEqual(book.sheets[0].bounds, (1, 1, 3, 3))


@unittest.skipUnless(os.environ.get('RV_TEST_EXCEL')=='1','Actual Excel opt-in required')
class ActualExcelTests(unittest.TestCase):
    def test_common_ab_multifile_engine_and_preserve(self):
        from v3_fixtures import fixtures
        from rangevalidator.workflow import run_comparison
        with tempfile.TemporaryDirectory() as folder:
            common,headers,details=fixtures(Path(folder))
            before={p:snapshot(p) for p in [common,*headers,*details]}
            native=run_comparison(common,headers,details,'openpyxl')
            diag=Diagnostics();com=run_comparison(common,headers,details,'com',diagnostics=diag)
            self.assertTrue(all(f.read_status=='SUCCESS' and f.compared for f in com.files))
            key=lambda r:(r.sheet,r.cell,r.struct,r.member,r.width,r.expected,r.status,r.candidate)
            self.assertEqual([key(r) for r in native.results],[key(r) for r in com.results])
            self.assertEqual(before,{p:snapshot(p) for p in before})
            self.assertEqual(sum(e.stage=='com.DispatchEx' and e.status=='ENTER' for e in diag.events),1)
            self.assertTrue(any(e.stage=='com.process.exit' and e.status=='OK' for e in diag.events))

    def test_actual_large_five_column_sheet(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'large.xlsx';w=Workbook();ws=w.active
            ws.append(['Input/Output','Property','Name','Data type','Scope'])
            ws.cell(15000,5,'tail');w.save(path);w.close()
            before=snapshot(path);book=read_com(path)
            self.assertEqual(book.sheets[0].bounds,(1,1,15000,5))
            self.assertEqual(book.sheets[0].cells['E15000'].value,'tail')
            self.assertEqual(before,snapshot(path))
