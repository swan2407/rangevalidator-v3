import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from openpyxl import Workbook
from rangevalidator.readers import read_openpyxl
from rangevalidator.a_extraction import scan_a, normalize_header
from rangevalidator.engine import catalog_headers, compare_book


def fixture(path, count=10, row=20, name_col=4, variant=None):
    wb=Workbook();ws=wb.active;ws.title='1.Main'
    ws.cell(row-2,name_col-2,'함수 변수 정보')
    columns=[name_col-2,name_col-1,name_col,name_col+1,name_col+3]
    labels=['Input/Output','Property','Name','Data Type','Scope']
    if variant=='spaced': labels=[' Input/\nOutput ','PROP ERTY','Na\tme','Data\rType','Sc\u3000ope']
    if variant=='alias': labels[-1]='범위'
    if variant=='missing_scope':labels[-1]=''
    if variant=='missing_name':labels[2]=''
    if variant=='description_name':labels[2]='Function Name'
    for col,label in zip(columns,labels):ws.cell(row,col,label)
    if variant=='duplicate_scope':ws.cell(row,name_col+5,'Scope')
    if variant=='duplicate_name':ws.cell(row,name_col+5,'NAME')
    if variant=='merged_header':ws.merge_cells(start_row=row,start_column=name_col+3,end_row=row,end_column=name_col+4)
    truth=[]
    for i in range(count):
        r=row+1+i+(3 if i>=count//2 else 0)
        name='obj.v001' if variant!='variable_spaces' else 'obj . v001'
        for col,value in zip(columns,['Input','Global',name,'unsigned int',None if i%2 else '0~3']):
            ws.cell(r,col,value)
        truth.append((ws.cell(r,name_col).coordinate,ws.cell(r,name_col+3).coordinate))
    if variant=='formula':ws.cell(row+1,name_col+3,'=IF(1=0,1,"")')
    if variant=='nonanchor':ws.merge_cells(start_row=row+1,start_column=name_col+2,end_row=row+1,end_column=name_col+3)
    ws.cell(row+count+8,name_col-2,'Function description')
    ws.cell(row+count+9,name_col,'obj.v001');ws.cell(row+count+9,name_col+3,'0~3')
    wb.save(path);wb.close();return truth


class AExtractionTests(unittest.TestCase):
    def setUp(self):
        self.temp=TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name);self.path=self.root/'sample.xlsx'
        source=self.root/'test.h';source.write_text('typedef struct {unsigned int v001:3;} T; extern T obj;')
        self.catalog=catalog_headers([source])
    def read(self,**kwargs):
        truth=fixture(self.path,**kwargs);book=read_openpyxl(self.path)
        headers,tables,candidates=scan_a(book.sheets[0])
        results,_=compare_book(None,self.catalog,book)
        return truth,tables,candidates,results
    def test_normalization_is_header_only(self):
        for value in ['Name','NAME',' name ','Na me','Na\nme','Na\r\tme','Na\u3000me','Na\u00a0me']:
            self.assertEqual(normalize_header(value),'name')
        self.assertNotEqual(normalize_header('Function Name'),'name')
        self.assertNotEqual(normalize_header('Variable Name'),'name')
        self.assertNotEqual(normalize_header('File Name'),'name')
        _,_,c,_=self.read(variant='variable_spaces')
        self.assertEqual(c[0]['name'],'obj . v001')
    def test_positions_sizes_and_blank_rows(self):
        for size in (1,10,50,200,500):
            for row,col in ((5,4),(20,6),(100,8)):
                with self.subTest(size=size,row=row,col=col):
                    truth,_,c,r=self.read(count=size,row=row,name_col=col,variant='spaced')
                    self.assertEqual([(v['name_cell'],v['range_cell']) for v in c],truth)
                    self.assertEqual([v.cell for v in r],[v[1] for v in truth])
    def test_alias_and_merged_header(self):
        for variant in ('alias','merged_header'):
            truth,_,c,r=self.read(variant=variant)
            self.assertEqual(len(r),10)
            self.assertEqual([(v['name_cell'],v['range_cell']) for v in c],truth)
    def test_incomplete_and_duplicate_headers_never_guess_scope(self):
        for variant in ('missing_scope','duplicate_scope','duplicate_name','missing_name','description_name'):
            with self.subTest(variant=variant):
                _,tables,c,r=self.read(variant=variant)
                if variant in ('missing_scope', 'duplicate_scope'):
                    self.assertEqual(len(r), 10)
                    self.assertTrue(all(v.status == 'REVIEW' and not v.candidate and not v.cell for v in r))
                else:
                    self.assertEqual(r,[])
                self.assertTrue(tables[0]['state']!='HEADER_FOUND')
                self.assertTrue(all(v['range_cell'] is None for v in c))
    def test_formula_and_nonanchor_keep_review(self):
        for variant in ('formula','nonanchor'):
            _,_,_,r=self.read(variant=variant)
            self.assertEqual(r[0].status,'REVIEW');self.assertFalse(r[0].candidate)
    def test_multiple_tables_keep_their_own_columns(self):
        fixture(self.path,count=1,row=5)
        from openpyxl import load_workbook
        w=load_workbook(self.path);s=w.active
        for col,label in zip((4,5,6,7,9),('Input/Output','Property','Name','Data Type','Scope')):s.cell(30,col,label)
        s['F31']='obj.v001';s['I31']='0~7';w.save(self.path);w.close()
        b=read_openpyxl(self.path);_,t,c=scan_a(b.sheets[0])
        self.assertEqual(len(t),2)
        self.assertEqual([(v['name_cell'],v['range_cell']) for v in c],[('D9','G9'),('F31','I31')])

