from copy import deepcopy
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook
from rangevalidator.engine import Result
from rangevalidator.manual_edits import ReviewRequest
from rangevalidator.report import create_report, MISMATCH_HEADERS, REVIEW_HEADERS
from rangevalidator.workflow import Session, FileStatus


class ReportTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name).resolve()
        self.paths = [str(self.root / folder / 'Design.xlsx') for folder in ('first', 'second')]
        a = Result('m1', self.paths[0], '1.Main', 'H24', 'device.mode', 'MISMATCH', 'range',
                   existing='0~3', expected='0~7', candidate=True)
        self.rows = [
            a, replace(a, id='m2', cell='H25'),
            replace(a, id='create', cell='H26', status='CREATE', existing=None),
            replace(a, id='manual', cell='H27', status='REVIEW', candidate=False,
                    expected='', reason='타입 불명'),
            replace(a, id='review', cell='H28', status='REVIEW', candidate=False, expected='', reason=''),
            replace(a, id='match', cell='H29', status='MATCH', candidate=False),
            replace(a, id='skip', cell='H30'),
            replace(a, id='fail', file=self.paths[1], sheet='2.Other'),
            replace(a, id='wrong', file=self.paths[1], sheet='2.Other', cell='H25'),
            replace(a, id='cover', sheet='표지'),
        ]
        self.session = Session((), tuple(self.paths), results=self.rows, files=[
            FileStatus(1,self.paths[0],status='COMPLETED',read_status='SUCCESS'),
            FileStatus(2,self.paths[1],status='PARTIAL',read_status='SUCCESS',partial=True)],
            coverage=[(self.paths[1],'3.Broken','TARGET','')],
            errors=[dict(file=self.paths[1],sheet_name='3.Broken',error_kind='SHEET_READ_ERROR',
                         stage='com.extract',reason='읽기 실패')])
        self.outcomes = {}
        for rid, state, kind, file_state in [
            ('m1','SAVED_VERIFIED','AUTO_MISMATCH','SAVED_UNVERIFIED'),
            ('create','SAVED_UNVERIFIED','AUTO_CREATE','SAVED_UNVERIFIED'),
            ('manual','SAVED_VERIFIED','MANUAL_REVIEW','SAVED_UNVERIFIED'),
            ('skip','SKIPPED','AUTO_MISMATCH','SAVED_UNVERIFIED'),
            ('fail','SAVE_FAILED','AUTO_MISMATCH','SAVE_FAILED'),
            ('wrong','SAVED_VERIFY_FAILED','AUTO_MISMATCH','SAVE_FAILED'),
        ]:
            row = next(r for r in self.rows if r.id == rid)
            self.outcomes[rid] = dict(file=row.file,status=state,edit_kind=kind,file_status=file_state,
                output=row.file.replace('.xlsx','_fixed.xlsx') if state.startswith('SAVED') else '',
                manual_value='-1~7' if rid=='manual' else '',reason='값 변경 제외' if state=='SKIPPED' else state,
                cleanup_warnings=[dict(phase='verification',api='com.Quit',code=5)] if rid=='manual' else [])

    def test_summary_counts_paths_and_read_errors(self):
        before = deepcopy(self.outcomes)
        report = create_report(self.session,self.outcomes,self.root / 'report.xlsx')
        workbook = load_workbook(report)
        self.addCleanup(workbook.close)
        rows = list(workbook['Summary'].values)
        metrics = {row[0]: row[1] for row in rows[:19]}
        self.assertEqual(metrics['분석 대상 Excel 파일 수'], 2)
        self.assertEqual(metrics['MATCH'], 1)
        self.assertEqual(metrics['MISMATCH'], 5)
        self.assertEqual(metrics['CREATE'], 1)
        self.assertEqual(metrics['REVIEW'], 2)
        self.assertEqual(metrics['자동 MISMATCH 수정 선택 건수'], 4)
        self.assertEqual(metrics['CREATE 수정 선택 건수'], 1)
        self.assertEqual(metrics['REVIEW 수동 수정 선택 건수'], 1)
        self.assertEqual(metrics['저장 및 검증 완료 건수'], 2)
        self.assertEqual(metrics['저장됨 / 검증 보류 건수'], 1)
        files = [row for row in rows if row[0] == 'Design.xlsx']
        self.assertEqual({row[1] for row in files}, set(self.paths))
        broken = next(row for row in files if row[2] == '3.Broken')
        sheet_header = next(row for row in rows if row[2] == '함수 시트명')
        self.assertEqual(broken[sheet_header.index('분석 상태')], 'ERROR')
        self.assertEqual(self.outcomes,before)

    def test_histories_filters_manual_value_and_original_reason(self):
        workbook = load_workbook(create_report(self.session,self.outcomes,self.root / 'report.xlsx'))
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.sheetnames, ['Summary','Mismatch','Review','Errors','Results'])
        mismatch = workbook['Mismatch']
        review = workbook['Review']
        self.assertEqual(tuple(mismatch.values)[0],tuple(MISMATCH_HEADERS))
        self.assertEqual(tuple(review.values)[0],tuple(REVIEW_HEADERS))
        records = list(mismatch.values)[1:]
        self.assertEqual(len(records), 7)
        self.assertFalse(any(row[11] == 'MATCH' for row in records))
        manual = next(row for row in records if row[12] == 'MANUAL_REVIEW')
        self.assertEqual(manual[8:16], ('0~3',None,'-1~7','REVIEW','MANUAL_REVIEW','수동 수정','선택','SAVED_VERIFIED'))
        unselected = next(row for row in records if row[3] == 'H25' and row[1] == self.paths[0])
        self.assertEqual(unselected[14:16], ('미선택','NOT_SELECTED'))
        skipped = next(row for row in records if row[15]=='SKIPPED')
        self.assertEqual(skipped[17],'값 변경 제외')
        reviews = list(review.values)[1:]
        self.assertEqual(len(reviews), 2)
        self.assertEqual(reviews[0][8:12], ('타입 불명','선택','-1~7','SAVED_VERIFIED'))
        self.assertIsNone(reviews[1][8])
        for ws in (mismatch, review):
            self.assertEqual(ws.freeze_panes, 'A2')
            self.assertEqual(ws.auto_filter.ref, ws.dimensions)
        self.assertTrue(any(row[3]=='CLEANUP_WARNING' for row in list(workbook['Errors'].values)[1:]))

    def test_empty_report_and_unselected_draft_are_safe(self):
        session = Session((), ())
        workbook = load_workbook(create_report(session,{},self.root / 'empty.xlsx'))
        self.addCleanup(workbook.close)
        for name in ('Mismatch','Review'):
            self.assertEqual(workbook[name].max_row, 1)
            self.assertEqual(workbook[name].auto_filter.ref, workbook[name].dimensions)
        workbook2 = load_workbook(create_report(self.session,{},self.root / 'draft.xlsx',
            manual_requests=[ReviewRequest('review','0~9')],selected_ids=['m2']))
        self.addCleanup(workbook2.close)
        row = next(row for row in list(workbook2['Review'].values)[1:] if row[3]=='H28')
        self.assertEqual(row[9:12], ('미선택','0~9','NOT_SELECTED'))
