from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

from openpyxl import load_workbook
from rangevalidator.engine import Result
from rangevalidator.gui import gui_save_text
from rangevalidator.report import create_report
from rangevalidator.save_display import CLEANUP_LABEL, SAVE_LABELS
from rangevalidator.writer import SaveOutcomes


class SaveDisplayTests(unittest.TestCase):
    def setUp(self):
        self.outcomes = SaveOutcomes()
        self.rows = []
        files = [('verified.xlsx', ['SAVED_VERIFIED'] * 8 + ['SKIPPED'] * 2),
                 ('pending.xlsx', ['SAVED_VERIFIED', 'SAVED_UNVERIFIED']),
                 ('wrong.xlsx', ['SAVED_VERIFY_FAILED']),
                 ('failed.xlsx', ['SAVE_FAILED'])]
        for file, states in files:
            file_state = next(state for state in reversed(states) if state != 'SKIPPED')
            for index, state in enumerate(states, 1):
                row = Result(f'{file}:{index}', file, '1.Main', f'C{index}', 'obj.bit',
                             'MISMATCH', '', candidate=True)
                self.rows.append(row)
                self.outcomes[row.id] = dict(file=file, file_status=file_state, status=state,
                                             output=file.replace('.xlsx', '_fixed.xlsx') if state.startswith('SAVED_') else '',
                                             reason='수식 셀' if state == 'SKIPPED' else state,
                                             cleanup_warnings=[dict(phase='verification', api='com.Quit', code=5)]
                                             if file == 'verified.xlsx' else [])
        self.rows.append(Result('review', 'verified.xlsx', '1.Main', 'C11', 'unknown.bit',
                                'REVIEW', '타입 불명'))

    def test_gui_uses_file_status_and_cell_counts(self):
        self.outcomes.report_warning = '보고서 생성 실패'
        summary, details = gui_save_text('output', self.outcomes)
        for state, label in SAVE_LABELS.items():
            if state != 'SKIPPED':
                self.assertIn(f'{label} 1개', summary)
        for value in ('선택 후보 10칸 / 저장 확인 8칸', '수정 제외 2칸',
                      '저장 및 검증 완료 8칸', 'pending_fixed.xlsx', '보고서 생성 실패'):
            self.assertIn(value, details)
        self.assertEqual(details.count(CLEANUP_LABEL), 1)

    def test_report_retains_raw_status_labels_counts_and_warnings(self):
        session = SimpleNamespace(id='test', results=self.rows, errors=[])
        with TemporaryDirectory() as directory:
            path = create_report(session, self.outcomes, Path(directory) / 'report.xlsx')
            workbook = load_workbook(path)
            try:
                rows = list(workbook['Summary'].values)
                file_header = next(row for row in rows if row[0] == '원본 파일명')
                verified = next(row for row in rows if row[0] == 'verified.xlsx')
                pending = next(row for row in rows if row[0] == 'pending.xlsx')
                self.assertEqual(verified[file_header.index('MISMATCH')], 10)
                self.assertEqual(verified[file_header.index(SAVE_LABELS['SAVED_VERIFIED'])], 8)
                self.assertEqual(verified[file_header.index(SAVE_LABELS['SKIPPED'])], 2)
                self.assertEqual(pending[file_header.index('파일 저장 상태')], 'SAVED_UNVERIFIED')
                self.assertIn('com.Quit', verified[file_header.index(CLEANUP_LABEL)])
                header, *cells = workbook['Results'].values
                raw, label = header.index('Write result'), header.index('셀 저장 상태')
                self.assertEqual({row[raw] for row in cells},
                                 {*SAVE_LABELS, 'NOT_SELECTED'})
                self.assertEqual(next(row[label] for row in cells if row[raw] == 'SKIPPED'), '수정 제외')
                self.assertEqual(cells[-1][header.index('Decision')], 'REVIEW')
            finally:
                workbook.close()
