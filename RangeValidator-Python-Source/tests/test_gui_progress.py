from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from threading import Thread
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rangevalidator.gui import App, source_files
from rangevalidator.models import Workbook
from rangevalidator.workflow import run_comparison


class ProgressTests(unittest.TestCase):
    def app(self, kind):
        app = App.__new__(App)
        app.progress_kind = kind
        app.progress_percent = 0
        app.queue = Queue()
        app.controls = []
        app.tabs = [None] * 4
        app.after_id = 'test'
        app.closing = False
        for name in ('progress', 'progress_text', 'progress_stage', 'root', 'encoding_combo',
                     'status', 'save_info', 'notebook', 'diagnostic_result', 'show_diagnostics'):
            setattr(app, name, Mock())
        return app

    def test_analysis_progress_uses_catalog_and_workbook_operations(self):
        with TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            source = root / 'test.h'
            source.touch()
            details = [root / 'good.xlsx', root / 'bad.xlsx']
            events = []

            def read(path, **_):
                if path == str(details[1]):
                    raise OSError('synthetic read failure')
                return Workbook(path, 'com')

            with patch('rangevalidator.workflow.read_com', side_effect=read):
                session = run_comparison(headers=[source], details=details, backend='com',
                                         progress=lambda *args: events.append(args))

        self.assertEqual([(done, total) for done, total, _ in events],
                         [(0, 6), (1, 6), (2, 6), (3, 6), (3, 6)])
        self.assertIn('카탈로그', events[0][2])
        self.assertIn('함수 시트 범위 비교', events[2][2])
        self.assertIn('결과 집계', events[-1][2])
        self.assertEqual(session.files[1].status, 'ERROR')

    def test_worker_progress_only_enqueues_and_display_never_regresses(self):
        app = self.app('save')
        def worker():
            app.save_progress_event(1, 2, 'detail.xlsx')
            app.progress_event(0, 3, '이전 단계')
            app.save_progress_event(2, 2, '완료')

        thread = Thread(target=worker)
        thread.start()
        thread.join()
        app.progress.configure.assert_not_called()
        app.drain()
        self.assertEqual([call.kwargs['value'] for call in app.progress.configure.call_args_list],
                         [33, 33, 66])
        app.progress_stage.set.assert_called_with('현재 작업: 저장 결과 보고서 생성 중')

    def test_only_successful_completion_sets_one_hundred(self):
        for outcome, expected in [('SAVED_VERIFIED', 100), ('SAVED_UNVERIFIED', 50),
                                  ('SAVE_FAILED', 50)]:
            with self.subTest(outcome=outcome):
                app = self.app('save')
                app.save_progress_event(1, 1, '완료')
                app.queue.put(('done', 'save', ('output', {'id': dict(
                    file='detail.xlsx', status=outcome, file_status=outcome, reason='test')})))
                app.drain()
                self.assertEqual(app.progress_percent, expected)
        app = self.app('compare')
        app.progress_event(2, 4, '비교 중')
        app.queue.put(('error', 'compare', RuntimeError('synthetic')))
        app.drain()
        self.assertEqual(app.progress_percent, 50)
        app.progress_stage.set.assert_called_with('현재 작업: 작업 중단 · 진단 로그 확인')

    def test_start_resets_progress_and_report_success_finishes(self):
        app = self.app('save')
        app.busy = False
        app.progress_percent = 100
        with patch('rangevalidator.gui.threading.Thread') as worker:
            app.start(lambda: None, 'report')
        app.progress_text.set.assert_called_with('보고서 진행률: 0%')
        worker.call_args.kwargs['target']()
        app.drain()
        self.assertEqual(app.progress_percent, 100)

    def test_folder_scan_and_registration_are_limited_to_supported_sources(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'nested').mkdir()
            inputs = [root / 'a.C', root / 'nested' / 'b.hpp', root / 'ignore.xlsx']
            for path in inputs:
                path.touch()
            app = SimpleNamespace(headers=[str(inputs[0])], sync_inputs=Mock(), invalidate=Mock())
            found = source_files(root)
            self.assertEqual(found, list(map(str, inputs[:2])))
            App.add_paths(app, 'headers', [*found, *found])
            self.assertEqual(app.headers, found)
            app.sync_inputs.assert_called_once()
