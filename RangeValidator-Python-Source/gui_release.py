"""Desktop UI for the DirectCompare release."""

import queue
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from rangevalidator.build_options import BuildOptionsError, load_build_options

from rangevalidator.gui import App as BaseApp
from rangevalidator.manual_edits import MANUAL_REVIEW
from rangevalidator.save_display import gui_save_text
from rangevalidator.workflow import run_comparison


class App(BaseApp):
    def build_inputs(self):
        super().build_inputs()
        self.build_options_path = ''
        box = ttk.LabelFrame(self.tabs[0], text='빌드옵션 (선택)', padding=6)
        box.pack(fill='x', pady=4)
        self.button(box, '빌드옵션 XML 선택', self.choose_build_options).pack(side='left', padx=3)
        self.button(box, '적용 해제', self.clear_build_options).pack(side='left', padx=3)
        self.build_options_text = tk.StringVar(value='미적용 · 기존 방식으로 모든 조건부 후보 비교')
        ttk.Label(box, textvariable=self.build_options_text, wraplength=900).pack(side='left', padx=8)

    def set_build_options(self, path):
        options = load_build_options(path)
        self.build_options_path = options.path
        self.build_options_text.set(f'{options.path} · 매크로 {len(options.macros)}개 · 값 없는 매크로는 정의 여부만 적용')
        self.invalidate()

    def choose_build_options(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(parent=self.root, title='Eclipse CDT 빌드옵션 XML 선택',
                                          filetypes=[('Eclipse CDT XML', '*.xml')])
        if path:
            try:
                self.set_build_options(path)
            except BuildOptionsError as exc:
                messagebox.showerror('빌드옵션 확인', str(exc), parent=self.root)

    def clear_build_options(self):
        if self.busy:
            return
        self.build_options_path = ''
        self.build_options_text.set('미적용 · 기존 방식으로 모든 조건부 후보 비교')
        self.invalidate()

    def __init__(self, root):
        super().__init__(root)
        progress_area = self.progress.master
        for widget in progress_area.master.winfo_children():
            if isinstance(widget, ttk.Button) and widget.cget('text') == '진단 로그 보기':
                widget.destroy()
        self.elapsed_text = tk.StringVar(master=root, value='소요 시간 —')
        ttk.Label(progress_area, textvariable=self.elapsed_text).pack(anchor='w')
        self.started_at = None
        self.finished_at = None

    def build_logs(self):
        # The compiled base app creates the fourth tab before calling this hook.
        tab = self.tabs.pop()
        self.notebook.forget(tab)
        tab.destroy()

    def compare(self):
        if self.busy:
            return
        if not self.headers or not self.details:
            messagebox.showinfo('파일 등록', 'C/H 파일과 상세설계서를 등록하세요.', parent=self.root)
            return
        self.invalidate()
        headers, details = tuple(self.headers), tuple(self.details)
        encoding = self.encoding.get()
        try:
            options = load_build_options(self.build_options_path) if self.build_options_path else None
        except BuildOptionsError as exc:
            messagebox.showerror('빌드옵션 확인', str(exc), parent=self.root)
            return
        self.start(lambda: run_comparison(headers=headers, details=details, encoding=encoding,
                                          progress=self.progress_event, build_options=options), 'compare')

    def start(self, func, kind):
        if self.busy:
            return
        self.started_at = time.perf_counter()
        self.finished_at = None
        self.elapsed_text.set('경과 시간 00:00:00.0')

        def timed_work():
            try:
                return func()
            finally:
                self.finished_at = time.perf_counter()

        super().start(timed_work, kind)

    def show_elapsed(self, complete=False):
        end = self.finished_at if complete and self.finished_at is not None else time.perf_counter()
        seconds, tenth = divmod(int((end - self.started_at) * 10), 10)
        hours, seconds = divmod(seconds, 3600)
        minutes, seconds = divmod(seconds, 60)
        label = '소요 시간' if complete else '경과 시간'
        self.elapsed_text.set(f'{label} {hours:02d}:{minutes:02d}:{seconds:02d}.{tenth}')

    def drain(self):
        self.root.after_cancel(self.after_id)
        if self.busy:
            self.show_elapsed()
        for _ in range(250):
            try:
                event = self.queue.get_nowait()
            except queue.Empty:
                break
            if event[0] == 'progress':
                _, completed, total, stage = event
                self.show_progress(min(99, 100 * completed // total) if total else 0, stage)
            elif event[0] in ('done', 'error'):
                self.finish_task(*event)
                if self.closing:
                    self.root.destroy()
                    return
        self.after_id = self.root.after(100, self.drain)

    def finish_task(self, event, operation, result):
        self.busy = False
        self.show_elapsed(complete=True)
        self.encoding_combo.configure(state='readonly')
        for control in self.controls:
            control.configure(state='normal')

        if event == 'error':
            self.show_progress(self.progress_percent, '작업 중단')
            message = {
                'compare': '비교를 완료하지 못했습니다. 등록한 C/H 파일과 Excel 파일을 확인하세요.',
                'sources': 'C/H 폴더를 읽지 못했습니다. 폴더 경로와 접근 권한을 확인하세요.',
                'save': '저장을 완료하지 못했습니다. 입력 파일과 출력 폴더를 확인하세요.',
                'report': '보고서를 저장하지 못했습니다. 저장 경로와 접근 권한을 확인하세요.',
            }[operation]
            self.status.set(message)
            if operation != 'compare' and not self.closing:
                messagebox.showerror('작업 중단', message, parent=self.root)
        elif operation == 'sources':
            if result:
                self.add_paths('headers', result)
            self.status.set(f'C/H 검색 완료 — 발견 {len(result)}개')
            self.show_progress(100, '폴더 검색 완료')
        elif operation == 'compare':
            files = getattr(result, 'files', ())
            errors = sum(file.status == 'ERROR' for file in files)
            incomplete = errors or any(file.partial for file in files) or result.errors
            self.show_progress(self.progress_percent if incomplete else 100,
                               '분석 종료 · 파일별 결과 확인' if incomplete else '분석 완료')
            self.session = result
            self.refresh()
            if errors:
                self.status.set(f'비교 종료 · 오류 파일 {errors}개. 파일별 상태를 확인하세요.')
            elif incomplete:
                self.status.set('비교 종료 · 일부 항목을 읽지 못했습니다. 파일별 상태를 확인하세요.')
            else:
                summary = result.catalog_summary or {}
                self.status.set(
                    f"비교 완료 · Excel {len(files)}개 · C 소스 {summary.get('files', 0)}개 · "
                    f"구조체 {summary.get('structs', 0)}개 · 비트필드 {summary.get('bitfields', 0)}개 "
                    f"(범위 확정 {summary.get('certified', 0)}개) · 변수 {summary.get('variables', 0)}개 · "
                    f"typedef {summary.get('typedefs', 0)}개")
            if getattr(result, 'build_options', None):
                options = result.build_options
                failures = len(options.get('errors', {}))
                self.status.set(self.status.get() + f" · 빌드옵션 {Path(options['file']).name}" +
                                (f' · 전처리 실패 {failures}개 C/H 문맥은 REVIEW (행 사유 확인)' if failures else ' 적용'))
            self.notebook.select(self.tabs[1])
            if incomplete:
                failed = next((node for node, (key, name) in self.nav_scopes.items()
                               if not name and self.result_index.files[key].record
                               and self.result_index.files[key].record.status in ('ERROR', 'PARTIAL')), None)
                if failed:
                    self.file_tree.selection_set(failed)
                    self.file_tree.see(failed)
                    self.show_file_status()
        elif operation == 'save':
            folder, self.outcomes = result
            incomplete = (getattr(self.outcomes, 'report_warning', '') or
                          any(item['status'] not in ('SAVED_VERIFIED', 'SKIPPED') for item in self.outcomes.values()))
            self.show_progress(self.progress_percent if incomplete else 100,
                               '저장 종료 · 결과 확인' if incomplete else '저장 작업 완료')
            if any(item.get('edit_kind') == MANUAL_REVIEW for item in self.outcomes.values()):
                self.render_rows()
            status, details = gui_save_text(folder, self.outcomes)
            self.status.set(status)
            self.save_info.configure(state='normal')
            self.save_info.delete('1.0', 'end')
            self.save_info.insert('end', details)
            self.save_info.configure(state='disabled')
            self.notebook.select(self.tabs[2])
        else:
            self.show_progress(100, '보고서 저장 완료')
            self.status.set('보고서 저장 완료')


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()
