"""Korean Tk/ttk GUI. All UI access stays on the main thread."""
from dataclasses import asdict, replace
import json
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from .diagnostics import Diagnostics, format_event, STAGES
from .workflow import run_comparison
from .writer import save_selected, export_all
from .save_display import gui_save_text
from .result_index import ResultIndex, STATES
from .manual_edits import MANUAL_REVIEW, ReviewRequest, is_review_target, manual_value
from .report import create_report


LABELS={'MATCH':'일치','MISMATCH':'불일치','REVIEW':'검토 필요','CREATE':'신규 입력','SKIP':'일반 멤버'}


SOURCE_EXTENSIONS = {'.c', '.h', '.hpp', '.inc'}
PROGRESS_LABELS = {'compare': '분석', 'save': '수정·저장', 'report': '보고서', 'sources': 'C/H 폴더 검색'}


def source_files(folder):
    """Find supported C/H inputs recursively without following directory symlinks."""
    def failed(exc):
        raise exc

    return sorted((str(Path(root) / name) for root, _, names in os.walk(folder, onerror=failed)
                   for name in names if Path(name).suffix.casefold() in SOURCE_EXTENSIONS),
                  key=str.casefold)


class App:
    def __init__(self,root):
        self.root=root
        root.title('RangeValidator Python V3.1 — 비트필드 범위 검증')
        root.geometry('1320x820')
        root.minsize(1000,640)
        style=ttk.Style(root)
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        style.configure('.',font=('맑은 고딕',9))
        self.queue=queue.Queue()
        self.busy=False
        self.closing=False
        self.session=None
        self.page=0
        self.page_size=1000
        self.selected=set()
        self.manual_requests = {}
        self.outcomes={}
        self.headers=[]
        self.details=[]
        self.controls=[]
        self.notebook=ttk.Notebook(root)
        self.notebook.pack(fill='both',expand=True,padx=10,pady=10)
        self.tabs=[]
        for title in ('파일 등록','비교 결과','수정 및 저장','진단 로그'):
            tab=ttk.Frame(self.notebook,padding=10)
            self.notebook.add(tab,text=title)
            self.tabs.append(tab)
        self.build_inputs()
        self.build_results()
        self.build_save()
        self.build_logs()
        bottom=ttk.Frame(root,padding=(10,0,10,10))
        bottom.pack(fill='x')
        self.status=tk.StringVar(value='파일을 등록한 뒤 비교를 실행하세요.')
        ttk.Button(bottom,text='진단 로그 보기',command=self.show_diagnostics).pack(side='right',padx=8)
        progress_area = ttk.Frame(bottom)
        progress_area.pack(side='right', padx=8)
        self.progress_text = tk.StringVar(value='분석 진행률: 0%')
        self.progress_stage = tk.StringVar(value='현재 작업: 대기')
        self.progress = ttk.Progressbar(progress_area, mode='determinate', maximum=100, length=220)
        ttk.Label(progress_area, textvariable=self.progress_text).pack(anchor='w')
        self.progress.pack(fill='x')
        ttk.Label(progress_area, textvariable=self.progress_stage, wraplength=350).pack(anchor='w')
        ttk.Label(bottom, textvariable=self.status, wraplength=640).pack(side='left', fill='x', expand=True)
        root.protocol('WM_DELETE_WINDOW',self.close)
        self.after_id=root.after(100,self.drain)
        root.bind('<Destroy>',self.on_destroy,add='+')

    def button(self,parent,text,command):
        b=ttk.Button(parent,text=text,command=command)
        self.controls.append(b)
        return b

    def invalidate(self):
        self.session=None
        self.page=0
        self.selected.clear()
        self.manual_requests.clear()
        self.outcomes={}
        self.refresh()

    def build_inputs(self):
        tab=self.tabs[0]
        for title,role in [('C 소스/헤더 파일 (여러 개 — 구조체 정의 .h 와 변수 선언이 있는 .c/.h)','headers'),('상세설계서 XLSX (여러 개)','details')]:
            box=ttk.LabelFrame(tab,text=title,padding=8)
            box.pack(fill='both',expand=True,pady=5)
            bar=ttk.Frame(box)
            bar.pack(fill='x')
            self.button(bar,'파일 추가',lambda r=role:self.add(r)).pack(side='left',padx=3)
            if role == 'headers':
                self.button(bar,'폴더 추가',self.add_source_folder).pack(side='left',padx=3)
            self.button(bar,'선택 삭제',lambda r=role:self.remove(r)).pack(side='left',padx=3)
            self.button(bar,'초기화',lambda r=role:self.clear(r)).pack(side='left',padx=3)
            frame=ttk.Frame(box)
            frame.pack(fill='both',expand=True)
            lb=tk.Listbox(frame,selectmode='extended',height=4,exportselection=False)
            lb.pack(side='left',fill='both',expand=True)
            scroll=ttk.Scrollbar(frame,command=lb.yview)
            scroll.pack(side='right',fill='y')
            lb.configure(yscrollcommand=scroll.set)
            setattr(self,role+'_list',lb)
        bar=ttk.Frame(tab)
        bar.pack(fill='x',pady=10)
        self.button(bar,'비교 실행',self.compare).pack(side='right')
        ttk.Label(bar,text='일반 XLSX 우선 읽기 → 실패 시 승인된 Excel COM 사용 / 원본은 읽기 전용').pack(side='left')

        self.encoding=tk.StringVar(value='utf-8-sig')
        ttk.Label(bar,text='C 헤더 인코딩').pack(side='left',padx=5)
        self.encoding_combo=ttk.Combobox(bar,textvariable=self.encoding,values=['utf-8-sig','cp949'],state='readonly',width=12)
        self.encoding_combo.pack(side='left')
        self.encoding_combo.bind('<<ComboboxSelected>>',lambda e:self.invalidate())

    def add(self,role):
        if self.busy:return
        types=[('C 소스/헤더','*.h *.c *.hpp *.inc'),('모든 파일','*.*')] if role=='headers' else [('Excel','*.xlsx')]
        ps=filedialog.askopenfilenames(parent=self.root,filetypes=types)
        if not ps:return
        self.add_paths(role, ps)

    def add_source_folder(self):
        if self.busy:
            return
        folder = filedialog.askdirectory(parent=self.root, title='C/H 소스 폴더 선택')
        if folder:
            self.start(lambda: source_files(folder), 'sources')

    def add_paths(self, role, paths):
        current = getattr(self, role)
        identities = {str(Path(p).resolve()).casefold() for p in current}
        for path in paths:
            identity = str(Path(path).resolve()).casefold()
            if identity not in identities:
                current.append(path)
                identities.add(identity)
        self.sync_inputs()
        self.invalidate()

    def remove(self,role):
        if self.busy:return
        lb=getattr(self,role+'_list')
        indexes=set(lb.curselection())
        setattr(self,role,[p for i,p in enumerate(getattr(self,role)) if i not in indexes])
        self.sync_inputs();self.invalidate()

    def clear(self,role):
        if self.busy:return
        setattr(self,role,[])
        self.sync_inputs();self.invalidate()

    def sync_inputs(self):
        for role in ('headers','details'):
            lb=getattr(self,role+'_list')
            lb.delete(0,'end')
            for p in getattr(self,role):
                lb.insert('end',p)

    def build_results(self):
        tab = self.tabs[1]
        self.result_index = ResultIndex(None)
        self.scope = (None, None)
        self.count = tk.StringVar()
        ttk.Label(tab, textvariable=self.count).pack(anchor='w', pady=4)
        panes = ttk.Panedwindow(tab, orient='horizontal')
        panes.pack(fill='both', expand=True)
        nav = ttk.Frame(panes, padding=4)
        detail = ttk.Frame(panes, padding=4)
        panes.add(nav, weight=1)
        panes.add(detail, weight=3)
        self.file_tree = ttk.Treeview(nav, columns=(*STATES, '상태'), selectmode='browse', height=8)
        self.file_tree.heading('#0', text='파일 / 함수 시트')
        self.file_tree.column('#0', width=175, minwidth=110)
        for col in (*STATES, '상태'):
            self.file_tree.heading(col, text='일반 멤버' if col == 'SKIP' else col)
            self.file_tree.column(col, width=75, minwidth=50, anchor='center')
        self.file_tree.grid(row=0, column=0, sticky='nsew')
        fy = ttk.Scrollbar(nav, orient='vertical', command=self.file_tree.yview)
        fx = ttk.Scrollbar(nav, orient='horizontal', command=self.file_tree.xview)
        fy.grid(row=0, column=1, sticky='ns')
        fx.grid(row=1, column=0, sticky='ew')
        self.file_tree.configure(yscrollcommand=fy.set, xscrollcommand=fx.set)
        nav.rowconfigure(0, weight=1)
        nav.columnconfigure(0, weight=1)
        self.file_detail = tk.StringVar()
        ttk.Label(nav, textvariable=self.file_detail, wraplength=340, justify='left').grid(
            row=2, column=0, sticky='w', pady=4)
        self.file_tree.bind('<<TreeviewSelect>>', self.show_file_status)

        bar = ttk.Frame(detail)
        bar.pack(fill='x', pady=4)
        self.filter = tk.StringVar(value='전체')
        ttk.Label(bar, text='판정').pack(side='left')
        combo = ttk.Combobox(bar, textvariable=self.filter,
                             values=('전체', *(('일반 멤버' if s == 'SKIP' else s) for s in STATES)),
                             state='readonly', width=11)
        combo.pack(side='left', padx=5)
        combo.bind('<<ComboboxSelected>>', lambda e: self.change_page(0, reset=True))
        self.search = tk.StringVar()
        ttk.Label(bar, text='변수 검색').pack(side='left')
        ttk.Entry(bar, textvariable=self.search, width=18).pack(side='left', padx=5)
        self.search.trace_add('write', lambda *_: self.change_page(0, reset=True))
        actions = ttk.Frame(detail)
        actions.pack(fill='x', pady=4)
        self.button(actions,'표시된 안전 후보 선택',self.select_visible).pack(side='left',padx=3)
        self.button(actions,'선택 해제',self.unselect).pack(side='left',padx=3)
        self.button(actions,'계산 범위 전체 내보내기',self.export).pack(side='right',padx=3)
        paging = ttk.Frame(detail)
        paging.pack(fill='x', pady=4)
        self.button(paging,'이전 페이지',lambda:self.change_page(-1)).pack(side='left',padx=3)
        self.button(paging,'다음 페이지',lambda:self.change_page(1)).pack(side='left',padx=3)
        self.page_label = tk.StringVar()
        ttk.Label(paging,textvariable=self.page_label).pack(side='left')
        cols = ('선택','판정','변수','원본 구조체','C/H 파일 및 선언 위치','비트폭',
                '부호','계산 범위','기존 범위','셀','사유','REVIEW 수동 입력')
        frame=ttk.Frame(detail);frame.pack(fill='both',expand=True)
        self.tree=ttk.Treeview(frame,columns=cols,show='headings',selectmode='extended')
        for c in cols:
            self.tree.heading(c,text=c)
            self.tree.column(c,width=105 if c!='사유' else 320,minwidth=50)
        self.tree.column('선택',width=75)
        self.tree.column('비트폭',width=55)
        self.tree.column('C/H 파일 및 선언 위치',width=240)
        y=ttk.Scrollbar(frame,orient='vertical',command=self.tree.yview)
        x=ttk.Scrollbar(frame,orient='horizontal',command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set,xscrollcommand=x.set)
        self.tree.grid(row=0,column=0,sticky='nsew');y.grid(row=0,column=1,sticky='ns');x.grid(row=1,column=0,sticky='ew')
        frame.rowconfigure(0,weight=1);frame.columnconfigure(0,weight=1)
        self.tree.bind('<Button-1>',self.toggle)
        self.tree.bind('<space>',self.toggle_selected)

    def build_save(self):
        tab=self.tabs[2]
        self.selection_text=tk.StringVar(value='선택된 수정 항목: 0개')
        ttk.Label(tab,textvariable=self.selection_text).pack(anchor='w',pady=10)
        ttk.Label(tab,text='선택한 MISMATCH/CREATE와 사용자 승인 REVIEW를 별도 _fixed 파일에 저장합니다.').pack(anchor='w',pady=5)
        ttk.Label(tab,text='"비교 결과" 탭의 [계산 범위 전체 내보내기] 는 선택 없이 모든 파일의 MISMATCH/CREATE를 확인 후 기입합니다. 여기서는 선택한 칸만 기입합니다.').pack(anchor='w',pady=5)
        self.output=tk.StringVar()
        row=ttk.Frame(tab);row.pack(fill='x',pady=15)
        ttk.Entry(row,textvariable=self.output,state='readonly').pack(side='left',fill='x',expand=True)
        self.button(row,'출력 폴더 선택',self.choose_output).pack(side='right',padx=5)
        row=ttk.Frame(tab);row.pack(fill='x',pady=10)
        self.button(row,'선택 수정본 및 보고서 저장',self.save).pack(side='left',padx=5)
        self.button(row,'보고서만 저장',self.report_only).pack(side='left',padx=5)
        self.save_info=tk.Text(tab,height=20,wrap='word',state='disabled')
        self.save_info.pack(fill='both',expand=True,pady=10)

    def build_logs(self):
        tab=self.tabs[3]
        self.last_diagnostics=None
        self.last_run_result=None
        self.diagnostic_window=None
        self.diagnostic_log=None
        self.log_font_size=11
        self.diagnostic_summary=tk.StringVar(value='실행 전 — 기본 마스킹 사용 / 실제 회사 DRM 미검증')
        self.diagnostic_result=tk.StringVar(value='마지막 실행 결과: 없음')
        ttk.Label(tab,textvariable=self.diagnostic_result,wraplength=1150).pack(anchor='w',pady=4)
        ttk.Label(tab,textvariable=self.diagnostic_summary,wraplength=1150,justify='left').pack(anchor='w',pady=4)
        bar=ttk.Frame(tab);bar.pack(fill='x')
        ttk.Label(bar,text='민감정보 기본 마스킹: 경로·셀 내용·원문 예외는 로그에 기록하지 않습니다.').pack(side='left')
        ttk.Button(bar,text='선택/전체 복사',command=self.copy_logs).pack(side='right',padx=4)
        self.button(bar,'로그 파일 저장 (선택)',self.export_logs).pack(side='right',padx=4)
        frame=ttk.Frame(tab);frame.pack(fill='both',expand=True,pady=8)
        self.log=tk.Text(frame,wrap='none',state='disabled',font=('맑은 고딕',self.log_font_size))
        self.log.grid(row=0,column=0,sticky='nsew')
        for orient,axis,command in [('vertical','y',self.log.yview),('horizontal','x',self.log.xview)]:
            sc=ttk.Scrollbar(frame,orient=orient,command=command)
            sc.grid(row=0 if axis=='y' else 1,column=1 if axis=='y' else 0,sticky='ns' if axis=='y' else 'ew')
            self.log.configure(**{axis+'scrollcommand':sc.set})
        frame.rowconfigure(0,weight=1);frame.columnconfigure(0,weight=1)

    def refresh(self):
        """Rebuild the browse index only when replacing or explicitly refreshing results."""
        previous = self.scope
        self.result_index = ResultIndex(self.session)
        self.file_tree.delete(*self.file_tree.get_children())
        self.nav_scopes = {}
        nodes = {}
        default = None
        for file, entry in self.result_index.files.items():
            f = entry.record
            node = self.file_tree.insert('', 'end', text=Path(entry.path).name,
                values=(*[entry.counts[state] for state in STATES], f.status if f else ''))
            self.nav_scopes[node] = (file, None)
            nodes[file, None] = node
            for name, sheet in entry.sheets.items():
                state = {'REVIEW': '형식 확인 필요', 'ERROR': '읽기 오류',
                         'PARTIAL': '읽기 불완전'}.get(sheet.status, '')
                child = self.file_tree.insert(node, 'end', text=name,
                    values=(*[sheet.counts[key] for key in STATES], state))
                self.nav_scopes[child] = (file, name)
                nodes[file, name] = child
                if default is None and sheet.rows:
                    default = child
        selected = nodes.get(previous) or default or next(iter(self.nav_scopes), None)
        if selected:
            self.file_tree.selection_set(selected)
            self.file_tree.see(selected)
            self.show_file_status()
        else:
            self.scope = (None, None)
            self.file_detail.set('비교 결과가 없습니다.')
            self.change_page(0, reset=True)
        self.refresh_selection()

    def render_rows(self):
        self.tree.delete(*self.tree.get_children())
        visible = self.result_index.visible(*self.scope, self.filter.get(), self.search.get())
        pages = max(1, (len(visible) + self.page_size - 1) // self.page_size)
        self.page = min(self.page, pages - 1)
        self.page_label.set(f'{self.page+1}/{pages} 페이지 · {len(visible)}개')
        for r in visible[self.page*self.page_size:(self.page+1)*self.page_size]:
            source = r.source + (f':{r.source_line}' if r.source and r.source_line else '')
            signed = '' if r.signed is None else 'signed' if r.signed else 'unsigned'
            self.tree.insert('', 'end', iid=r.id, values=(
                self.selection_mark(r),
                LABELS.get(r.status, r.status) + f' ({r.status})', r.expression, r.struct,
                source, r.width or '', signed, r.expected, r.existing, r.cell, r.reason, self.manual_text(r)))

    def selection_mark(self, row):
        if is_review_target(row):
            request = self.manual_requests.get(row.id)
            return '☑ 수동' if request and request.selected else '☐ 수동'
        return '☑' if row.id in self.selected else '☐' if self.is_editable(row) else ''

    def manual_text(self, row):
        request = self.manual_requests.get(row.id)
        if request is None:
            return ''
        outcome = self.outcomes.get(row.id, {})
        completed = (outcome.get('status') == 'SAVED_VERIFIED'
                     and outcome.get('manual_value') == request.value)
        return f'{request.value} (수동 수정 완료)' if completed else request.value

    def toggle_review(self, row):
        previous = self.manual_requests.get(row.id)
        if previous and previous.selected:
            self.manual_requests[row.id] = replace(previous, selected=False)
            return True
        value = simpledialog.askstring(
            'REVIEW 수동 수정 선택',
            f'{row.sheet}!{row.cell} · {row.expression}\n'
            f'검토 사유: {row.reason}\n기존 값: {row.existing}\n'
            '사용자 승인 범위 (예: 0~7)를 입력하고 확인을 누르면 수정 대상으로 선택됩니다.',
            initialvalue=previous.value if previous else '', parent=self.root)
        if value is None:
            return False
        try:
            normalized = manual_value(value)
        except ValueError as exc:
            self.manual_requests[row.id] = ReviewRequest(row.id, value)
            messagebox.showinfo('수동 수정값 확인', str(exc), parent=self.root)
            return True
        self.manual_requests[row.id] = ReviewRequest(row.id, normalized, selected=True)
        return True

    def refresh_selection(self):
        manual = [request for request in self.manual_requests.values() if request.selected]
        counts = self.result_index.counts
        errors = sum(f.record is not None and f.record.status == 'ERROR'
                     for f in self.result_index.files.values())
        self.count.set(' | '.join(f'{"일반 멤버" if state == "SKIP" else state} {counts[state]}' for state in STATES)
                       + f' | 파일 ERROR {errors} | 자동 선택 {len(self.selected)} | 수동 선택 {len(manual)}')
        self.selection_text.set(f'자동 MISMATCH/CREATE {len(self.selected)}개 / REVIEW 수동 {len(manual)}개')
        if not self.busy:
            self.save_info.configure(state='normal')
            self.save_info.delete('1.0','end')
            for rid in sorted(self.selected):
                r = self.result_index.rows.get(rid)
                if r:
                    self.save_info.insert('end', f'{Path(r.file).name} / {r.sheet}!{r.cell}: {r.existing} → {r.expected}\n')
            for request in manual:
                r = self.result_index.rows.get(request.row_id)
                if r:
                    self.save_info.insert('end', f'[REVIEW 수동] {Path(r.file).name} / {r.sheet}!{r.cell}: {r.existing} → {request.value}\n')
            self.save_info.configure(state='disabled')

    def show_file_status(self,event=None):
        selection = self.file_tree.selection()
        if not selection or selection[0] not in self.nav_scopes:
            return
        scope = self.nav_scopes[selection[0]]
        file, name = scope
        entry = self.result_index.files[file]
        f = entry.record
        text = [entry.path]
        if f:
            code = f'0x{f.code&0xffffffff:08X}' if isinstance(f.code, int) else ''
            text.append(f'{f.reader} / {f.read_status} / {f.status}')
            details = (f.reason,) if f.read_status == 'FAILED' else (f.failed_api, code, f.exception_type, f.reason)
            text.extend(part for part in details if part)
            if f.sheet_index or f.search_phase:
                text.append(f'시트 #{f.sheet_index} / {f.search_phase}')
        if name:
            text.extend(part for part in (name, entry.sheets[name].reason) if part)
        else:
            self.file_tree.item(selection[0], open=True)
        self.file_detail.set('\n'.join(text))
        self.change_page(0, reset=(self.scope != scope), scope=scope)

    def change_page(self,delta,reset=False,scope=None):
        if scope is not None:
            self.scope = scope
        self.page = 0 if reset else max(0, self.page + delta)
        self.render_rows()

    @staticmethod
    def is_editable(row):
        return row.status in ('MISMATCH', 'CREATE') and row.candidate

    def toggle(self,event):
        if self.busy or self.tree.identify_column(event.x) != '#1':
            return
        item = self.tree.identify_row(event.y)
        if item:
            self.toggle_item(item)

    def toggle_item(self,item):
        self.toggle_items((item,))

    def toggle_items(self, items):
        changed = False
        for item in items:
            row = self.result_index.rows.get(item)
            if row and self.is_editable(row):
                self.selected.symmetric_difference_update((row.id,))
                self.tree.set(item, '선택', '☑' if row.id in self.selected else '☐')
                changed = True
            elif row and is_review_target(row) and self.toggle_review(row):
                self.tree.set(item, '선택', self.selection_mark(row))
                self.tree.set(item, 'REVIEW 수동 입력', self.manual_text(row))
                changed = True
        if changed:
            self.refresh_selection()

    def toggle_selected(self,event):
        if not self.busy and self.session:
            self.toggle_items(self.tree.selection())
        return 'break'

    def select_visible(self):
        if self.busy or not self.session:
            return
        for item in self.tree.get_children():
            if self.is_editable(self.result_index.rows[item]):
                self.selected.add(item)
                self.tree.set(item, '선택', '☑')
        self.refresh_selection()

    def unselect(self):
        if self.busy:
            return
        self.selected.clear()
        self.manual_requests = {rid: replace(request, selected=False)
                                for rid, request in self.manual_requests.items()}
        self.render_rows()
        self.refresh_selection()

    def choose_output(self):
        p=filedialog.askdirectory(parent=self.root,title='출력 폴더 선택')
        if p:self.output.set(p)

    def start(self,func,kind):
        if self.busy:return
        self.busy=True
        for c in self.controls:c.configure(state='disabled')
        self.encoding_combo.configure(state='disabled')
        self.progress_kind = kind
        self.progress_percent = 0
        self.show_progress(0, '파일 검색 중 · 전체 파일 수 확인 전' if kind == 'sources' else '작업 준비 중')
        def work():
            try:self.queue.put(('done',kind,func()))
            except Exception as exc:self.queue.put(('error',kind,exc))
        threading.Thread(target=work,name='RangeValidator-STA',daemon=False).start()

    def show_progress(self, percent, stage):
        self.progress_percent = max(self.progress_percent, percent)
        self.progress.configure(value=self.progress_percent)
        self.progress_text.set(f'{PROGRESS_LABELS[self.progress_kind]} 진행률: {self.progress_percent}%')
        self.progress_stage.set(f'현재 작업: {stage}')

    def progress_event(self,i,n,name):
        self.queue.put(('progress',i,n,name))

    def save_progress_event(self, i, n, name):
        # File work finishes before the report is written.
        stage = '저장 결과 보고서 생성 중' if i == n else f'Excel 수정·저장·검증 중 — {name}'
        self.progress_event(i, n + 1, stage)

    def compare(self):
        if not self.headers or not self.details:
            messagebox.showinfo('파일 등록','C 헤더와 상세설계서는 등록해야 합니다.',parent=self.root);return
        self.invalidate()
        diag=Diagnostics(lambda e:self.queue.put(('log',e)))
        self.last_diagnostics=diag
        self.diagnostic_summary.set('현재 실행 ID: '+diag.run_id+' — 읽기 진행 중')
        headers,details=list(self.headers),list(self.details)
        encoding=self.encoding.get()
        self.start(lambda:run_comparison(None,headers,details,diagnostics=diag,encoding=encoding,progress=self.progress_event),'compare')

    def save(self):
        manual = tuple(request for request in self.manual_requests.values() if request.selected)
        if not self.session or not (self.selected or manual) or not self.output.get():
            messagebox.showinfo('수정 및 저장','안전한 수정 후보와 출력 폴더를 선택하세요.',parent=self.root);return
        records={f.file:f for f in getattr(self.session,'files',[])}
        chosen=[r for r in self.session.results if r.id in self.selected]
        if len(chosen)!=len(self.selected) or any(not r.candidate or (r.file in records and (records[r.file].status=='ERROR' or records[r.file].read_status!='SUCCESS' or records[r.file].save_blocked)) for r in chosen):
            messagebox.showerror('저장 차단','읽기 실패 또는 안전하지 않은 대상은 저장할 수 없습니다.',parent=self.root)
            return
        if not messagebox.askyesno('선택 항목 확인',f'자동 MISMATCH/CREATE {len(self.selected)}개와 REVIEW 수동 {len(manual)}개를 새 파일에 저장하시겠습니까?\n원본은 변경하지 않습니다.',parent=self.root):return
        session,selected,parent=self.session,set(self.selected),self.output.get()
        self.start(lambda:save_selected(session,selected,parent,self.save_progress_event,manual_requests=manual),'save')

    def export(self):
        """모든 상세설계서의 불일치·공란(MISMATCH/CREATE) 칸에 계산 범위를 기입해 새 폴더로 저장"""
        if self.busy or not self.session:
            messagebox.showinfo('내보내기','먼저 비교를 실행하세요.',parent=self.root);return
        n=sum(1 for r in self.session.results if r.status in ('MISMATCH','CREATE') and r.candidate)
        files=len({r.file for r in self.session.results if r.status in ('MISMATCH','CREATE') and r.candidate})
        if not n:
            messagebox.showinfo('내보내기','수정할 MISMATCH/CREATE가 없습니다. 내보낼 파일이 없습니다.',parent=self.root);return
        p=filedialog.askdirectory(parent=self.root,title='내보낼 폴더 선택 (하위에 RangeValidator_날짜시각 폴더를 만듭니다)')
        if not p:return
        if not messagebox.askyesno('전체 내보내기',f'상세설계서 {files}개 파일의 {n}개 MISMATCH/CREATE 칸에 계산 범위를 기입해 새 파일로 저장합니다.\n원본은 변경하지 않습니다. 진행할까요?',parent=self.root):return
        self.output.set(p)
        session=self.session
        self.start(lambda:export_all(session,p,self.save_progress_event),'save')

    def report_only(self):
        if not self.session:return
        path=filedialog.asksaveasfilename(parent=self.root,initialfile='RangeValidator_Report.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if not path:return
        session, outcomes = self.session, self.outcomes
        manual, selected = tuple(self.manual_requests.values()), tuple(self.selected)
        self.start(lambda:create_report(session,outcomes,path,manual_requests=manual,selected_ids=selected),'report')

    def on_destroy(self,event):
        if event.widget is self.root:
            try:self.root.after_cancel(self.after_id)
            except tk.TclError:pass

    def drain(self):
        try:self.root.after_cancel(self.after_id)
        except tk.TclError:pass
        try:
            for _ in range(250):
                event=self.queue.get_nowait()
                kind=event[0]
                if kind=='log':
                    e=event[1]
                    self.append_diagnostic(format_event(e))
                    if e.status=='FAILED':
                        self.diagnostic_summary.set(f'실행 ID: {e.run_id} | 실패 API: {e.api or e.stage} | 입력 {e.input_role} #{e.file_index} / 시트 #{e.sheet_index} / {e.search_phase} | 예외: {e.exception_type or "없음"} | 분류: {e.error_class or "해당 없음"} | HRESULT: {e.hresult or "없음"} | Win32: {e.win32_code}\n마지막 성공: {e.last_success or "없음"} | Open 진입: {"예" if e.open_entered else "아니오"}')
                        self.show_diagnostics()
                    elif e.stage=='com.process.exit':
                        self.diagnostic_summary.set(self.diagnostic_summary.get().split('\nExcel 종료:')[0]+'\nExcel 종료: '+('성공' if e.status=='OK' else '확인 필요'))
                    elif e.stage in ('openpyxl.complete','com.complete') and e.status=='OK':
                        self.diagnostic_summary.set(f'실행 ID: {e.run_id} | Reader: {e.reader} | 읽기 성공 | 마지막 성공: {e.last_success} | Excel 종료: {e.excel_exit or "해당 없음"}')
                elif kind=='progress':
                    _, completed, total, stage = event
                    percent = min(99, 100 * completed // total) if total else 0
                    self.show_progress(percent, stage)
                else:
                    self.busy=False
                    self.encoding_combo.configure(state='readonly')
                    for c in self.controls:c.configure(state='normal')
                    if kind=='error':
                        self.show_progress(self.progress_percent, '작업 중단 · 진단 로그 확인')
                        exc=event[2]
                        stage=getattr(exc,'stage','작업 검증')
                        self.status.set(f'확인 필요 — {stage} / 코드 {getattr(exc,"code",None)}')
                        self.notebook.select(self.tabs[3])
                        self.last_run_result={'operation':event[1],'status':'FAILED','stage':stage if stage in STAGES else 'gui.run','code':getattr(exc,'code',None)}
                        self.diagnostic_result.set('마지막 실행 결과: 실패 — '+self.last_run_result['stage'])
                        self.show_diagnostics()
                        if event[1]!='compare':
                            messagebox.showerror('작업 중단','입력 변경 또는 읽기/저장/안전 조건을 확인할 수 없어 작업을 중단했습니다. 진단 로그를 확인하세요.',parent=self.root)
                    elif event[1] == 'sources':
                        paths = event[2]
                        if paths:
                            self.add_paths('headers', paths)
                        self.status.set(f'C/H 검색 완료 — 발견 {len(paths)}개')
                        self.show_progress(100, '폴더 검색 완료')
                    elif event[1]=='compare':
                        files=getattr(event[2],'files',[])
                        errors=sum(f.status=='ERROR' for f in files)
                        incomplete = errors or any(f.partial for f in files) or event[2].errors
                        self.show_progress(self.progress_percent if incomplete else 100,
                                           '분석 종료 · 읽기/비교 오류 확인' if incomplete else '분석 완료')
                        self.last_run_result={'operation':'compare','status':'PARTIAL_ERROR' if errors else 'COMPLETED',
                            'file_errors':errors,'read_errors':sum(f.read_status=='FAILED' for f in files)}
                        self.diagnostic_result.set(f'마지막 실행 결과: 파일 ERROR {errors} — 파일별 표와 로그 확인')
                        cs=getattr(event[2],'catalog_summary',{}) or {}
                        self.session=event[2];self.refresh()
                        self.status.set(f"비교 완료 — C 소스 {cs.get('files',0)}개: 구조체 {cs.get('structs',0)} / 비트필드 {cs.get('bitfields',0)} (범위 확정 {cs.get('certified',0)}) / 변수 {cs.get('variables',0)} / typedef {cs.get('typedefs',0)} | 파일 ERROR {errors}")
                        if not any(e.status=='FAILED' for e in self.session.diagnostics.events):self.notebook.select(self.tabs[1])
                    elif event[1]=='save':
                        folder,self.outcomes=event[2]
                        incomplete = (getattr(self.outcomes, 'report_warning', '') or
                                      any(v['status'] not in ('SAVED_VERIFIED', 'SKIPPED')
                                          for v in self.outcomes.values()))
                        self.show_progress(self.progress_percent if incomplete else 100,
                                           '저장 종료 · 실패 또는 검증 보류 확인' if incomplete else '저장 작업 완료')
                        if any(v.get('edit_kind') == MANUAL_REVIEW for v in self.outcomes.values()):
                            self.render_rows()
                        status, details = gui_save_text(folder, self.outcomes)
                        self.status.set(status)
                        self.save_info.configure(state='normal');self.save_info.delete('1.0','end');self.save_info.insert('end',details);self.save_info.configure(state='disabled')
                        self.notebook.select(self.tabs[2])
                    else:
                        self.show_progress(100, '보고서 저장 완료')
                        self.status.set('보고서 저장 완료')
                    if self.closing:
                        self.root.destroy();return
        except queue.Empty:pass
        self.after_id=self.root.after(100,self.drain)

    def append_diagnostic(self,text):
        for widget in (self.log,self.diagnostic_log):
            if widget is not None and widget.winfo_exists():
                widget.configure(state='normal')
                widget.insert('end',text)
                if int(widget.index('end-1c').split('.')[0])>5000:widget.delete('1.0','1001.0')
                widget.see('end')
                widget.configure(state='disabled')

    def show_diagnostics(self):
        self.notebook.select(self.tabs[3])
        if self.diagnostic_window is not None and self.diagnostic_window.winfo_exists():
            self.diagnostic_window.deiconify();self.diagnostic_window.lift();return
        win=tk.Toplevel(self.root)
        self.diagnostic_window=win
        win.title('RangeValidator — 진단 로그 보기 (기본 마스킹)')
        win.geometry('1120x720');win.minsize(800,480)
        ttk.Label(win,textvariable=self.diagnostic_result,wraplength=1050,justify='left').pack(anchor='w',padx=12,pady=5)
        ttk.Label(win,textvariable=self.diagnostic_summary,wraplength=1050,justify='left').pack(anchor='w',padx=12,pady=5)
        bar=ttk.Frame(win);bar.pack(fill='x',padx=10,pady=5)
        ttk.Button(bar,text='복사',command=self.copy_logs).pack(side='left',padx=4)
        ttk.Button(bar,text='글자 크게',command=lambda:self.zoom_logs(1)).pack(side='left',padx=4)
        ttk.Button(bar,text='글자 작게',command=lambda:self.zoom_logs(-1)).pack(side='left',padx=4)
        ttk.Label(bar,text='경로·파일명·셀 값·원문 오류 생략 / 창 크기 조절 후 화면 촬영 가능').pack(side='left',padx=8)
        frame=ttk.Frame(win);frame.pack(fill='both',expand=True,padx=10,pady=(0,10))
        log=tk.Text(frame,wrap='none',font=('맑은 고딕',self.log_font_size))
        log.grid(row=0,column=0,sticky='nsew')
        self.diagnostic_log=log
        log.insert('end',self.log.get('1.0','end-1c'));log.configure(state='disabled')
        for orient,axis,command in [('vertical','y',log.yview),('horizontal','x',log.xview)]:
            scroll=ttk.Scrollbar(frame,orient=orient,command=command)
            scroll.grid(row=0 if axis=='y' else 1,column=1 if axis=='y' else 0,sticky='ns' if axis=='y' else 'ew')
            log.configure(**{axis+'scrollcommand':scroll.set})
        frame.rowconfigure(0,weight=1);frame.columnconfigure(0,weight=1)

    def zoom_logs(self,delta):
        self.log_font_size=max(9,min(20,self.log_font_size+delta))
        for widget in (self.log,self.diagnostic_log):
            if widget is not None and widget.winfo_exists():widget.configure(font=('맑은 고딕',self.log_font_size))

    def copy_logs(self):
        widget=self.diagnostic_log if self.diagnostic_log is not None and self.diagnostic_log.winfo_exists() else self.log
        try:data=widget.get('sel.first','sel.last')
        except tk.TclError:data=widget.get('1.0','end')
        self.root.clipboard_clear();self.root.clipboard_append(data)

    def export_logs(self):
        p=filedialog.asksaveasfilename(parent=self.root,defaultextension='.log',filetypes=[('로그','*.log')])
        if p:
            try:
                with open(p,'x',encoding='utf-8') as f:f.write(self.log.get('1.0','end'))
            except OSError:messagebox.showerror('로그 저장','기존 파일 또는 접근 권한을 확인하세요.',parent=self.root)

    def close(self):
        if self.busy:
            self.closing=True
            self.status.set('작업 및 Excel 정리 후 종료합니다. 잠시 기다려 주세요.')
        else:self.root.destroy()


def main():
    root=tk.Tk()
    App(root)
    root.mainloop()
