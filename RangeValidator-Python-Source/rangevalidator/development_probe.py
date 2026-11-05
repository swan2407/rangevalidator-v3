"""Explicit frozen integration test. Uses generated data and actual Excel only."""
from collections import Counter
import json
from pathlib import Path
import sys
import time
import tkinter as tk
from openpyxl import Workbook, load_workbook
from .gui import App, source_files
from .manual_edits import ReviewRequest
from .diagnostics import Diagnostics
from .readers import read_com
from .workflow import run_comparison
from .writer import save_selected, _write_groups


def run(receipt):
    receipt = Path(receipt).resolve()
    root = receipt.with_suffix('')
    root.mkdir()
    source = root / '합성.h'
    source.write_text('typedef struct { unsigned int mode:3; unsigned int state:4; unsigned int active:1; int normal; } Device; typedef struct { unsigned int mode:5; unsigned int state:2; unsigned int active:1; } Sensor; extern Device device; extern Sensor sensor;', encoding='utf-8')
    details = []
    for name, base, mode in [('BMS_합성', 'device', '0~7'), ('Sensor_합성', 'sensor', '0~31')]:
        path = root / (name + '.xlsx')
        w = Workbook()
        s = w.active
        s.title = '1.합성검증'
        for row in [('Name','Type','Range'), (base+'.mode','unsigned int',mode),
                    (base+'.mode','unsigned int','0~0'),(base+'.state','unsigned int',None),
                    ('unknown_p->active','unsigned int','0~1'),('unselected->active','unsigned int','0~1'),
                    ('count','int','0~1'),(base+'.normal','int','0~1'),
                    (base+'.active','unsigned int','=1+1')]:
            s.append(row)
        if base == 'device':
            s.append((base+'.active','unsigned int',None))
            s.merge_cells('B10:C10')
        s = w.create_sheet('2.추가검증')
        s.append(('Name','Type','Range'))
        s.append((base+'.active','unsigned int','0~1'))
        w.create_sheet('commonparameter')['A1'] = '비교 제외'
        w.save(path)
        w.close()
        details.append(path)
    before = {str(p): p.read_bytes() for p in details}
    window = tk.Tk()
    app = App(window)
    result = dict(frozen=getattr(sys,'frozen',False), executable=sys.executable,
                  cwd=str(Path.cwd()), direct_dialog_automation=False)
    try:
        app.add_paths('headers', [str(source)])
        app.add_paths('details', list(map(str, details)))
        diag = Diagnostics(lambda event: app.queue.put(('log', event)))
        app.last_diagnostics = diag
        app.start(lambda: run_comparison(headers=[source],details=details,backend='com',
                                         diagnostics=diag,progress=app.progress_event), 'compare')
        deadline = time.monotonic() + 180
        while app.busy and time.monotonic() < deadline:
            window.update()
            time.sleep(.02)
        if app.busy or app.session is None:
            raise RuntimeError('Analysis did not complete')
        session = app.session
        counts = Counter(r.status for r in session.results)
        assert counts == {'MATCH':4, 'MISMATCH':2, 'CREATE':2, 'REVIEW':7}, counts
        assert app.progress_percent == 100
        for node, (file, name) in app.nav_scopes.items():
            if name:
                app.file_tree.selection_set(node)
                app.show_file_status()
                for status in ('MATCH','MISMATCH','REVIEW','CREATE'):
                    app.filter.set(status)
                    app.change_page(0,reset=True)
                    expected = app.result_index.visible(file,name,status,'')
                    assert len(app.tree.get_children()) == len(expected)
                app.search.set('unknown')
                app.filter.set('전체')
                app.change_page(0,reset=True)
                assert all('unknown' in app.result_index.rows[rid].expression for rid in app.tree.get_children())
                app.search.set('')
        selected = [r.id for r in session.results if r.candidate]
        manual = [ReviewRequest(r.id,'0~7',True) for r in session.results
                  if r.expression == 'unknown_p->active' or r.cell in ('C9','C10')]
        app.manual_requests = {r.row_id:r for r in manual}
        app.selected = set(selected)
        app.start(lambda: save_selected(session,selected,root/'outputs',app.save_progress_event,
                                        manual_requests=manual), 'save')
        deadline = time.monotonic()+240
        while app.busy and time.monotonic()<deadline:
            window.update()
            time.sleep(.02)
        if app.busy:
            raise RuntimeError('Save did not complete')
        outcomes = app.outcomes
        assert Counter(o['status'] for o in outcomes.values()) == {'SAVED_VERIFIED':6,'SKIPPED':3}, outcomes
        assert {o['edit_kind'] for o in outcomes.values()} == {'AUTO_MISMATCH','AUTO_CREATE','MANUAL_REVIEW'}
        assert app.progress_percent == 100
        outputs = sorted({o['output'] for o in outcomes.values() if o.get('output')})
        for path in outputs:
            book = read_com(path)
            s = next(s for s in book.sheets if s.name=='1.합성검증')
            sensor = Path(path).name.startswith('Sensor')
            expected = {'C2':'0~31' if sensor else '0~7','C3':'0~31' if sensor else '0~7',
                        'C4':'0~3' if sensor else '0~15','C5':'0~7','C6':'0~1'}
            for address,value in expected.items():
                assert s.cells[address].value == value, (path,address,s.cells[address])
                assert not s.cells[address].is_formula
            assert s.cells['C9'].is_formula
        assert all(p.read_bytes()==before[str(p)] for p in details)
        report = Path(outputs[0]).parent / 'RangeValidator_Report.xlsx'
        w = load_workbook(report)
        try:
            assert w.sheetnames == ['Summary','Mismatch','Review','Errors','Results']
            for name in ('Mismatch','Review'):
                assert w[name].freeze_panes=='A2'
                assert w[name].auto_filter.ref==w[name].dimensions
            changes = list(w['Mismatch'].values)[1:]
            assert {r[12] for r in changes} == {'AUTO_MISMATCH','AUTO_CREATE','MANUAL_REVIEW'}
            assert sum(r[15]=='SAVED_VERIFIED' for r in changes)==6
            assert sum(r[15]=='SKIPPED' for r in changes)==3
        finally:
            w.close()
        collision = root / 'collision'
        collision.mkdir()
        (collision/'BMS_합성_fixed.xlsx').write_bytes(b'collision sentinel')
        isolated = _write_groups(session, [r for r in session.results if r.candidate],collision)
        assert all(o['status']==('SAVE_FAILED' if Path(o['file']).name.startswith('BMS') else 'SAVED_VERIFIED')
                   for o in isolated.values())
        assert all(p.read_bytes()==before[str(p)] for p in details)
        from .diagnostics import format_event
        (root/'diagnostics.log').write_text(''.join(format_event(e) for e in diag.events),encoding='utf-8')
        result.update(analysis=dict(counts),save=dict(Counter(o['status'] for o in outcomes.values())),
                      outputs=outputs,report=str(report),original_preserved=True,
                      gui_callback_navigation_filter_search=True,progress_complete=True,
                      collision_isolation=True,source_readers=session.readers,
                      kinds=sorted({o['edit_kind'] for o in outcomes.values()}))
    finally:
        window.destroy()
    receipt.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result
