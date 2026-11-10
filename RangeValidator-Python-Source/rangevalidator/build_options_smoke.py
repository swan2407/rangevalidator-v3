"""Synthetic end-to-end check usable from both Python and the frozen release."""
import json
from pathlib import Path
import tempfile
import time
import traceback


def run(receipt, example=None):
    import tkinter as tk
    from openpyxl import load_workbook
    from gui_release import App
    from selftest_patch import build
    from .build_options import load_build_options
    from .report import create_report
    from .workflow import run_comparison

    root = None
    result = {'successful': False}
    try:
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            options_file = folder / 'build.xml'
            options_file.write_text('<cdtprojectproperties><section name="settings.Macros">'
                                    '<language name="C Source File"><macro><name>OPT_PRESSURE_SENSOR</name>'
                                    '<value/></macro></language></section></cdtprojectproperties>', encoding='utf-8')
            if example:
                options_file = Path(example)
            options = load_build_options(options_file)
            assert 'OPT_PRESSURE_SENSOR' in dict(options.macros)
            header = '''#ifndef API_H
#define API_H
typedef unsigned char u1;
typedef unsigned int uint;
#ifdef OPT_PRESSURE_SENSOR
typedef struct { uint pressure:12; uint version:4; } BpsDataStateCycle_t;
#else
typedef struct { unsigned short pressure; unsigned char version; } BpsDataStateCycle_t;
#endif
extern BpsDataStateCycle_t state[1];
void Probe(u1 (*data)[16], BpsDataStateCycle_t *status);
#endif
'''
            paths, book = build(folder, {'api.h': header,
                                'motor.c': '#include "api.h"\nvoid Probe(u1 (*data)[16], BpsDataStateCycle_t *status) {}'},
                                [('state[99].pressure', 'uint', '0~4095'),
                                 ('state[2].version', 'uint', '0~7'),
                                 ('status->version', 'uint', '')], sheet='1. Probe')
            baseline = run_comparison(headers=paths, details=[book], backend='openpyxl')
            assert [r.status for r in baseline.results] == ['REVIEW'] * 3
            root = tk.Tk()
            root.withdraw()
            app = App(root)
            app.add_paths('headers', paths)
            app.add_paths('details', [book])
            app.set_build_options(options_file)
            app.compare()
            assert app.busy
            assert all(str(control.cget('state')) == 'disabled' for control in app.controls)
            deadline = time.monotonic() + 60
            while app.busy and time.monotonic() < deadline:
                root.update()
                time.sleep(0.02)
            assert not app.busy and app.session is not None, app.status.get()
            rows = app.session.results
            assert [r.status for r in rows] == ['MATCH', 'MISMATCH', 'CREATE'], [r.reason for r in rows]
            assert [r.expected for r in rows] == ['0~4095', '0~15', '0~15']
            assert all(r.source == str(folder / 'api.h') and r.source_line == 6 for r in rows)
            assert app.session.build_options['sha256'] == options.sha256
            app.session.assert_unchanged()
            report = folder / 'report.xlsx'
            create_report(app.session, {}, report)
            report_book = load_workbook(report, read_only=True)
            try:
                summary = list(report_book['Summary'].values)
                assert any(options.sha256 in row for row in summary)
            finally:
                report_book.close()
            result.update(statuses=[r.status for r in rows], expected=[r.expected for r in rows],
                          xml=options.receipt(), gui_status=app.status.get(), source_locations_verified=True,
                          array_bounds_ignored=True, array_pointer_signature_linked=True)
            app.clear_build_options()
            assert app.session is None and not app.build_options_path
            disabled = folder / 'disabled.xml'
            disabled.write_text('<cdtprojectproperties><section name="settings.Macros">'
                                '<language name="C Source File"/></section></cdtprojectproperties>')
            ordinary = run_comparison(headers=paths, details=[book], backend='openpyxl', build_options=disabled)
            assert [r.status for r in ordinary.results] == ['SKIP'] * 3
            Path(paths[-1]).write_text('#include "missing.h"\n', encoding='utf-8')
            failed = run_comparison(headers=paths, details=[book], backend='openpyxl', build_options=options)
            assert [r.status for r in failed.results] == ['REVIEW'] * 3
            assert all(not r.candidate and not r.expected for r in failed.results)
            result.update(successful=True, ordinary_branch_verified=True, missing_include_blocked=True,
                          gui_option_clear_invalidates=True, report_created=True)
    except Exception:
        result['error'] = traceback.format_exc()
        raise
    finally:
        if root is not None:
            root.destroy()
        Path(receipt).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
