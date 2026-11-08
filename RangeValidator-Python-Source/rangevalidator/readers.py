"""Read-only backends. COM is explicit: native read failure never implies DRM."""
import gc
import os
import re
from pathlib import Path
from datetime import datetime, date, time, timedelta
from .com_read import read_call, ReadBudgetExceeded, FindBoundsError
from .models import Cell, Sheet, Workbook, ReaderLimits, ReadFailure, address


def input_path(path):
    # No native file-open/stat prerequisite: policy may allow Excel-only access.
    raw = os.fspath(path)
    if not raw or '\x00' in raw:
        raise ValueError('Empty/NUL input path')
    p = Path(os.path.abspath(raw))
    if p.suffix.lower() != '.xlsx':
        raise ValueError('Only .xlsx inputs supported')
    return p


def bounds_for(cells):
    if not cells:
        return None
    return (min(c.row for c in cells), min(c.column for c in cells),
            max(c.row for c in cells), max(c.column for c in cells))


def check_limits(limits):
    if min(limits.max_sheets, limits.max_cells, limits.max_text) <= 0:
        raise ValueError('Reader limits must be positive')


def _merge(cell, ranges):
    for region in ranges:
        if region.min_row <= cell.row <= region.max_row and region.min_col <= cell.column <= region.max_col:
            return str(region), address(region.min_row, region.min_col)
    return None, None


def read_openpyxl(path, limits=ReaderLimits()):
    from time import perf_counter
    stage, formula_book, value_book = 'path', None, None
    failure, result, cleanup = None, None, []
    try:
        check_limits(limits)
        p = input_path(path)
        stage = 'dependency.openpyxl'
        from openpyxl import load_workbook
        from openpyxl.utils.datetime import to_excel, time_to_days, timedelta_to_days
        stage = 'openpyxl.open'
        # Normal mode exposes merge metadata; input file handles are read-only.
        # Never call Workbook.save(). Separate load gets stored cached values.
        tick = perf_counter()
        formula_book = load_workbook(p, read_only=False, data_only=False, keep_links=False)
        value_book = load_workbook(p, read_only=False, data_only=True, keep_links=False)
        open_seconds = perf_counter() - tick
        stage = 'openpyxl.extract'
        if len(formula_book.worksheets) > limits.max_sheets:
            raise ValueError('Sheet budget exceeded')
        result = Workbook(str(p), 'openpyxl')
        result.source_sheet_count = len(formula_book.worksheets)
        result.read_timings['open_seconds'] = open_seconds
        tick = perf_counter()
        total, text = 0, 0
        for ws in formula_book.worksheets:
            values = value_book[ws.title]
            ranges = list(ws.merged_cells.ranges)
            sheet = Sheet(ws.title, ws.sheet_state, merged_ranges=sorted(map(str, ranges)))
            # Sparse loaded-cell index avoids an enormous style-only max_row/col.
            # Pinned openpyxl 3.x private API: explicit compatibility test required.
            active = [c for c in ws._cells.values() if c.value is not None]
            bounds = bounds_for(active)
            sheet.bounds = bounds
            if bounds:
                r1, c1, r2, c2 = bounds
                total += (r2 - r1 + 1) * (c2 - c1 + 1)
                if total > limits.max_cells:
                    raise ValueError('Content rectangle cell budget exceeded')
                for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
                    for c in row:
                        is_formula = c.data_type == 'f'
                        formula = c.value if is_formula and isinstance(c.value, str) else None
                        if is_formula and formula is None:
                            # Array/data-table formulas retain a diagnostic marker.
                            formula = getattr(c.value, 'text', None) or '<non-scalar formula>'
                        value = values.cell(c.row, c.column).value
                        if isinstance(value, (datetime, date)):
                            value = to_excel(value, formula_book.epoch)
                        elif isinstance(value, time):
                            value = time_to_days(value)
                        elif isinstance(value, timedelta):
                            value = timedelta_to_days(value)
                        region, anchor = _merge(c, ranges)
                        kind = 'error' if c.data_type == 'e' else ('formula' if is_formula else 'value')
                        cell = Cell(c.row, c.column, c.coordinate, value, is_formula, formula, region, anchor, kind)
                        text += len(value) if isinstance(value, str) else 0
                        text += len(formula or '')
                        if text > limits.max_text:
                            raise ValueError('Text budget exceeded')
                        sheet.cells[cell.address] = cell
            result.sheets.append(sheet)
        result.read_timings['body_read_seconds'] = perf_counter() - tick
    except Exception as exc:
        failure = ReadFailure(stage, exc)
    finally:
        for label, book in [('values.close', value_book), ('formulas.close', formula_book)]:
            if book is not None:
                try:
                    book.close()
                except Exception as exc:
                    cleanup.append(ReadFailure(label, exc).as_dict())
    if failure or cleanup:
        failure = failure or ReadFailure('cleanup', RuntimeError('Workbook cleanup failed'))
        failure.cleanup_errors.extend(cleanup)
        raise failure
    return result


class Win32Runtime:
    """Late imports permit ordinary XLSX use without Excel/pywin32."""
    def __init__(self):
        import pythoncom
        import win32com.client
        import win32process
        self.pythoncom = pythoncom
        self.client = win32com.client
        self.process = win32process

    def initialize(self):
        self.pythoncom.CoInitialize()

    def uninitialize(self):
        self.pythoncom.CoUninitialize()

    def process_ids(self):
        return set(self.process.EnumProcesses())

    def dispatch(self):
        return self.client.DispatchEx('Excel.Application')

    def dispatch_owned(self):
        """Prove this Reader can safely configure and later quit the new process."""
        existing = self.process_ids()
        app = self.dispatch()
        try:
            pid = self.pid(app)
            if not pid or pid in existing:
                raise RuntimeError('Cannot prove new Excel ownership')
        except Exception as exc:
            raise ReadFailure('com.instance.verify', exc) from None
        return app, pid

    def pid(self, app):
        return self.process.GetWindowThreadProcessId(app.Hwnd)[1]

    def wait_exit(self, pid):
        import win32api
        import win32event
        # Synchronize-only handle; never terminate any process.
        try:
            handle = win32api.OpenProcess(0x00100000, False, pid)
        except Exception as exc:
            if getattr(exc, 'winerror', None) in (87, 1168):
                return
            raise
        try:
            if win32event.WaitForSingleObject(handle, 30_000) != win32event.WAIT_OBJECT_0:
                raise TimeoutError('Owned Excel process did not exit within 30 seconds; no process killed')
        finally:
            handle.Close()


def _com_bounds(ws, _trace=None):
    # Writer preservation calls remain strict on uncertain fallback bounds.
    from .com_values import _last_cell
    row, col = _last_cell(ws, _trace)
    return (1, 1, row, col) if row and col else None


def _merge_coordinates(region):
    m=re.fullmatch(r'([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)',region)
    if not m:raise ValueError('Unsupported COM merge address')
    def col(s):
        value=0
        for c in s:value=value*26+ord(c)-64
        return value
    bounds=(int(m[2]),col(m[1]),int(m[4]),col(m[3]))
    if not (bounds[0]<=bounds[2]<=1048576 and bounds[1]<=bounds[3]<=16384):
        raise ValueError('Invalid COM merge bounds')
    return bounds


def read_com(path, limits=ReaderLimits(), *, _runtime=None, _trace=None, function_sheets_only=False):
    from .com_values import read_workbook_values
    return read_workbook_values(path, limits, _runtime=_runtime, _trace=_trace,
                                function_sheets_only=function_sheets_only)


def read_workbook(path, backend='openpyxl', limits=ReaderLimits()):
    if backend == 'openpyxl':
        return read_openpyxl(path, limits)
    if backend == 'com':
        return read_com(path, limits)
    raise ReadFailure('backend', ValueError('Choose openpyxl or com; no automatic DRM fallback'))



def read_auto(path, limits=ReaderLimits(), diagnostics=None, *, function_sheets_only=False):
    """Native failure is only a backend failure, never a DRM diagnosis."""
    if diagnostics:
        diagnostics.emit('openpyxl','openpyxl.open')
    try:
        book = read_openpyxl(path,limits)
        if diagnostics:
            diagnostics.emit('openpyxl','openpyxl.complete','OK')
        return book
    except ReadFailure as exc:
        exc.backend='openpyxl'
        if diagnostics:
            diagnostics.emit('openpyxl',exc.stage,'FAILED',exc.code)
        # Bad format or native access denial may still be readable by approved Excel.
        # Invalid/missing paths, dependency errors, budget and extraction failures don't fallback.
        if exc.stage != 'openpyxl.open' or exc.code in (2,3):
            raise
        if diagnostics:
            diagnostics.emit('COM','fallback','ENTER',message='Native Reader failed; DRM status unknown')
        if function_sheets_only:
            return read_com(path,limits,_trace=diagnostics,function_sheets_only=True)
        return read_com(path,limits,_trace=diagnostics)
