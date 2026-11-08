"""Process sampling and non-invasive timing for the actual COM stress runner."""
from collections import Counter
from contextlib import ExitStack
from functools import wraps
import os
import threading
from time import perf_counter
from unittest.mock import patch

import win32api
import win32process

from rangevalidator import a_extraction, b_extraction, com_values, engine, readers, workflow
from rangevalidator.diagnostics import Diagnostics


class Trace(Diagnostics):
    def __init__(self, probe):
        super().__init__()
        self.probe = probe
        self.started = {}

    def emit(self, reader, stage, status='ENTER', code=None, **metadata):
        now = perf_counter()
        if reader == 'writer':
            if stage == 'writer.saveas' and status == 'ENTER':
                self.started[stage] = now
            elif stage == 'writer.saveas' and status == 'OK' and stage in self.started:
                self.probe.times['save_as_seconds'] += now - self.started.pop(stage)
            elif stage == 'writer.reopen' and status == 'ENTER':
                self.started[stage] = now
            elif stage == 'writer.cleanup' and status == 'OK' and 'writer.verify' in self.started:
                self.probe.times['writer_verify_with_cleanup_seconds'] += now - self.started.pop('writer.verify')
        if reader == 'COM' and stage == 'writer.verify' and status == 'ENTER' and 'writer.reopen' in self.started:
            self.probe.times['writer_reopen_with_startup_seconds'] += now - self.started.pop('writer.reopen')
            self.started[stage] = now
        if reader == 'COM' and status == 'ENTER':
            self.probe.operations[metadata.get('api', stage)] += 1
        return super().emit(reader, stage, status, code, **metadata)


class Probe:
    def __init__(self):
        self.times = Counter()
        self.operations = Counter()
        self.pids = set()
        self.peak = {'python': 0, 'owned_excel': 0}
        self.samples = Counter()
        self.stop = threading.Event()
        self.stack = ExitStack()

    def _timed(self, module, name, metric):
        original = getattr(module, name)

        @wraps(original)
        def measured(*args, **kwargs):
            start = perf_counter()
            try:
                return original(*args, **kwargs)
            finally:
                self.times[metric] += perf_counter() - start
                self.operations[name] += 1

        self.stack.enter_context(patch.object(module, name, measured))

    def _sample(self):
        while not self.stop.is_set():
            for label, pids in [('python', [os.getpid()]), ('owned_excel', list(self.pids))]:
                total = 0
                for pid in pids:
                    try:
                        handle = win32api.OpenProcess(0x0400 | 0x0010, False, pid)
                        try:
                            total += win32process.GetProcessMemoryInfo(handle)['WorkingSetSize']
                        finally:
                            handle.Close()
                    except Exception:
                        continue  # A process may exit between samples.
                if total:
                    self.peak[label] = max(self.peak[label], total)
                    self.samples[label] += 1
            self.stop.wait(0.05)

    def __enter__(self):
        pid_of = readers.Win32Runtime.pid

        def record_pid(runtime, app):
            pid = pid_of(runtime, app)
            self.pids.add(pid)
            return pid

        self.stack.enter_context(patch.object(readers.Win32Runtime, 'pid', record_pid))
        for module, name, metric in (
            (com_values, '_excel', 'reader_com_init_seconds'),
            (com_values, '_open', 'reader_open_with_init_seconds'),
            (com_values, 'close_excel', 'reader_com_cleanup_seconds'),
            (com_values, '_count_a', 'count_a_seconds'),
            (com_values, '_merge_map', 'merge_scan_seconds'),
            (com_values, '_merge_state', 'merge_state_seconds'),
            (a_extraction, 'scan_a', 'a_scan_seconds'),
            (b_extraction, 'scan_b', 'b_scan_seconds'),
            (workflow, 'compare_book', 'comparison_pipeline_seconds'),
            (engine, '_catalog_match', 'c_binding_seconds'),
            (engine, 'interval_status', 'range_decision_seconds'),
            (readers.Win32Runtime, 'wait_exit', 'owned_excel_exit_wait_seconds'),
        ):
            self._timed(module, name, metric)
        self.thread = threading.Thread(target=self._sample, daemon=True)
        self.thread.start()
        return self

    def __exit__(self, *exc):
        self.stop.set()
        self.thread.join()
        self.stack.close()

    def result(self):
        times = dict(self.times)
        if 'reader_open_with_init_seconds' in times:
            times['workbooks_open_seconds'] = max(0, times['reader_open_with_init_seconds']
                                                   - times.get('reader_com_init_seconds', 0))
        return dict(seconds=times, observed_operations=dict(self.operations),
                    peak_working_set_bytes={key: value or None for key, value in self.peak.items()},
                    memory_samples=dict(self.samples), memory_interval_seconds=0.05,
                    excel_pids=sorted(self.pids))
