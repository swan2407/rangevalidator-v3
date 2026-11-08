"""Safety contracts for merge discovery and actual runtime ownership checks."""
from collections import Counter
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import weakref

from rangevalidator.com_values import _merge_map
from rangevalidator.models import ReadFailure
from rangevalidator.readers import Win32Runtime, read_com
from test_readers import FakeRuntime, range_ends


class MergeSheet:
    def __init__(self):
        self.areas = [
            (2, 1, 2, 4, '$A$2:$D$2'),
            (3, 3, 70, 4, '$C$3:$D$70'),
            (65, 6, 66, 7, '$F$65:$G$66'),
        ]
        self.area_calls = Counter()
        self.queries = 0
        self.unknown_leaf = False

    def Range(self, requested):
        self.queries += 1
        first, last = range_ends(SimpleNamespace(Item=lambda r, c: SimpleNamespace(Row=r, Column=c)),
                                 requested, None)
        selected = [(r, c) for r in range(first.Row, last.Row + 1)
                    for c in range(first.Column, last.Column + 1)]
        hits = [next((area for area in self.areas if area[0] <= r <= area[2] and area[1] <= c <= area[3]),
                     None) for r, c in selected]
        if len(selected) != 1:
            return SimpleNamespace(MergeCells=None if any(hits) else False)
        if self.unknown_leaf and hits[0]:
            return SimpleNamespace(MergeCells=None)
        if not hits[0]:
            return SimpleNamespace(MergeCells=False)
        self.area_calls[hits[0][4]] += 1
        return SimpleNamespace(MergeCells=True, MergeArea=SimpleNamespace(Address=hits[0][4]))


class Excel6ReaderTests(unittest.TestCase):
    def test_geometry_crossing_partitions_and_clipped_edges(self):
        sheet = MergeSheet()
        mapping, regions = _merge_map(sheet, 68, 6)
        expected = {}
        for r0, c0, r1, c1, addr in sheet.areas:
            region = addr.replace('$', '')
            anchor = region.split(':')[0]
            for row in range(r0, min(68, r1) + 1):
                for col in range(c0, min(6, c1) + 1):
                    expected[(row, col)] = (region, anchor)
        self.assertEqual(mapping, expected)
        self.assertEqual(regions, {v[0] for v in expected.values()})
        self.assertEqual(sheet.area_calls, {area[4]: 1 for area in sheet.areas})

    def test_empty_sheet_has_one_probe_and_unknown_leaf_fails(self):
        empty = MergeSheet()
        empty.areas.clear()
        self.assertEqual(_merge_map(empty, 50, 6), ({}, set()))
        self.assertEqual(empty.queries, 1)
        unknown = MergeSheet()
        unknown.unknown_leaf = True
        with self.assertRaises(ReadFailure) as error:
            _merge_map(unknown, 50, 6)
        self.assertEqual(error.exception.stage, 'com.merge')
        self.assertEqual(error.exception.context['attempted_bounds'], (2, 1, 2, 1))

    def test_existing_pid_never_gets_configured_opened_or_quit(self):
        fake = FakeRuntime('ownership')
        runtime = Win32Runtime.__new__(Win32Runtime)
        for name in ('initialize', 'uninitialize', 'dispatch', 'pid', 'process_ids', 'wait_exit'):
            setattr(runtime, name, getattr(fake, name))
        with self.assertRaises(ReadFailure) as error:
            read_com('owned-elsewhere.xlsx', _runtime=runtime)
        self.assertEqual(error.exception.stage, 'com.instance.verify')
        self.assertEqual([event[0] for event in fake.events], ['init', 'dispatch', 'uninit'])

    def test_partial_read_drops_exception_frames_before_waiting_for_excel(self):
        class FrameOwnedReference:
            pass

        observed = []

        def fail_sheet(*args, **kwargs):
            com_proxy_in_frame = FrameOwnedReference()
            observed.append(weakref.ref(com_proxy_in_frame))
            raise ReadFailure('com.Value2', RuntimeError('synthetic partial read'))

        runtime = FakeRuntime()
        runtime.wait_exit = lambda pid: self.assertIsNone(observed[0]())
        with patch('rangevalidator.com_values._copy_sheet', fail_sheet):
            book = read_com('synthetic.xlsx', _runtime=runtime)
        self.assertTrue(book.partial)
        self.assertEqual(book.read_errors[0]['stage'], 'com.Value2')
