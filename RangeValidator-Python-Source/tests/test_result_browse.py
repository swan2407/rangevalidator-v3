from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rangevalidator.engine import Result
from rangevalidator.gui import App
from rangevalidator.result_index import ResultIndex
from rangevalidator.workflow import FileStatus, canonical


class MemoryTree:
    def __init__(self):
        self.items = {}
        self.selected = ()

    def insert(self, parent, _, iid=None, **options):
        iid = iid if iid is not None else f'node{len(self.items)}'
        self.items[iid] = dict(parent=parent, **options)
        return iid

    def get_children(self, parent=''):
        return tuple(key for key, value in self.items.items() if value['parent'] == parent)

    def delete(self, *keys):
        for key in keys:
            self.delete(*self.get_children(key))
            del self.items[key]

    def selection(self):
        return self.selected

    def selection_set(self, item):
        self.selected = (item,)

    def item(self, item, **options):
        self.items[item].update(options)

    def see(self, item):
        pass

    def set(self, item, column, value):
        self.items[item][column] = value


class ResultBrowseTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name).resolve()
        self.first, self.second = [str(root / folder / 'Design.xlsx') for folder in ('one', 'two')]
        self.keys = [canonical(self.first), canonical(self.second)]
        result = Result('stable-1', self.first, '1.Process', 'H24', 'device.active',
                        'MISMATCH', '범위 불일치', existing='0~0', expected='0~1', struct='Device',
                        width=1, signed=False, source='device.h', source_line=7, candidate=True)
        self.session = SimpleNamespace(
            details=(self.first, self.second),
            results=[
                result,
                replace(result, id='stable-2', status='MATCH', expression='device.other', candidate=False),
                replace(result, id='review', status='REVIEW', expression='unknown.active',
                        struct='', source='', signed=None, candidate=False),
                replace(result, id='stable-3', sheet='2.Read', expression='sensor.active'),
                replace(result, id='stable-4', file=self.second, expression='remote.active'),
                replace(result, id='excluded', sheet='commonparameter'),
            ],
            files=[FileStatus(1, self.first, status='PARTIAL', read_status='SUCCESS'),
                   FileStatus(2, self.second, status='COMPLETED', read_status='SUCCESS')],
            coverage=[(self.first, '4.Empty', 'SKIP', '빈 시트')],
            errors=[dict(file=self.first, sheet_name=name, error_kind='SHEET_READ_ERROR', reason='읽기 실패')
                    for name in ('2.Read', '3.Failed', '표지')],
        )

    def test_index_groups_counts_and_partial_errors_by_full_path(self):
        index = ResultIndex(self.session)
        self.assertEqual(list(index.files), self.keys)
        first, second = index.files.values()
        self.assertEqual(first.counts, {'MATCH': 1, 'MISMATCH': 2, 'REVIEW': 1})
        self.assertEqual(second.counts, {'MISMATCH': 1})
        self.assertEqual(first.sheets['1.Process'].counts, {'MATCH': 1, 'MISMATCH': 1, 'REVIEW': 1})
        self.assertEqual(second.sheets['1.Process'].counts, {'MISMATCH': 1})
        self.assertEqual(first.sheets['2.Read'].status, 'PARTIAL')
        self.assertEqual(first.sheets['3.Failed'].status, 'ERROR')
        self.assertFalse(first.sheets['4.Empty'].rows)
        self.assertNotIn('commonparameter', first.sheets)
        self.assertNotIn('표지', first.sheets)
        self.assertEqual(len(self.session.results), 6)

    def app(self):
        app = App.__new__(App)
        app.session = self.session
        app.scope = (None, None)
        app.selected = set()
        app.manual_requests = {}
        app.outcomes = {}
        app.root = Mock()
        app.busy = False
        app.page = 0
        app.page_size = 1000
        app.file_tree, app.tree = MemoryTree(), MemoryTree()
        for name in ('filter', 'search', 'page_label', 'file_detail', 'count', 'selection_text', 'save_info'):
            setattr(app, name, Mock())
        app.filter.get.return_value = '전체'
        app.search.get.return_value = ''
        app.refresh()
        return app

    def test_navigation_filters_search_and_selection_use_stable_ids(self):
        app = self.app()
        first, second = app.file_tree.get_children()
        first_sheet = next(i for i in app.file_tree.get_children(first)
                           if app.file_tree.items[i]['text'] == '1.Process')
        second_sheet = app.file_tree.get_children(second)[0]
        self.assertEqual(app.tree.get_children(), ('stable-1', 'stable-2', 'review'))
        self.assertIn('device.h:7', app.tree.items['stable-1']['values'])
        self.assertIn('unsigned', app.tree.items['stable-1']['values'])
        app.toggle_item('stable-1')
        with patch('rangevalidator.gui.simpledialog.askstring', return_value=None):
            app.toggle_item('review')
        for state, expected in [('MATCH', 'stable-2'), ('MISMATCH', 'stable-1'), ('REVIEW', 'review')]:
            app.filter.get.return_value = state
            app.change_page(0, reset=True)
            self.assertEqual(app.tree.get_children(), (expected,))
            self.assertEqual(app.selected, {'stable-1'})
        app.filter.get.return_value = '전체'
        app.search.get.return_value = ' UNKNOWN '
        app.change_page(0, reset=True)
        self.assertEqual(app.tree.get_children(), ('review',))
        app.search.get.return_value = ''

        app.file_tree.selection_set(second)
        app.show_file_status()
        self.assertTrue(app.file_tree.items[second]['open'])
        self.assertFalse(app.tree.get_children())
        app.file_detail.set.assert_called_with(self.second + '\nunknown / SUCCESS / COMPLETED')
        app.file_tree.selection_set(second_sheet)
        app.show_file_status()
        self.assertEqual(app.tree.get_children(), ('stable-4',))
        app.select_visible()
        app.file_tree.selection_set(first_sheet)
        app.show_file_status()
        self.assertEqual(app.tree.items['stable-1']['values'][0], '☑')
        self.assertEqual(app.selected, {'stable-1', 'stable-4'})

    def test_tree_shows_read_errors_without_synthetic_result_rows(self):
        app = self.app()
        first = app.file_tree.get_children()[0]
        children = {app.file_tree.items[i]['text']: i for i in app.file_tree.get_children(first)}
        self.assertEqual(app.file_tree.items[children['2.Read']]['values'][-1], '읽기 불완전')
        failed = children['3.Failed']
        self.assertEqual(app.file_tree.items[failed]['values'][-1], '읽기 오류')
        app.file_tree.selection_set(failed)
        app.show_file_status()
        self.assertFalse(app.tree.get_children())
        self.assertIn('읽기 실패', app.file_detail.set.call_args.args[0])

    def test_review_input_and_selection_survive_navigation(self):
        app = self.app()
        first, second = app.file_tree.get_children()
        first_sheet = app.file_tree.selection()[0]
        self.assertEqual(app.tree.items['review']['values'][0], '☐ 수동')
        with patch('rangevalidator.gui.simpledialog.askstring', return_value='[0, 7]'):
            app.toggle_item('review')
        self.assertEqual((app.manual_requests['review'].value, app.manual_requests['review'].selected), ('0~7', True))
        app.filter.get.return_value = 'MATCH'
        app.change_page(0, reset=True)
        self.assertNotIn('review', app.tree.items)
        for node in (app.file_tree.get_children(first)[0], app.file_tree.get_children(second)[0], first_sheet):
            app.file_tree.selection_set(node)
            app.show_file_status()
        app.filter.get.return_value = 'REVIEW'
        app.change_page(0, reset=True)
        self.assertEqual(app.tree.items['review']['values'][0], '☑ 수동')
        self.assertEqual(app.tree.items['review']['values'][-1], '0~7')
        self.assertEqual(self.session.results[2].status, 'REVIEW')
        app.toggle_item('review')
        self.assertFalse(app.manual_requests['review'].selected)
        self.assertEqual(app.manual_requests['review'].value, '0~7')
        with patch('rangevalidator.gui.simpledialog.askstring', return_value='invalid'), \
                patch('rangevalidator.gui.messagebox.showinfo'):
            app.toggle_item('review')
        self.assertFalse(app.manual_requests['review'].selected)

    def test_create_filter_counts_and_selection_preserve_manual_requests(self):
        row = replace(self.session.results[0], id='create', status='CREATE', cell='H25',
                      existing=None, expected='0~7')
        self.session.results.append(row)
        app = self.app()
        index = app.result_index
        self.assertEqual(index.files[self.keys[0]].counts['CREATE'], 1)
        self.assertEqual(index.files[self.keys[0]].sheets['1.Process'].counts['CREATE'], 1)
        app.filter.get.return_value = 'CREATE'
        app.change_page(0, reset=True)
        self.assertEqual(app.tree.get_children(), ('create',))
        app.toggle_item('create')
        from rangevalidator.manual_edits import ReviewRequest
        app.manual_requests['review'] = ReviewRequest('review', '0~7', selected=True)
        for node in app.file_tree.get_children():
            app.file_tree.selection_set(node)
            app.show_file_status()
        app.refresh()
        self.assertEqual(app.selected, {'create'})
        self.assertTrue(app.manual_requests['review'].selected)
