"""Deterministic workbook fixtures. Ground truth is assigned at write time."""
import random

from openpyxl import Workbook
from openpyxl.styles import Font


EXCLUDED = ('Common Parameter', 'Cover', 'Summary', 'Notes', 'Index', 'History')


def _book():
    book = Workbook()
    book.remove(book.active)
    for name in EXCLUDED:
        sheet = book.create_sheet(name)
        sheet.append(['Name', 'Type', 'Range'])
        sheet.append(['obj.v001', 'unsigned int', '0~3'])
    return book


def _table(sheet, truth, filename, fmt, row, base, count, *, merged=False,
           unresolved=False, editable_only=False):
    section = '함수 변수 정보' if fmt == 'A' else ('Return' if unresolved else 'Global(Input)')
    sheet.cell(row, base, section)
    if merged:
        sheet.merge_cells(start_row=row, start_column=base, end_row=row, end_column=base + 7)
    row += 2
    name_col, type_col, range_col = (base + 2, base + 4, base + 6) if fmt == 'A' else (base, base + 2, base + 4)
    for col, value in ((name_col, 'Na\nme'), (type_col, 'Data Type' if fmt == 'A' else 'Type')):
        sheet.cell(row, col, value)
    if fmt == 'A':
        sheet.cell(row, base, 'Input/Output')
        sheet.cell(row, base + 1, 'Property')
    if not unresolved:
        sheet.cell(row, range_col, 'Scope' if fmt == 'A' else 'Range')
    if merged:
        sheet.merge_cells(start_row=row, start_column=name_col, end_row=row, end_column=name_col + 1)
        if not unresolved:
            sheet.merge_cells(start_row=row, start_column=range_col, end_row=row, end_column=range_col + 1)
    row += 1
    for index in range(count):
        if index % 13 == 4:
            row += 2
        name = f'obj.v{index % 200 + 1:03d}'
        name_cell = sheet.cell(row, name_col, name)
        sheet.cell(row, type_col, 'unsigned int')
        if fmt == 'A':
            sheet.cell(row, base, 'Input')
            sheet.cell(row, base + 1, 'Global')
        original = ('0~3', None)[index % 2] if editable_only else ('0~3', None, '0~7')[index % 3]
        status = 'CREATE' if original is None else 'MATCH' if original == '0~7' else 'MISMATCH'
        reason = ''
        target = None
        if unresolved:
            status, reason = 'UNRESOLVED', 'RANGE_CELL_UNRESOLVED'
            original = None
        else:
            if not editable_only and index == 2:
                original, status, reason = '=IF(1=0,1,"")', 'REVIEW', 'FORMULA'
            cell = sheet.cell(row, range_col, original)
            target = cell.coordinate
            if merged and index == 0:
                sheet.merge_cells(start_row=row, start_column=range_col,
                                  end_row=row, end_column=range_col + 1)
            if merged and not editable_only and index == 3:
                sheet.cell(row, range_col - 1, '0~3')
                sheet.merge_cells(start_row=row, start_column=range_col - 1,
                                  end_row=row, end_column=range_col)
                original, status, reason = None, 'REVIEW', 'MERGE_NONANCHOR'
        truth.append(dict(file=filename, sheet=sheet.title, format=fmt, section=section,
                          name=name, name_cell=name_cell.coordinate, range_cell=target,
                          original_range=original, status=status,
                          candidate=status in ('MISMATCH', 'CREATE'), reason=reason))
        row += 1
    return row


def build_mixed(path, count):
    book = _book()
    order = ['A', 'B'] * (count // 2)
    random.Random(805 + count).shuffle(order)
    truth = []
    for index, fmt in enumerate(order):
        sheet = book.create_sheet(f'{index + 1}. Function_{index:03d}')
        size = 1000 if index in (1, 33) else 500 if index in (4, 28, 71, 102) else (1, 9, 18, 35, 52, 70)[index % 6]
        base = 2 + index % 4
        row = _table(sheet, truth, path.name, fmt, 4 + index % 8, base, size,
                     merged=index % 4 == 0)
        if index % 23 == 0:
            _table(sheet, truth, path.name, fmt, row + 3, base + 1, 3, unresolved=True)
        if index in (0, 1):
            sheet.cell(900_000, 16_384).font = Font(bold=True)
    book.save(path)
    book.close()
    return truth


def build_write_file(path, count, fmt='B'):
    book = _book()
    sheet = book.create_sheet('1. Save_target')
    truth = []
    end = _table(sheet, truth, path.name, fmt, 3, 2, count, merged=True, editable_only=True)
    _table(sheet, truth, path.name, fmt, end + 3, 3, 4, merged=True)
    end = max(sheet.max_row, end) + 4
    _table(sheet, truth, path.name, fmt, end, 4, 2, unresolved=True)
    book.save(path)
    book.close()
    return truth


def build_limits(path):
    """Confirmed content past column cap plus a 3,100-row valid sibling."""
    book = _book()
    wide = book.create_sheet('1. Wide_limit')
    truth = []
    _table(wide, truth, path.name, 'B', 3, 2, 2)
    wide.cell(20, 257, 'Out of bounds content')
    long = book.create_sheet('2. Long_table')
    _table(long, truth, path.name, 'B', 3, 2, 3100)
    long.cell(1_000_000, 16_384).font = Font(bold=True)
    book.save(path)
    book.close()
    return truth
