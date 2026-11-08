"""Synthetic B section layouts and their independent expected coordinates."""
from openpyxl import Workbook


def _headers(sheet, row, columns, labels):
    for column, label in zip(columns, labels):
        sheet.cell(row, column, label)


def _variable(sheet, row, columns, field, value, section, section_cell, *, status=None):
    name_col, type_col, range_col = columns
    name = sheet.cell(row, name_col, f'obj.v{field:03d}')
    sheet.cell(row, type_col, 'unsigned int')
    target = sheet.cell(row, range_col, value)
    return dict(sheet=sheet.title, section=section, section_cell=section_cell,
                name=name.value, name_cell=name.coordinate, range_cell=target.coordinate,
                comparison_status=status or ('CREATE' if value is None else 'MATCH' if value == '0~7' else 'MISMATCH'))


def build_b_workbook(path, sheet_count):
    book = Workbook()
    book.active.title = 'Common Parameter'
    book.active['A1'] = 'Excluded content'
    truth = []
    for index in range(sheet_count):
        sheet = book.create_sheet(f'{index + 1}. B_function')
        row = 4 + index % 7
        sheet.cell(row, 2, 'The Global Input description is not a section')
        row += 2
        # Repeated titles and headers get their own columns, including reversed Range/Name.
        for group, (title, written, count, columns) in enumerate((
            ('Parameter(Input)', ' Pa ra meter \n( In put ) ', 0, (2, 4, 6)),
            ('Global(Output)', '\u3000Glo BAL ( Out put )\t', (1, 10, 200, 500)[index % 4], (8, 5, 2)),
            ('Global(Output)', 'Global(Output)', 1, (3, 5, 7)),
            ('Return', 'RETURN', 1, (5, 7, 9)),
        )):
            title_cell = sheet.cell(row, 2, written)
            sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
            row += 2
            labels = (' Na\nme ', 'Ty\tpe', '범위' if group % 2 else 'Ran ge')
            _headers(sheet, row, columns, labels)
            # Both Name and Range can be horizontal merge anchors.
            for column in (columns[0], columns[2]):
                sheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
            row += 1
            for offset in range(count):
                if offset % 11 == 3:
                    row += 2
                value = ('0~3', None, '0~7')[offset % 3]
                truth.append(_variable(sheet, row, columns, offset % 200 + 1, value,
                                       title, title_cell.coordinate))
                row += 1
            if group == 2:
                # A repeated header without a second title stays within the same section.
                columns = (6, 4, 2)
                _headers(sheet, row, columns, ('NAME', 'TYPE', 'RAN\nGE'))
                row += 1
                truth.append(_variable(sheet, row, columns, 1, '0~3', title, title_cell.coordinate))
                row += 1
            row += 3
    book.save(path)
    book.close()
    return truth


def build_b_safety(path):
    book = Workbook()
    sheet = book.active
    sheet.title = '1. Safety'
    book.create_sheet('Common Parameter')['A1'] = 'Excluded content'
    sheet['B3'] = 'Parameter (Output)'
    _headers(sheet, 5, (2, 3, 5), ('Na me', 'Type', '범위'))
    truth = [
        _variable(sheet, 6, (2, 3, 5), 1, '0~3', 'Parameter(Output)', 'B3'),
        _variable(sheet, 8, (2, 3, 5), 2, None, 'Parameter(Output)', 'B3'),
        _variable(sheet, 9, (2, 3, 5), 3, '=IF(1=0,"x","")', 'Parameter(Output)', 'B3', status='REVIEW'),
        _variable(sheet, 10, (2, 3, 5), 4, None, 'Parameter(Output)', 'B3', status='REVIEW'),
    ]
    sheet['B7'], sheet['C7'], sheet['E7'] = 'get_status()', 'unsigned int', '0~3'
    sheet.merge_cells('E6:F6')
    sheet['D10'] = '0~3'
    sheet.merge_cells('D10:E10')

    sheet['D12'] = 'Global (Input)'
    _headers(sheet, 14, (4, 6), ('Name', 'Type'))
    sheet['D15'], sheet['F15'] = 'obj.v005', 'unsigned int'
    sheet['C17'] = 'Local'
    _headers(sheet, 18, (3, 4, 6, 8), ('Name', 'Type', 'Range', '범위'))
    sheet['C19'], sheet['D19'], sheet['F19'], sheet['H19'] = 'obj.v006', 'unsigned int', '0~3', '0~7'
    sheet['C21'] = 'Return'
    _headers(sheet, 22, (3, 4, 8), ('Name', 'Type', 'Range'))
    truth.append(_variable(sheet, 23, (3, 4, 8), 7, '0~3', 'Return', 'C21'))
    sheet['C25'] = 'Description of Global (Input) and Return'
    _variable(sheet, 26, (3, 4, 8), 8, '0~3', '', None)
    _headers(sheet, 28, (3, 4, 8), ('Name', 'Type', '범위'))
    sheet['C29'], sheet['D29'], sheet['H29'] = 'obj.v009', 'unsigned int', '0~3'
    book.save(path)
    book.close()
    return truth
