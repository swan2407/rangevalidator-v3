"""Conservative semantic preservation check for native XLSX, never saves inputs."""
from copy import copy
import hashlib
from pathlib import Path
import zipfile
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string
from xml.etree.ElementTree import tostring


def xml(obj):
    return tostring(copy(obj).to_tree(),encoding='unicode')


def value(cell):
    v = cell.value
    if cell.data_type == 'f' and not isinstance(v,str):
        return ('formula',getattr(v,'text',None),getattr(v,'ref',None))
    return v


def style(c):
    return tuple(xml(o) for o in (c.font,c.fill,c.border,c.alignment,c.protection)) + (c.number_format,)


def manifest(path):
    """Captures semantic values/formulas/styles, dimensions, merges, links, images.
    Unsupported objects are explicitly blocked rather than assumed preserved.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        risky = [n for n in names if any(part in n.lower() for part in ('vba','activex','ctrlprops','slicer','pivot','externallink','embeddings','customxml'))]
        allowed=('[Content_Types].xml','_rels/','docProps/','xl/workbook.xml','xl/_rels/','xl/worksheets/','xl/styles.xml','xl/sharedStrings.xml','xl/theme/','xl/media/','xl/drawings/','xl/charts/','xl/calcChain.xml','xl/printerSettings/')
        risky += [n for n in names if not n.endswith('/') and not n.startswith(allowed)]
        if risky:
            raise ValueError('Unsupported workbook objects require preservation review')
        media = sorted(hashlib.sha256(z.read(n)).hexdigest() for n in names if n.startswith('xl/media/'))
        charts = sorted(z.read(n) for n in names if n.startswith('xl/charts/') and n.endswith('.xml'))
    wb = load_workbook(path,data_only=False,keep_links=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            if len(ws.conditional_formatting) or (ws.data_validations and ws.data_validations.count):
                raise ValueError('Conditional formatting/data validation requires preservation review')
            cells = {}
            for c in ws._cells.values():
                if c.value is not None or c.has_style or c.hyperlink or c.comment:
                    cells[c.coordinate] = (value(c), style(c), c.hyperlink.target if c.hyperlink else None,
                                           (c.comment.text,c.comment.author) if c.comment else None)
            columns = {}
            for label,d in ws.column_dimensions.items():
                a,b = d.min or column_index_from_string(label),d.max or column_index_from_string(label)
                for col in range(a,b+1):
                    columns[col] = (d.width,bool(d.hidden),d.outlineLevel,bool(d.collapsed))
            rows = {r:(d.height,bool(d.hidden),d.outlineLevel,bool(d.collapsed)) for r,d in ws.row_dimensions.items()
                    if d.height is not None or d.hidden or d.outlineLevel or d.collapsed}
            sheets.append(dict(name=ws.title,state=ws.sheet_state,cells=cells,columns=columns,rows=rows,
                               shapes=None, merges=sorted(str(x) for x in ws.merged_cells.ranges), freeze=str(ws.freeze_panes or ''),
                               autofilter=ws.auto_filter.ref or '',images=len(ws._images),charts=len(ws._charts)))
        return dict(sheets=sheets,media=media,charts=charts)
    finally:
        wb.close()


def verify(before, after, edits):
    """Returns detailed mismatch category, with no cell values in diagnostics."""
    expected = copy(before)
    expected['sheets'] = [dict(s,cells=dict(s['cells'])) for s in before['sheets']]
    for sheet,address,old,new in edits:
        target = next(s for s in expected['sheets'] if s['name'] == sheet)
        if address not in target['cells']:
            # New cell inherits workbook default formatting, never invent a neighbor style.
            from openpyxl import Workbook
            w = Workbook()
            default = style(w.active['A1'])
            w.close()
            target['cells'][address] = (new,default,None,None)
        else:
            prior = target['cells'][address]
            if prior[0] != old:
                raise ValueError('Pre-save value disagrees with source manifest')
            target['cells'][address] = (new,*prior[1:])
    if expected['media'] != after['media']:
        raise ValueError('Image payload preservation failed')
    if expected['charts'] != after['charts']:
        raise ValueError('Chart package changed; preservation needs review')
    if len(expected['sheets']) != len(after['sheets']):
        raise ValueError('Sheet inventory changed')
    for a,b in zip(expected['sheets'],after['sheets']):
        for attr in ('name','state','cells','columns','rows','merges','freeze','autofilter','images','charts','shapes'):
            if a[attr] != b[attr]:
                raise ValueError(f'Preservation failed: {attr}')
    return True
