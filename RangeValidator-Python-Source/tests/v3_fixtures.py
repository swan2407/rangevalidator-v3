"""Fresh synthetic inputs only. Shared by regression, V1 oracle and frozen checks."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side


def fixtures(root, images=False):
    root=Path(root).resolve();root.mkdir(parents=True,exist_ok=True)
    h=root/'types.h'
    h.write_text('typedef struct DeviceTag { unsigned int active:1; unsigned int mode:3; unsigned int normal; } Device_t;\ntypedef struct SensorTag { unsigned int active:3; signed int offset:8; } Sensor_t;\nextern const Device_t * const p; extern Sensor_t *q; extern Device_t obj; extern Device_t arr[10];\n',encoding='utf-8')
    common=root/'Common Parameter.xlsx'
    w=Workbook();s=w.active;s.title='Common'
    s['A1']='Global Variable'
    s.append(['File','Name','Type','Value','Range'])
    for row in [('motor.c','p','const Device_t * const','&obj','retained'),('motor.c','q','Sensor_t *','',''),('motor.c','obj','Device_t','',''),('motor.c','arr[10]','Device_t','','')]:s.append(row)
    w.save(common);w.close()
    details=[]
    for fmt in ('A','B'):
        path=root/('motor_Detail.xlsx' if fmt=='A' else 'motor.c_DetailDesign.xlsx')
        w=Workbook();s=w.active;s.title='1.Design '+fmt
        s['A1']='motor.c';s['A2']='Global(Input)'
        labels=['Name','Type','Range'] if fmt=='A' else ['Input/Output','Property','Name','Data type','Scope']
        for col,label in enumerate(labels,1):s.cell(4,col,label)
        data=[('p->active','unsigned int','0~9'),('p->mode','unsigned int','0~7'),('q->active','unsigned int','0~99'),
              ('p->normal','unsigned int','0~9'),('missing->active','unsigned int','0~9'),('p.active','unsigned int','0~9'),
              ('p->active','unsigned int','=1+1'),('p->active','unsigned int','0~9'),('p->mode','unsigned int',None),
              ('p->active',None,'0~9'),('p->active','unsigned int','0.0~1e0'),('arr[i].mode','uint','0~7'),
              ('q->offset','signed int','-128~127'),('-','uint','0~9')]
        for row,vals in enumerate(data,5):
            vals=list(vals) if fmt=='A' else ['Input','Global',*vals]
            for col,v in enumerate(vals,1):
                c=s.cell(row,col,v)
                c.font=Font(name='맑은 고딕',size=10,color='113355')
                c.fill=PatternFill('solid',fgColor='EFF6FA')
                c.alignment=Alignment(vertical='center',wrap_text=True)
                c.border=Border(bottom=Side(style='thin',color='CCCCCC'))
        rng=3 if fmt=='A' else 5
        s.merge_cells(start_row=12,start_column=rng,end_row=12,end_column=rng+1)
        s.column_dimensions['A'].width=28;s.column_dimensions['B'].width=24
        s.row_dimensions[5].height=27;s.freeze_panes='A5';s.auto_filter.ref=f'A4:{"C" if fmt=="A" else "E"}18'
        hidden=w.create_sheet('Hidden');hidden.sheet_state='hidden';hidden['A1']='Name';hidden['B1']='Type';hidden['C1']='Range';hidden['A2']='Global';hidden['A3']='p->mode';hidden['B3']='uint';hidden['C3']='0~7'
        w.create_sheet('Empty').sheet_state='veryHidden'
        if images:
            from PIL import Image
            from openpyxl.drawing.image import Image as ExcelImage
            png=root/'synthetic.png';Image.new('RGB',(16,12),(10,80,160)).save(png)
            s.add_image(ExcelImage(str(png)),'H3')
        w.save(path);w.close();details.append(path)
    return common,[h],details
