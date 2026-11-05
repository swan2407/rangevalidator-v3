from pathlib import Path
from openpyxl import Workbook
root=Path('build/integration-6-2').resolve()
(root/'sources').mkdir(exist_ok=True)
(root/'sources'/'synthetic.h').write_text('typedef struct { unsigned int mode:3; unsigned int state:4; unsigned int active:1; int normal; } Device; typedef struct { unsigned int mode:5; unsigned int state:2; unsigned int active:1; } Sensor; extern Device device; extern Sensor sensor;',encoding='utf-8')
for name,base,mode,state in [('BMS_합성','device','0~7','0~15'),('Sensor_합성','sensor','0~31','0~3')]:
 w=Workbook();s=w.active;s.title='1.합성검증'
 for row in [('Name','Type','Range'),(base+'.mode','unsigned int',mode),(base+'.mode','unsigned int','0~0'),(base+'.state','unsigned int',None),('unknown_p->active','unsigned int','0~1'),('unselected->active','unsigned int','0~1'),('count','int','0~1'),(base+'.normal','int','0~1'),(base+'.active','unsigned int','=1+1')]:s.append(row)
 if base=='device':
  s.append((base+'.active','unsigned int',None));s.merge_cells('B10:C10')
 t=w.create_sheet('2.추가검증');t.append(('Name','Type','Range'));t.append((base+'.active','unsigned int','0~1'))
 w.create_sheet('commonparameter')['A1']='비교 제외'
 w.save(root/(name+'.xlsx'));w.close()
print(root)
