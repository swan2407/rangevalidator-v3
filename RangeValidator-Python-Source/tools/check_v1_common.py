"""Current V1 sources, read-only; all compilation/fixtures/results stay in V3."""
from pathlib import Path
import hashlib,json,os,subprocess,sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT));sys.path.insert(0,str(ROOT/'tests'))
from v3_fixtures import fixtures
from rangevalidator.workflow import run_comparison
from openpyxl import load_workbook

def main():
 v1=Path('C:/Dev/RangeValidator');out=ROOT/'build/v1-common';out.mkdir(parents=True,exist_ok=True)
 sources=['bitfield_parser','type_matcher','variable_expression_parser','common_parameter_reader','common_type_resolver','common_range_analysis','range_comparator','xlsx_reader','xlsx_package']
 refs=[v1/'src'/f'{n}.cpp' for n in sources]+list((v1/'include').glob('*.h'))
 hashes=lambda:{str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in refs}
 before=hashes()
 env=dict(os.environ,PATH='C:/msys64/ucrt64/bin;'+os.environ.get('PATH',''),TEMP=str(out),TMP=str(out),TMPDIR=str(out),GIT_OPTIONAL_LOCKS='0')
 git=lambda:subprocess.run(['git','--no-optional-locks','-C',str(v1),'status','--porcelain=v1','--untracked-files=no'],capture_output=True,env=env).stdout.decode('utf-8',errors='replace')
 git_before=git()
 extracted=out/'common_range_analysis_readonly.cpp'
 original=(v1/'src/common_range_analysis.cpp').read_text(encoding='utf-8-sig')
 extracted.write_text(original.split('AutoFixResult save_common_corrections(')[0]+'\n}\n',encoding='utf-8')
 exe=out/'probe.exe'
 cmd=['C:/msys64/ucrt64/bin/g++.exe','-std=c++17','-O1','-static','-ffunction-sections','-fdata-sections','-Wl,--gc-sections','-I'+str(v1/'include'),str(ROOT/'tools/v1_common_probe.cpp'),*[str(extracted if n=='common_range_analysis' else v1/'src'/f'{n}.cpp') for n in sources],'-lexpat','-lz','-o',str(exe)]
 compile=subprocess.run(cmd,capture_output=True,env=env)
 if compile.returncode:
  print(compile.stderr.decode(errors='replace'));return 1
 records=[];diffs=[]
 for scenario in ('basic','duplicate_common','duplicate_type','conditional','scalar_reversed','no_section','merged_header','classification_formula','internal_conflict'):
  c,h,ds=fixtures(out/scenario)
  if scenario=='duplicate_common':
   w=load_workbook(c);w.active.append(['motor.c','p','Device_t *','','']);w.save(c);w.close()
  if scenario in ('duplicate_type','conditional'):
   p=out/scenario/'other.h';p.write_text('typedef struct { unsigned int active:5; } Sensor_t;' if scenario=='duplicate_type' else '#if CONFIG\nstruct X { unsigned int x:1; };\n#endif');h.append(p)
  for d in ds:
   w=load_workbook(d);s=w.active;col=3 if 'Design A'==s.title else 5
   if scenario=='scalar_reversed':s.cell(5,col,3);s.cell(7,col,'7~0');s.cell(13,col,'3e0')
   if scenario=='no_section':s['A2']=None
   if scenario=='merged_header':s.merge_cells(start_row=4,start_column=col,end_row=4,end_column=col+1)
   if scenario=='classification_formula' and col==5:s['A5']='="Input"'
   if scenario=='internal_conflict':s['A1']='other.c'
   w.save(d);w.close()
  py=run_comparison(c,h,ds,'openpyxl')
  for d in ds:
   run=subprocess.run([str(exe),str(c),str(d),*map(str,h)],capture_output=True,encoding='utf-8',env=env,check=True)
   expected=[line.split('\t')[:5] for line in run.stdout.splitlines() if not line.startswith('BLOCK\t')]
   actual=[[r.sheet,r.cell,r.expression,r.status,r.expected] for r in py.results if r.file==str(d) and r.cell]
   # V1 may retain field evidence on REVIEW; compare decision/coordinates/expression for all, range for certified rows.
   norm=lambda values:[row[:4]+([row[4]] if row[3] in ('MATCH','MISMATCH') else ['']) for row in values]
   equal=norm(expected)==norm(actual)
   rec=dict(scenario=scenario,file=d.name,rows=len(expected),equal=equal)
   records.append(rec)
   if not equal:diffs.append(dict(rec,v1=expected,python=actual))
 report=dict(documents=len(records),rows=sum(r['rows'] for r in records),differences=diffs,reference_files_unchanged=before==hashes(),git_state_unchanged=git_before==git(),source_sha256=before,results=records)
 (ROOT/'docs/V1_COMMON_DIFFERENTIAL.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps({k:report[k] for k in ('documents','rows','reference_files_unchanged','git_state_unchanged')}));print('differences',len(diffs))
 if diffs:print(json.dumps(diffs,ensure_ascii=False,indent=2))
 return int(bool(diffs) or not report['reference_files_unchanged'] or not report['git_state_unchanged'])
if __name__=='__main__':raise SystemExit(main())

