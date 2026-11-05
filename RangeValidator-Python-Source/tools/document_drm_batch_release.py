from pathlib import Path
import ast,hashlib,json
ROOT=Path(__file__).resolve().parents[1]
load=lambda name:json.loads((ROOT/'docs'/name).read_text(encoding='utf-8-sig'))
t=load('DRM_BATCH_TEST_RUN.json');d=load('DISTRIBUTION.json');v=load('DRM_BATCH_DISTRIBUTION_VALIDATION.json');scope=load('DRM_BATCH_SCOPE_CHECK.json')
current={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'rangevalidator').glob('*.py')}
assert t['successful'] and t['tests']==41 and not t['skipped'] and not t['failures'] and not t['errors']
assert current==t['source_sha256']==d['source_sha256']==v['source_sha256']
assert not t['full_suite_run'] and not t['writer_report_tests_run']
for key in ('zip','exe'):assert hashlib.sha256(Path(d[key]).read_bytes()).hexdigest()==d[key+'_sha256']==v[key+'_sha256']
# Existing Workbook data classes and Excel security/apartment runtime are unchanged.
def nodes(path):return {x.name:ast.dump(x,include_attributes=False) for x in ast.parse(path.read_text(encoding='utf-8-sig')).body if isinstance(x,(ast.ClassDef,ast.FunctionDef))}
base=ROOT/'build/drm-batch-before'
old,new=nodes(base/'models.py'),nodes(ROOT/'rangevalidator/models.py')
assert all(old[k]==new[k] for k in ('Cell','Sheet','Workbook','ReaderLimits','address'))
old,new=nodes(base/'readers.py'),nodes(ROOT/'rangevalidator/readers.py')
assert old['Win32Runtime']==new['Win32Runtime'] and old['read_openpyxl']==new['read_openpyxl']
scope.update(workbook_data_model_ast_unchanged=True,sta_ownership_runtime_ast_unchanged=True,native_reader_ast_unchanged=True)
(ROOT/'docs/DRM_BATCH_SCOPE_CHECK.json').write_text(json.dumps(scope,ensure_ascii=False,indent=2),encoding='utf-8')
report=f'''# DRM 대량 분석 실패: 확인 결과와 최신 배포

## 확인된 사실과 미확정 사항

회사 PC 자료: 전체 REVIEW 사유가 Input changed during analysis였고 일부 Open
성공 뒤 Find 실패. Find 코드/Win32/HRESULT는 None이며 예외 타입은 이전 로그에
없었습니다. 별도 실행의 merge/Quit 0x80010001 및 Close 0x80010108은 **Find의
오류 코드가 아니며 Find 원인으로 재사용하지 않았습니다.**

현재 환경에는 실제 회사 DRM 파일이 없습니다. 회사 Find의 정확한 원인,
DRM/UNC에서 왜 상태가 달라졌는지, 어떤 파일·tuple 요소가 바뀌었는지는
확정할 수 없습니다. 다음 버전의 파일별 근거와 예외/검색 단계로 확인해야 합니다.

### 소스에서 확인한 원인

이전 workflow.py의 fingerprint는 `(SHA-256, st_size, st_mtime_ns)`를 반환합니다.
분석 뒤 `any(fingerprint(p) != stamp ...)`가 True이면 모든 Result의 status를
REVIEW, candidate를 False, reason을 Input changed during analysis로 교체했습니다.
**한 파일의 수정시간만 달라도 전체 행의 사유가 덮였습니다.** 어떤 파일·요소인지
남기지 않았으며 any는 첫 차이에서 멈춰 나머지 입력 상태를 확인하지 않았습니다.

이 replace가 변수/구조체/셀 필드 자체를 지우지는 않았습니다. 빈 근거 행은
이미 Reader 실패의 가짜 REVIEW 행이거나 미인식 시트/미추출 행일 수 있습니다.
자료만으로 20개 파일 모두의 읽기 실패라고 단정하지 않습니다. 기존 코드는
Reader 오류도 `...|error`라는 빈 REVIEW Result로 만들어 변수 REVIEW와 섞었습니다.

기존 최초 fingerprint 예외는 입력 hash 불가로 후보를 막았고, 최종 fingerprint
예외는 처리하지 않아 실행 실패로 전파했습니다. **최종 예외만으로 Input changed
문자열을 만들지는 않습니다.** 이 문구는 실제 tuple 비교가 달랐을 때 만들어집니다.
DRM이 mtime만 바꾼다는 설명은 검증되지 않은 가능성입니다.

Find는 첫 검색 None을 이미 빈 시트로 반환했습니다. 실제 API 예외와 None을
같게 취급한 것이 아닙니다. 하지만 첫/마지막 행·열 단계, 파일/시트 번호가 부족했고,
예외 타입을 기록하지 않았으며, Find 이후 예산/범위 검증 예외도 com.Find로
표시할 수 있었습니다. Python TypeError/ValueError/AttributeError에는 원래 HRESULT가
없을 수 있습니다. 표준 COM hresult와 별도로 com_error.args[0]의 정수 HRESULT를
보완했고, 정보가 없는 Python 예외에 HRESULT를 만들어 넣지 않습니다.

## 수정

- 입력 상태를 STABLE / CONTENT_CHANGED(SHA 또는 크기) /
  METADATA_CHANGED(mtime만) / STATE_UNAVAILABLE로 분리.
- 조회 실패와 해시 도중 불안정은 변경으로 확정하지 않음. 전후 조회 실패,
  예외 타입, 오류 코드 및 SHA/크기/mtime 변경 여부를 입력별로 기록.
- 비교 분류·변수·구조체·멤버·셀·범위와 원래 사유를 유지. 변경/조회 불가 시
  읽기 당시 snapshot임을 표시하고 후보 및 세션 전체 저장을 보수적으로 차단.
  hash/mtime 재확인과 Writer의 기존 잠금·셀 재검증은 삭제하거나 완화하지 않음.
- 읽기 실패는 FileStatus.ERROR, read_status=FAILED, compared=False,
  추출 0 / 변수 판정 수 0으로 기록. 가짜 REVIEW Result를 만들지 않음.
- 읽기는 성공했지만 변수 행이 없으면 NO_VARIABLES/변수 비교 미실행.
  Common 실패는 종속 비교 오류로 구분하고 성공 상세 문서의 읽기 상태는 유지.
- GUI에 파일 번호·Reader·읽기/실행 상태·API/코드/타입·시트/검색·비교 여부·
  추출 수·4개 판정 수·입력 상태·저장 차단을 표시. ERROR 필터와 상세 선택 제공.
- Find를 cells/limits/first_row/last_row/first_column/last_column/coordinates로
  나누어 기록. 첫 None은 정상 빈 시트, 이후 불일치 None은 FIND_BOUNDS_INCONSISTENT.
- 읽기 예산은 com.bounds.limit / com.sheet.limit / com.text.limit 및
  READ_BUDGET_EXCEEDED로 분리하여 Find/merge COM 오류로 잘못 부르지 않음.
- 0x80010001 / 0x8001010A만 명명된 읽기 전용 API에서 동일 STA로 최대 3회,
  추가 대기 .1+.25초. 연결 끊김/기타 오류는 즉시 실패. Open/DispatchEx/
  SaveAs/Close/Quit는 재시도하지 않음. 정리 실패는 별도 기록하며 프로세스 강제 종료 없음.
- 병합 영역을 읽기 사각형/셀 예산 안의 좌표 캐시에 저장해 반복 MergeCells/
  MergeArea/Address 접근을 줄임. 캐시 조회는 dict이며 거대한 병합을 전체 시트로 확장하지 않음.
- 로그는 파일명 대신 input_role/번호, 시트 번호, 검색 단계, 예외 타입,
  외부/내부 HRESULT, 오류 분류, 시도 횟수를 기록. 원문 예외/셀 값/해시는 생략.

## 변경 범위

workflow.py, readers.py, diagnostics.py, gui.py, ReadFailure 오류 메타데이터(models.py),
새 com_read.py/error_info.py, 읽기 전용 batch_probe.py 및 GUI 진입 검사 분기,
한정 테스트/배포 도구와 관련 문서. 파서·engine·Common·Writer·보고서·보존 모듈은
해시 동일합니다. Workbook/Sheet/Cell/ReaderLimits, STA/소유권 Runtime 및
openpyxl 읽기 함수는 AST 동일합니다. V1/V2는 접근하거나 수정하지 않았습니다.
근거: DRM_BATCH_SCOPE_CHECK.json.

## 이번 검증

관련 unittest **{t['tests']}/{t['tests']} 통과**, 실패/오류/건너뜀 0, {t['seconds']}초.
전체 회귀 및 Writer/SaveAs/보고서 테스트는 실행하지 않았습니다.
목록/로그: DRM_BATCH_TEST_RUN.json, DRM_BATCH_TEST_RUN.log.

입력 tuple 요소별 차이, mtime-only, 전후 조회 거부/UNC 모의 오류,
해시 도중 상태 변화, 20개 Reader 실패와 메타 변경 동시 상황, 부분 실패,
Common 종속 오류, 정상 비트필드 비교, 빈 문서, GUI ERROR/저장 선택 차단,
Find/merge 거부 후 복구·3회 종료, 연결 끊김 1회 종료, 빈/불일치 Find,
HRESULT 없는 TypeError, args HRESULT, 읽기 예산 구분, 캐시와 기존 Reader
계약/정리, 실제 일반 Excel 읽기·누락/손상 경로를 검사했습니다.

새 ZIP을 풀어 Windows 시스템 PATH만 남기고 Python 환경변수를 제거한 EXE에서:
- 일반 문서: openpyxl SUCCESS, 비교 1건, MISMATCH 1, 구조체/멤버/범위 근거 있음.
- 손상 문서: COM fallback 뒤 Open 실패, 파일 ERROR, 비교 미실행, 추출 0,
  비트필드 REVIEW 0. 성공 문서 결과 유지.
- 파일별 GUI 표/ERROR 필터, 원본 SHA/mtime 유지, 패키지 내부 pywin32 로드,
  ZIP 전체 체크섬을 확인. Writer/보고서는 실행하지 않음.
근거: DRM_BATCH_DISTRIBUTION_VALIDATION.json. 테스트·빌드·현재 소스 해시 일치.

이는 합성 일반/손상 문서와 모의 COM/UNC 검증이며 **실제 회사 DRM·UNC·대량
성능 검증은 아닙니다.** company Find 원인이나 transient 회복을 검증했다고
보고하지 않습니다. 동기 COM 호출 자체가 정지하는 경우의 supervisor는 미지원입니다.

## 배포 및 재확인

기존 배포 폴더가 WinError 32로 잠겨 이동되지 않았습니다. 기존 앱을 종료/교체하지
않고 그대로 유지했으며 새 버전 onedir 폴더를 생성했습니다. 기존 바로가기를
사용하지 말고 **새 ZIP을 새 폴더에 전부 풀어 그 안의 EXE를 실행**하세요.

EXE: `{d['exe']}`
SHA-256: `{d['exe_sha256']}`

ZIP: `{d['zip']}`
SHA-256: `{d['zip_sha256']}`

회사에서 파일별 표와 진단 로그의 detail 번호/시트 번호/API/검색 단계/예외 타입/
분류/HRESULT 및 입력 변경 flag를 확인하세요. None HRESULT는 해당 예외에
정보가 없다는 뜻이며, 다른 API의 HRESULT로 대체하지 않습니다.

참조한 공식 API 정의:
- https://learn.microsoft.com/en-us/office/vba/api/excel.range.find
- https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-3
'''
(ROOT/'docs/DRM_BATCH_AUDIT.md').write_text(report,encoding='utf-8')
p=ROOT/'docs/TASKS.md'
with p.open('a',encoding='utf-8') as f:f.write('''

## 실제 DRM 대량 읽기 실패 보고에 대한 코드 수정 (2026-11-03)

- [x] Input changed 문자열/tuple 판정/예외 전파 원인 조사, 미확정 환경 원인 구분.
- [x] 입력별 변경·메타데이터 변경·조회 불가 분리, 근거 유지 및 기존 저장 차단 유지.
- [x] Reader 실패를 파일 ERROR로 분리, 부분 성공·미추출·미비교 상태와 GUI 집계.
- [x] Find 단계/예외 타입/HRESULT/분류/파일·시트 번호, bounded read-only 재시도와 병합 캐시.
- [x] HRESULT 없는 Python 예외/예산/불일치 Find를 COM 거부와 구분.
- [x] 관련 41개 테스트 통과. 전체/Writer/보고서 테스트는 실행하지 않음.
- [x] 새 버전 onedir EXE/ZIP, 추출본 Windows-only PATH 정상 비교+파일 ERROR 분리 확인.
- [x] 파서·엔진·Common·Writer·보고서 및 Workbook 모델 유지, V1/V2 접근/수정 없음.
- [ ] 실제 회사 Find 예외 타입/검색 단계/분류와 입력별 변경 flag 재확인.
- [ ] 실제 회사 DRM/UNC에서 읽기·retry 회복·대량 성능은 별도 미검증.

최신 근거/해시: DRM_BATCH_AUDIT.md 및 DRM_BATCH_* 기록.
''')
for rel in ('README.md','docs/TEST_RESULTS.md'):
    p=ROOT/rel;s=p.read_text(encoding='utf-8-sig')
    if s.startswith('> 최신 배포'):s=s.split('\n\n',1)[1]
    link=('docs/' if rel=='README.md' else '')+'DRM_BATCH_AUDIT.md'
    p.write_text('> 최신 배포는 DRM 대량 읽기 상태/오류 구분 수정본입니다. 이번에는 관련 **41개 테스트만** 실행했으며 전체/Writer/보고서 회귀는 실행하지 않았습니다. [최신 감사 결과]('+link+')를 확인하세요.\n\n'+s,encoding='utf-8')
print(json.dumps(dict(related_tests=t['tests'],source_test_package_equal=True,scope=scope,exe=d['exe'],exe_sha256=d['exe_sha256'],zip_sha256=d['zip_sha256']),ensure_ascii=False,indent=2))
