from pathlib import Path
import hashlib,json
root=Path(__file__).resolve().parents[1]
load=lambda name:json.loads((root/'docs'/name).read_text(encoding='utf-8-sig'))
t=load('DIAGNOSTIC_TEST_RUN.json');d=load('DISTRIBUTION.json');v=load('DIAGNOSTIC_DISTRIBUTION_VALIDATION.json');scope=load('DIAGNOSTIC_SCOPE_CHECK.json')
current={str(p.relative_to(root)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (root/'rangevalidator').glob('*.py')}
assert t['tests']==27 and t['successful'] and not t['failures'] and not t['errors'] and not t['skipped']
assert not t['full_suite_run'] and not t['writer_tests_run'] and not t['report_tests_run']
assert current==t['source_sha256']==d['source_sha256']==v['source_sha256']
assert scope['reader_non_diagnostic_ast_identical'] and scope['engine_writer_report_workflow_models_byte_identical']
for kind in ('exe','zip'):
    assert hashlib.sha256(Path(d[kind]).read_bytes()).hexdigest()==d[kind+'_sha256']==v[kind+'_sha256']
probe=v['result']['diagnostic-test']
report=f'''# 최신 진단 로그 감사 및 보완 결과

대상: C:/Dev/RangeValidator-Python, 최신 PyInstaller onedir EXE/ZIP.
V1/V2는 이번 작업에서 접근하거나 수정하지 않았습니다.

| 확인 항목 | 수정 전 소스/기존 EXE | 최신 EXE 확인 결과 |
|---|---|---|
| GUI 로그 | 진단 탭 있음 | 탭 + 항상 접근 가능한 진단 로그 보기 버튼 + 별도 창 |
| 일반/COM 읽기 성공·실패 | 기본 기록 있음, 단계별 성공 부족 | Reader별 진입/성공/실패 및 fallback 양쪽 기록 |
| DispatchEx/Open/Find/Value2/Close/Quit | Close/Quit 기록 누락 | 각 단계 구분, 미실행은 SKIP, 정리 오류 별도 기록 |
| HRESULT/API/마지막 성공 | 일부 제공, 실패 후 성공 추론 문제 가능 | HRESULT·Win32 구분, 실패 단계/마지막 성공/Excel 종료 표시 |
| 실패 후 유지 | 텍스트는 남지만 진단 객체 직접 보관 없음 | 마지막 진단 객체·실행 결과 보관, 입력 변경 후에도 로그 유지 |
| 기본 마스킹 | 알려진 호출부는 원문 생략, 메시지 필터 미강제 | Reader/단계/상태/메시지 허용 목록, 경로·파일명·셀·원문 예외 생략 |
| 배포 EXE 포함 | 기존 4탭 GUI/pywin32 실행 확인 | ZIP 추출 EXE의 실제 진단/Reader 검사 통과 |

실제 회사 DRM 파일은 미검증입니다. DRM용 COM과 동일한 코드 경로의 합성
일반 문서/손상 파일을 확인했으며, 읽기 실패를 DRM 거부로 분류하지 않습니다.

## 변경 파일과 범위

- rangevalidator/diagnostics.py: 단계 완료/실패 상태, HRESULT·Win32, 마스킹,
  마지막 성공 및 Excel 종료 기록. 실패 API를 정리 성공으로 인증하지 않음.
- rangevalidator/gui.py: 진단 보기 버튼, 자동 팝업, 마지막 실행 결과·진단 객체,
  읽기 쉬운 로그·스크롤·글자 확대·복사. 비교의 읽기 실패는 비모달 진단 표시.
- rangevalidator/readers.py: **진단 호출만 추가**. 기존 COM 호출·보안/Open 플래그,
  Find/Value2/Close(False)/Quit 순서와 데이터 반환은 유지.
- gui_entry.py + rangevalidator/diagnostic_probe.py: 명시적 `--diagnostic-test`
  EXE 검사. 엔진/Writer/보고서를 실행하지 않는 합성 Reader/진단창 검사.
- tests/test_diagnostics.py 및 진단 전용 실행/범위/배포 검사 도구 추가.
- 진단 안내/감사/작업 문서와 배포 해시 기록 갱신.

로그 코드를 제거한 Reader AST가 이전 버전과 같음을 확인했습니다.
분석 엔진, 파서, Common, workflow, Writer, 보고서, 보존 검사, Workbook 모델은
바이트 단위 해시가 동일합니다. 근거: DIAGNOSTIC_SCOPE_CHECK.json.

## 이번에 실행한 테스트

관련 unittest **{t['tests']}/{t['tests']} 통과**, 실패/오류/건너뜀 0,
{t['seconds']}초. 목록은 DIAGNOSTIC_TEST_RUN.json, 상세 로그는
DIAGNOSTIC_TEST_RUN.log에 있습니다.

- 단계 기록, 반복 Find/마지막 성공, Close/Quit 실패 HRESULT, 마스킹,
  native 성공/실패 및 fallback 기록, 로그 sink 오류와 Reader 호출/결과 동일성.
- GUI 버튼·팝업 자동 표시·실패 후 객체/결과 유지·확대/스크롤·복사·작업 중 접근.
- 기존 COM Reader 계약/예외 정리, 실제 합성 XLSX 읽기, 누락/손상 경로,
  MTA 거부, 기존 GUI COM 오류 표시 관련 테스트만 선택 실행.

**전체 회귀, 분석 엔진, Writer/SaveAs, 보고서 테스트는 실행하지 않았습니다.**
기존 TEST_RUN.json의 106개 결과는 이전 전체 기능 구현 단계 기록이며 이번
진단 수정의 전체 회귀 결과가 아닙니다.

## 최신 배포 EXE 검사

ZIP 추출 전 파일 해시 대조 후 Windows/System32만 PATH에 남기고 Python
환경변수를 제거했습니다. 기본 GUI 실행 및 진단 전용 EXE 검사가 통과했습니다.
pywin32는 추출 패키지 내부에서 로드됩니다.

실제 COM 합성 Open 실패: API `{probe['failure_api']}`, HRESULT
`{probe['hresult']}`, 마지막 성공 `{probe['last_success']}`.
진단창 자동 표시, 실패 로그/마지막 결과 유지, 민감 데이터 생략,
DispatchEx/Open/Find/Value2/Close/Quit 성공 로그, 원본 해시/mtime 보존을 확인했습니다.
EXE 검사에서 Writer/보고서는 실행하지 않았습니다.
근거: DIAGNOSTIC_DISTRIBUTION_VALIDATION.json.

EXE: `{d['exe']}`
SHA-256: `{d['exe_sha256']}`

ZIP: `{d['zip']}`
SHA-256: `{d['zip_sha256']}`

이전 배포물은 build/release-history에 보관했습니다. Python 설치 없이 실행하는
onedir 구조를 유지했고 원본 회사 파일/평문 임시 XLSX를 생성하거나 수정하지
않았습니다. 합성 테스트 입력만 새로 생성했습니다.

## 회사 PC에서 확인

1. ZIP 전체를 새 폴더에 풀고 EXE를 실행합니다. `_internal`을 함께 보관합니다.
2. 파일 등록 후 비교를 실행합니다. 읽기 실패 시 진단창이 자동으로 열립니다.
3. 언제든 오른쪽 아래 **진단 로그 보기**를 누를 수 있습니다.
4. 실행 ID, Reader, 실패 API, HRESULT/Win32, 마지막 성공, Open 진입,
   Close/Quit/Excel 종료를 확인합니다. 창 확대와 글자 크게로 사진을 촬영할 수 있습니다.
5. 필요하면 복사/선택 로그 저장을 사용합니다. 기본 로그는 경로/파일명/셀 값/
   원문 예외를 포함하지 않습니다. 파일 등록·비교 결과 화면의 경로는 별도입니다.

로그는 앱 실행 중의 제한된 메모리에 유지되며 재시작 자동 복원은 제공하지
않습니다. 실제 회사 DRM 허용/거부 정책 및 물리적 클린 PC는 미검증입니다.
'''
(root/'docs/DIAGNOSTIC_AUDIT.md').write_text(report,encoding='utf-8')
p=root/'docs/TASKS.md'
with p.open('a',encoding='utf-8') as f:f.write('''

## 진단 로그 한정 수정 (2026-11-03)

- [x] 기존 소스/EXE 진단 탭 및 누락 단계 확인, 기존 배포 해시/실행 증거 보관.
- [x] 진단 보기 버튼·별도 창·읽기 실패 자동 표시·실패 후 진단/결과 유지.
- [x] 단계별 성공/실패·HRESULT/Win32·마지막 성공·Close/Quit/종료 및 마스킹 보완.
- [x] Reader 비진단 AST 동일, 엔진/Writer/보고서/공통 모델 해시 불변.
- [x] 관련 27개 테스트만 통과. 전체/분석/Writer/보고서 회귀는 실행하지 않음.
- [x] onedir 재패키징, 새 ZIP 추출·Windows-only PATH 진단 전용 EXE 검사 통과.
- [ ] 실제 회사 DRM 읽기·정책 거부 및 별도 클린 PC 확인.

최신 진단 감사와 해시: DIAGNOSTIC_AUDIT.md. 이전 106개 전체 회귀는 과거 기록입니다.
''')
for rel in ('README.md','docs/TEST_RESULTS.md'):
    p=root/rel;s=p.read_text(encoding='utf-8-sig')
    banner=('> 최신 배포는 진단 로그 한정 수정본입니다. 이번에는 관련 **27개 테스트만** '
            '실행했으며 전체/Writer/보고서 회귀는 실행하지 않았습니다. '
            '[최신 진단 감사]('+('docs/' if rel=='README.md' else '')+'DIAGNOSTIC_AUDIT.md)를 확인하세요.\n\n')
    p.write_text(banner+s,encoding='utf-8')
print(json.dumps(dict(related_tests=t['tests'],full_suite_run=False,writer_tests_run=False,report_tests_run=False,
    exe_sha256=d['exe_sha256'],zip_sha256=d['zip_sha256'],packaged_source_matches_tested_source=True),indent=2))
