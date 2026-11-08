# Excel-2 — 함수 시트 우선 COM 읽기

2026-11-08, 실제 작업 경로: `C:\Users\ido11\range\rangevalidator-v3\RangeValidator-Python-Source`.

## 구현

기존 비교용 COM 경로는 시트 이름과 본문을 차례로 읽은 후 `engine.compare_book()`에서 함수 시트를 선별했다. 이제 `read_workbook_values()`가 전체 시트의 `Worksheets.Item(i).Name`을 먼저 수집하고, `engine.is_function_sheet()`를 그대로 호출해 대상 목록을 만든 다음 대상의 본문만 읽는다. 공통 판정은 기존 `숫자.함수명` 정규식이며 새 규칙을 만들지 않았다.

`workflow._run_comparison()`은 직접 COM 읽기와 `read_auto()`의 COM fallback에 `function_sheets_only=True`를 전달한다. 일반 `read_com()` API의 기본값은 기존 전체 읽기로 유지하여 독립적인 Reader 사용 및 저장 검증과 호환된다. openpyxl의 읽기 방식은 이번 단계에서 바꾸지 않았다. 따라서 **함수 시트 본문만 읽는 정책은 비교 실행의 COM 경로에 적용된다.**

제외 시트에는 `_copy_sheet()`를 호출하지 않는다. Value2, Formula, UsedRange, Find, 병합 속성 등의 본문 접근도 수행하지 않는다. Common Parameter는 이름 조회와 `SHEET_EXCLUDED` 이벤트만 기록된다. 함수 시트 안의 `_last_cell()`, `_read_blocks()`, `_merge_map()`은 그대로다.

## 변경 파일과 함수

| 파일 | 변경 |
|---|---|
| `rangevalidator/com_values.py` | `read_workbook_values()` 이름 목록 선조회, 공통 함수로 선별, 대상만 본문 읽기, 이벤트 |
| `rangevalidator/readers.py` | `read_com()`, `read_auto()` 선택 옵션 전달 |
| `rangevalidator/workflow.py` | 비교용 COM 및 fallback에 선택 옵션 활성화 |
| `rangevalidator/models.py` | 비교 동등성에 영향을 주지 않는 `sheet_observations` 추가 |
| `rangevalidator/diagnostics.py` | 이름 발견·포함·제외·본문 시작·완료·실패 이벤트 허용 |
| `rangevalidator/extraction_diagnostics.py` | 본문을 읽지 않은 제외 시트 관찰 정보와 실제 읽기 수 집계 |
| `tests/test_excel2_selection.py` | 선조회 순서, 126시트, 제외 본문 접근 차단, 이름/본문 오류, 전부 제외·전부 포함 테스트 |
| `tools/diagnose_excel2.py` | 합성 12파일 실제 COM 전후 대조와 원본 해시 확인 |

## 오류 처리

- 제외는 `READ_ERROR`나 `partial`의 원인이 아니며 기존 보고서 Errors에도 추가하지 않는다.
- 함수 시트 본문 실패는 시트명을 가진 기존 `read_errors`로 남고 다음 시트를 계속 읽는다. 기존 workflow의 시트별 후보 차단 규칙을 유지한다.
- 시트명 조회 실패는 빈 시트명 오류와 `partial=True`로 기록한다. 범위를 확정하지 못하므로 기존 workflow가 자동 수정 후보를 보수적으로 차단한다.
- 부분 본문 읽기와 기존 예산 오류 처리는 그대로 유지한다.

## 실제 Excel COM 결과

회사 DRM 문서는 사용하지 않았다. Excel-1 산출물은 덮어쓰지 않았으며 `../excel2-diagnostic/`에 새 합성 파일과 결과를 생성했다. 대량 파일에는 비함수 시트 6개, A형 60개, B형 60개가 있다. 함수별 변수 수는 1~120개로 서로 다르다.

| 항목 | 결과 |
|---|---:|
| 전체 시트 이름 | 126 |
| 함수 시트 본문 읽기 | 120 |
| 제외 시트 | 6 |
| 제외 시트 본문 읽기 진입 | 0 |
| 대량 파일 변수 정확 추출 | 7,260 / 7,260 |
| 합성 전체 정확 추출 | 7,772 / 7,793 |
| 기존 누락 | 21 |
| 잘못된 Range/Scope 셀 연결 | 0 |
| 12파일 전후 결과 행 전체 비교 | 모두 동일 |
| 12파일 원본 바이트 SHA-256 | 모두 동일 |

모의 COM은 제외 시트에서 Name 이외 속성에 접근하면 예외를 내도록 구성해 본문 접근 차단을 검증했다. 실제 COM에서는 모든 본문 획득을 수행하는 `_copy_sheet()` 진입 시트명을 계측했고 제외 시트 진입은 0회였다. 개별 COM 메서드마다 별도 계수기를 설치한 것은 아니다. 양쪽 검증과 분기 코드로 제외 시트 본문 COM 호출 0회를 확인했다.

전후 대조는 같은 합성 파일을 전체 읽기 옵션과 함수만 읽기 옵션으로 각각 실행했다. 모든 Result 데이터 필드를 대조하므로 파일·시트별 MATCH/MISMATCH/REVIEW/CREATE 집계도 동일하다. A/B 판별은 계속 각 함수 시트에서 기존 엔진이 수행한다.

## 시간

- 보존한 Excel-1 기준: 120 함수 시트 전체 13.893초, 본문 9.109초(당시 전체 시트 122개).
- 이번 동일 126시트 파일 전체 읽기 옵션: 13.861초.
- 이번 함수만 읽기 옵션: 13.799초, 본문 9.020초.

시간 차이는 작다. 이름 126개 조회, 함수 본문 120개 읽기, 제외 본문 0회 및 결과 동일성을 완료 근거로 삼았다. 상세 수치는 `../excel2-diagnostic/validation.json`에 있다.

## 검증과 남은 단계

현재 실행 가능한 Windows Python 3.11.9와 이전 단계의 임시 의존성을 사용했다. 요청서의 CPython 3.13 경로는 존재하지 않는다.

- 신규 모의 COM 테스트 5개 통과.
- GUI 포함 전체 회귀 205개: 189개 통과, 기존 옵션 16개 건너뜀, 실패·오류 0개.
- 실제 Excel COM: 합성 12개 파일을 전후 각 1회 실행, 결과·원본 동일 및 제외 시트 미읽기 확인.
- 초기 실제 검증에서 테스트의 Mock이 COM 프록시를 보관해 종료 오류가 났다. 실제 계측 래퍼를 참조를 저장하지 않는 함수로 바꾼 뒤 검증을 완료했다.

비트필드 파서·타입 추적·A/B 헤더 및 표 추출·판정·Writer·SaveAs·파일명·GUI·5시트 보고서 구현은 변경하지 않았다. 원본 해시 확인은 합성 검증 스크립트에서만 하며 저장 차단 정책으로 재도입하지 않았다.

Excel-3에서는 A형 헤더 내부 공백·줄바꿈 정규화와 구역/헤더 조합을 확인할 필요가 있다. 한글 범위 별칭, 누락 헤더 및 B형 반복 표 경계는 이후 범위에서 검토해야 한다. 기존 누락 21개는 이번 단계에서 유지됐으며 Excel-3·4 파서 개선은 수행하지 않았다.
