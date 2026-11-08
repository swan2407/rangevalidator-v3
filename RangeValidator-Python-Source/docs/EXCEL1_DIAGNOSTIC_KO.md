# Excel-1: A/B 혼합 상세설계서 추출 진단

검증일: 2026-11-08. 실제 작업 경로는 `C:\Users\ido11\range\rangevalidator-v3\RangeValidator-Python-Source`이다. 요청서의 `C:\range\...` 및 `C:\Dev\RangeValidator-Python\.build-python\python.exe`는 이 PC에 없다. Python 3.11.9와 임시 설치한 openpyxl 3.1.5, Pillow 13.0.0, pywin32 311을 사용했다. 제공된 `A형.xlsx`, `B형.xlsx`는 읽기만 했고 수정하지 않았다.

## 실제 호출 순서와 필터링

1. `workflow.run_comparison()`이 `com_values.excel_session()` 안에서 C/H 카탈로그를 만든다.
2. 각 상세설계서를 `read_auto()`로 연다. 일반 XLSX는 `read_openpyxl()`이 수식/캐시 값을 위해 통합 문서를 두 번 로드하고 **모든 시트**를 순회한다. 허용된 열기 실패에서만 `read_com()`으로 넘어간다.
3. COM 경로는 `read_workbook_values()` → `_open()` → `book.Worksheets.Count` → 각 `Worksheets.Item(i).Name` → `_copy_sheet()` → `_last_cell()`/`_read_blocks()`/`_merge_map()` 순이다. 함수 시트 판정 전에 모든 시트의 본문 Value2/Formula를 읽는다.
4. 모델이 완성된 후 `workflow._run_comparison()`이 `engine.compare_book()`을 호출한다. 여기서 `is_function_sheet()`가 `숫자.이름` 형태의 시트만 분석 대상으로 등록한다. 나머지 시트는 이 단계에서 제외된다.
5. `compare_book()`은 대상 시트의 각 행에서 헤더 조합을 찾고, 발견한 헤더의 Name 및 Range/Scope 열을 같은 행에서 읽는다. C/H 비트필드 카탈로그에 해당 멤버명이 없으면 결과 행을 만들기 전에 건너뛴다. 이후 변수 선언 연결과 범위 비교가 진행된다.

따라서 `commonparameter`가 COM 읽기 로그에 나타나는 것은 단순 이름 조회만을 뜻하지 않는다. 현재 COM Reader는 그 시트의 이름을 조회하고 본문도 읽는다. 합성 파일의 `commonparameter`는 진단에서 `SHEET_EXCLUDED`이지만 `body_read_success`에 포함된다. `compare_book()`은 이 시트를 함수 시트로 등록하거나 변수 비교하지 않는다. 실제 회사 파일의 특정 로그 한 줄만으로 읽기 완료 여부를 단정할 수는 없으므로, 개별 시트의 성공/실패 단계와 함께 봐야 한다.

**양식 명칭 주의:** 요청서의 A형은 `Input/Output, Property, Name, Data Type, Scope`이고 현재 `engine.py` 내부에서는 `B`로 표기한다. 요청서의 B형은 구역별 `Name, Type, Range`이고 엔진 내부에서는 `A`로 표기한다. 새 진단 JSON은 요청서의 A/B 명칭을 사용한다.

## 샘플 관찰

| 읽기 전용 샘플 | 전체 시트 | 함수 시트 | 진단 양식 | 표 | Name/Range 또는 Scope 헤더 | 관찰한 변수명 셀 |
|---|---:|---|---|---:|---|---:|
| A형.xlsx | 7 | `1. charger_init` | A | 1 | D24 / G24 | 10 |
| B형.xlsx | 7 | `1.calc_chg_acc` | B | 6 | B20/D20, B24/D24, B28/D28, B38/D38, B45/D45, B51/D51 | 15 |

샘플에는 대응 C/H 입력이 없으므로 이 수치는 Excel 표 관찰 결과이다. C/H 변수 연결이나 비트필드 비교 성공 수치로 해석하지 않는다.

## 합성 정답 대조

`tools/excel1_synthetic.py`는 샘플의 헤더 배치를 참고해 별도 XLSX 12개와 독립 정답 `ground_truth.json`을 생성한다. 정답에는 파일·시트·양식·구역·변수명·Name 셀·Range/Scope 셀·기존 범위값을 기록한다. 진단기는 읽힌 Workbook과 변경 없는 엔진 결과를 관찰하며, `assess()`는 정답과 이름 및 양쪽 셀 주소를 함께 대조한다.

| 사례 | 기대 변수 | openpyxl 추출 | 실제 Excel COM 추출 | 결과 |
|---|---:|---:|---:|---|
| A형 1시트 | 10 | 10 | 10 | 일치 |
| B형 구역별 3표 | 15 | 15 | 15 | 일치 |
| A/B/A/B 혼합 | 24 | 24 | 24 | 일치 |
| B/A/B/A 혼합 | 20 | 20 | 20 | 일치 |
| 헤더 행·열 이동 | 20 | 20 | 20 | 일치 |
| 공란·병합 헤더·반복 표·동일 변수명 | 17 | 17 | 17 | 일치 |
| Range 헤더 누락 | 5 | 0 | 0 | 5개 누락 |
| 수식이 빈 문자열을 반환하는 범위 | 6 | 6 | 6 | 1개 REVIEW, 연결은 유지 |
| 헤더 내부 공백·줄바꿈 | 12 | 0 | 0 | 12개 누락 |
| 영문 Name/Type + 한글 범위 헤더 | 4 | 0 | 0 | 4개 누락 |
| A/B 혼합 함수 시트 120개 | 7260 | 7260 | 7260 | 일치 |
| A/B 각 200개 변수 | 400 | 400 | 400 | 일치 |
| **합계** | **7793** | **7772** | **7772** | **정확 연결 7772, 누락 21** |

중복·과잉 추출 0개, 잘못된 Range/Scope 셀 연결 0개, 양식 판별 오류 2시트이다. 판별 오류는 Range 헤더 누락과 한글 범위 헤더 사례이며, 헤더 내부 공백 사례는 진단기에서는 양식을 찾았지만 현재 엔진은 찾지 못했다. 잘못된 셀 연결을 의도적으로 주입한 단위 테스트에서는 정답 대조기가 1건을 검출했다.

### 누락 변수와 원인

- `missing_header.xlsx`, `1. missing_range`: `obj.v001`~`obj.v005`, Name B21:B25, 정답 범위 D21:D25. Range 헤더가 없어 실제 결과 0개. 진단은 `RANGE_CELL_UNRESOLVED`로 기록한다.
- `labels_spaced.xlsx`, `1. spaced_a`: Name D25:D30, 정답 Scope G25:G30의 6개. `2.spaced_b`: Name C29:C34, 정답 Range E29:E34의 6개. 헤더 텍스트 안의 공백·줄바꿈 때문에 현재 엔진의 정확한 헤더 조합 검사가 실패한다.
- `korean_mixed_header.xlsx`, `1. korean_range`: Name B26:B29, 정답 범위 D26:D29의 4개. 현재 엔진은 한글 `범위`를 `Range`로 인식하지 않는다.

위 21개에서 실제로 **잘못된 다른 범위 셀에 연결된 사례는 없다**. 9개는 헤더 열을 확정하지 못했고, 12개는 진단기가 양쪽 셀을 찾았지만 비교 엔진이 표 헤더를 등록하지 못했다. 완전한 개별 항목은 산출물 `assessment.json`의 `missing_rows`에 있다. 변수는 Excel에서 추출됐지만 C/H 선언 연결이 실패한 경우에는 진단의 `extracted=true`, `comparison_status=REVIEW`, `comparison_reason`으로 별도 확인한다.

## 대량 시트와 성능

Windows Excel COM에서 합성 파일을 실제로 읽었다. `large_120.xlsx`는 전체 122시트 중 함수 120, 제외 2, 본문 읽기 성공 122, 실패 0, 시트명 판별 실패 0이었다. 함수 시트별 변수 수는 각각 1개부터 120개까지 모두 다르며 7260/7260개를 정확히 연결했다. 15개 시트에는 병합 헤더가 있다. `large_200.xlsx`는 함수 2시트의 A/B 각 200개, 합계 400/400개를 정확히 연결했다.

| COM 측정 | 120 함수 시트 | A/B 각 200개 |
|---|---:|---:|
| 전체 진단 실행 | 13.893초 | 3.382초 |
| 통합 문서 열기 | 1.024초 | 0.951초 |
| 시트 수 조회 + 시트명 조회 | 0.324초 | 0.013초 |
| 본문 읽기 | 9.109초 | 0.158초 |
| 별도 진단 변수 관찰 스캔 | 0.067초 | 0.004초 |
| 현재 비교 엔진(헤더·변수 추출 포함) | 0.265초 | 0.014초 |
| 읽기 오류 / 누락 | 0 / 0 | 0 / 0 |

각 파일마다 Excel COM 인스턴스 시작·정리와 C/H 카탈로그 생성도 전체 실행 시간에 포함된다. 단계별 시간은 해당 단계만 측정했으므로 합계가 전체와 정확히 같지 않다. 별도 진단 스캔은 Workbook에서 변수 셀을 관찰한 시간이다. **현재 생산 엔진의 `compare_book()`은 표 인식·변수 추출·비교를 한 함수에서 수행하므로 생산 변수 추출 시간만 독립적으로 측정할 수 없다.** 그 값을 추정해 표에 넣지 않았다. 메모리 사용량도 측정하지 않았다. 같은 대량 파일의 openpyxl 진단은 각각 약 1.638초와 0.204초였다. 숫자는 이 합성 파일과 이 PC의 관찰값이며 회사 DRM 파일의 성능 추정치는 아니다.

## 기존 기능과 테스트

Reader의 시트 선택·셀 값·수식 처리·병합 처리·판정·Writer·보고서 로직은 변경하지 않았다. `Workbook`에 비교에서 제외되는 관찰 필드(원본 시트 수·읽기 시간)를 추가하고, COM/openpyxl 및 workflow에 시간 측정만 넣었다. 별도 `extraction_diagnostics.py`와 진단 실행기, 합성 생성기, 4개 테스트를 추가했다. 진단 보고서는 `Session`의 저장 후보와 결과 서명에 영향을 주지 않는다.

- 신규 진단 테스트 4개 통과. 이 중 하나는 진단 전후 기존 결과와 세션 서명이 같음을 확인한다.
- 기존 Reader·모의 COM 관련 테스트를 별도 실행: `test_readers.py` 22개 중 17개 통과·5개 옵션 건너뜀, `test_com_reference.py` 17개 중 15개 통과·2개 옵션 건너뜀.
- GUI 포함 전체 회귀 200개 실행: 184개 통과, 16개 기존 옵션 테스트 건너뜀, 실패·오류 0개.
- 실제 Excel COM 합성 진단 12개 파일 완료. 단위·모의 COM 결과와 별도 기록했다. 이번 단계에서 실제 회사 DRM 파일이나 사용자 원본 A/B 샘플을 COM으로 열거나 저장하지 않았다.

## 다음 단계의 수정 우선순위

1. **Excel-2 — 헤더 텍스트 변형:** `engine.compare_book()`의 A/B 헤더 인식에 국소적인 공백·줄바꿈 정규화와 승인된 한글 헤더 별칭을 넣을 지점이다. `common.norm()`을 전역 변경하면 다른 판정에도 영향을 주므로 별도 검토가 필요하다. 한글 `범위`를 모든 Scope처럼 취급하지 말고 표 구역 및 전체 헤더 조합으로 확인해야 한다.
2. **Excel-3 — 표 경계와 불확실성:** `engine.compare_book()`의 `headers` 수집, `section` 갱신, Name/Range 같은 행 결합을 중심으로 다중 표·병합·헤더 누락 사례를 확장한다. 헤더가 없는 표의 범위 열은 추정 기입하지 않고 진단 또는 REVIEW 근거로 남겨야 한다. 현재는 한 시트에 내부 A/B 두 양식이 함께 있으면 `Mixed A/B tables within sheet`로 시트를 REVIEW 처리한다.
3. **Excel-4 — 읽기 대상과 성능:** `com_values.read_workbook_values()`의 전체 시트 순회와 `_copy_sheet()`의 본문 읽기, `readers.read_openpyxl()`의 전체 시트 로드가 검토 지점이다. 함수 시트 판정은 현재 `engine.is_function_sheet()`에서 읽기 후에 실행된다. 함수 외 시트의 읽기를 생략할 경우 부분 읽기 오류·시트 수·진단 의미를 보존하는 설계와 회사 파일 검증이 필요하다. 합성 120시트 COM 측정에서는 본문 읽기 9.109초가 가장 큰 구간이다.

진단 JSON과 XLSX는 `../excel1-diagnostic/`(일반 XLSX), `../excel1-com-diagnostic/`(실제 COM)에 있다. 실회사 파일 원인이나 DRM 동작은 이 결과만으로 확정하지 않는다.
