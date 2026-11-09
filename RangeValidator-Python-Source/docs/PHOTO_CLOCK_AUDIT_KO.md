# 사진 코드 원인 분석 및 ConditionalCompare 변경 사항

2026-11-09. 배포 경로: `dist/RangeValidator-ConditionalCompare/RangeValidator-ConditionalCompare.exe`.
실행 파일과 `_internal` 폴더를 함께 보관한다. 기존 COMRetry3의 읽기 3회 시도와 FlexibleCatalog의 타입 추정 기능을 포함한다.

## 사진에서 실제로 확인한 내용

입력은 `6_20261109_141012.jpg_261109.zip`의 두 사진이다.

- `20261109_141012.jpg`: 결과 사유에 `비트필드 'can_hu_clock_01_t.hu_display...' 확인` 다음으로 `변수 선언 연결이 불확실함 (CONDITIONAL)`이 표시된다.
- `20261109_141006.jpg`: `canfd_comm.h`의 2334~2352행에 `typedef struct { ... } can_hu_clock_01_t;`가 있고, 아래 네 멤버의 이름과 비트폭이 보인다.

즉, 이 네 행은 엑셀 변수명이나 구조체 멤버를 처음부터 못 읽은 문제가 아니다. 이미 찾은 비트필드를 조건부 선언이라는 이유로 비교 대상에서 제외한 문제다. 이전 테스트도 이 일괄 REVIEW 동작을 기대해서 문제를 잡지 못했다.

| 엑셀 변수의 멤버 | 코드 행 | 사진의 선언 | unsigned 해석 시 범위 |
| --- | --- | --- | --- |
| `hu_displayday` | 2349 | `uint hu_displayday:8;` | `0~255` |
| `hu_displaymonth` | 2346 | `uint hu_displaymonth:4;` | `0~15` |
| `hu_displayyear1` | 2347 | `uint hu_displayyear1:2;` | `0~3` |
| `hu_displayyear2` | 2348 | `uint hu_displayyear2:8;` | `0~255` |

사진에는 `uint`의 typedef, 바깥쪽 `#if`, 변수 선언 전체가 보이지 않는다. 따라서 실제 프로젝트의 어느 조건문이 CONDITIONAL을 만들었는지는 사진만으로 확정하지 않았다. 엑셀 기존 범위도 `0~4294...`로 잘려 있어 정확한 전체 값은 확정하지 않았다. 사진에 소스가 없는 나머지 결과 행은 개별 검증하지 않았다.

## 잘못 중단하던 지점

1. `engine._catalog_match`: 함수, 변수 또는 구조체가 CONDITIONAL이면 범위 비교 전에 REVIEW로 반환했다. 타입·멤버를 찾았더라도 중단됐다.
2. `bitfield_catalog._field`: 구조체 또는 기본 타입에 조건이 붙으면, 숫자로 적힌 `:8`, `:4`, `:2`도 범위를 계산하지 않았다.
3. 숫자 폭 대신 매크로를 쓴 경우에도 정의가 하나뿐인데 조건부라는 이유만으로 계산하지 않았다.
4. `uint` 정의 자체를 등록하지 않은 경우에는 부호를 결정할 수 없었다. 사진 코드만 독립적으로 넣었을 때 추가로 확인한 문제다. 실제 사용자 프로젝트에서 typedef가 누락됐다는 의미는 아니다.

## 변경한 비교 규칙

- 소스에 선언된 비트폭을 비교한다. `#if`가 있다는 사실만으로 막지 않고, 계산 가능한 타입·부호·폭이 일치하면 비교한다. 결과 사유에 `조건부 선언의 비트폭 기준 비교 (CONDITIONAL)`을 남긴다. 실제 빌드에서 어느 분기가 활성화되는지 판정한 것은 아니다.
- 같은 타입의 후보가 여러 개면 해당 멤버를 모두 확인한다. 후보의 범위가 다르면 REVIEW에 양쪽 범위와 위치를 표시한다. 다른 멤버까지 함께 REVIEW로 만들지 않는다.
- 비트폭 매크로의 수집된 정의가 같은 본문이면 계산할 수 있다. 본문이 다르거나 값을 계산할 수 없으면 REVIEW를 유지한다.
- `uint`는 실제 typedef/매크로가 있으면 그 정의를 우선한다. 등록된 파일 전체에도 정의가 없을 때만 `unsigned int` 관례를 적용하고, 사유에 `uint 정의 미수집: 관용 타입 unsigned int 적용`을 표시한다. 명시적인 signed 정의, 충돌하거나 미해결인 정의를 덮어쓰지 않는다.
- 실제 변수 선언을 우선한다. 없으면 `can_hu_clock_01` → `can_hu_clock_01_t` 순으로 타입을 추정하고 해당 타입 안에서만 멤버를 찾는다. 관련 없는 구조체의 동명 멤버로 대체하지 않는다.
- 일반 멤버는 계속 `일반 멤버 (SKIP)`로 표시한다. 실제 타입/범위 충돌, 미해결 부호·폭, 접근 형태 불일치 등은 REVIEW로 남는다.

## 사진 코드로 재현한 전후 결과

`tools/photo_clock_audit.py`에 사진의 구조체 전체를 옮겼으며, 네 멤버의 원래 소스 행 번호를 유지했다. 실제 XLSX 파일의 H31~H34를 읽어 비교한다. 사진에 없는 조건문과 typedef는 아래 시나리오별로 별도 추가했다. 재현 엑셀의 기존 범위 `0~4294967295`와 헤더는 검증용 입력이다.

| 재현 입력 | 수정 전 | 수정 후 |
| --- | --- | --- |
| 사진 구조체만, uint 정의/변수 선언 없음 | 4행 REVIEW | 4행 MISMATCH, 관용 uint 사용 표시 |
| unsigned uint 정의만 추가, 변수 선언 없음 | 4행 MISMATCH | 동일한 4개 범위 유지 |
| 변수 선언에 조건 추가 | 4행 REVIEW | 4행 MISMATCH |
| 구조체 정의에 조건 추가 | 4행 REVIEW | 4행 MISMATCH |
| include에 조건 추가 | 4행 REVIEW | 4행 MISMATCH |
| 함수 정의에 조건 추가 | 4행 REVIEW | 4행 MISMATCH |
| uint typedef에 조건 추가 | 4행 REVIEW | 4행 MISMATCH |
| 다른 분기에서 day만 4비트로 변경 | 4행 REVIEW | day만 REVIEW(0~255 / 0~15), 나머지 3행 MISMATCH |

MISMATCH는 연결 실패가 아니라, 위 표의 계산 범위와 재현 엑셀에 넣은 기존 범위가 다르다는 정상 비교 결과다.

OpenPyXL과 실제 Excel COM 양쪽에서 8가지 입력의 32개 행을 확인했다. 두 읽기 경로의 결과가 위 표와 일치했고, 재현 엑셀의 비교 전후 SHA-256이 같았다. 새 배포 EXE도 실행하여 GUI 4개 탭과 번들 내부의 Excel COM 모듈 로드를 확인했다.

검증 자료:

- `build/photo-clock-audit/before-openpyxl.json`: 수정 전 수집된 타입·멤버·변수 연결과 비교 결과
- `build/photo-clock-audit/after-openpyxl.json`: 수정 후 동일 입력 결과
- `build/photo-clock-audit/after-com.json`: 실제 Excel COM 읽기 결과 및 원본 해시
- `tests/test_photo_clock.py`: 사진의 원본 행·폭·엑셀 셀을 대조하고, signed uint/충돌/포인터/매크로 분기를 추가 확인
- `build/photo-clock-tests.log`: 전체 277개 중 265개 통과, 선택 실행 조건에 따른 12개 건너뜀
- `build/conditional-compare-build.log`, `build/conditional-compare-smoke.json`: 배포 빌드 및 실행 확인

이 기록은 사진에 보이는 네 멤버와 추가 조건 시나리오의 검증이다. 원본 프로젝트 전체 C/H와 실제 상세설계서를 실행한 결과로 확대 해석하지 않는다.
