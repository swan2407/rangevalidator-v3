# FlexibleCatalog 배포 변경 사항

2026-11-09. 실행 파일은 `dist/RangeValidator-FlexibleCatalog/RangeValidator-FlexibleCatalog.exe`.
실행 파일과 같은 폴더의 `_internal`을 함께 보관한다.

## 입력을 읽는 방식

- `STRUCT`, `Struct`, `struct`를 같은 구조체 키워드로 인식한다. `#define STRUCT struct`는 필요하지 않다.
- `TYPEDEF`, `EXTERN`, `CONST`, `UNSIGNED INT` 등 선언 키워드의 대소문자 차이를 허용한다.
- 등록한 C/H 전체를 모아 타입과 매크로를 찾는다. 파일 간 `#include`가 없어도 없는 타입 정의를 다른 등록 파일에서 연결한다.
- 변수명, typedef 이름, 함수명, 멤버명은 정확히 일치하는 이름을 먼저 사용하고, 없으면 대소문자를 구분하지 않고 찾는다.
- 엑셀의 `CAN_CHG . READY`, `can_chg_p - > ready`, `can_chg [ 0 ] . ready`를 읽는다.
  연산자 주변의 공백·탭·줄바꿈을 허용하며 결과에는 원래 엑셀 표현식을 보존한다.

## 구조체 타입을 먼저 찾는 순서

실제 변수 선언이 있으면 선언된 타입을 우선 사용한다. 변수명에서 추정한 타입으로 덮어쓰지 않는다.
변수 선언이 없으면 다음 이름 규칙으로 타입을 찾고, 그 타입 안에서만 멤버를 찾는다.

| 엑셀 표현식 | 우선 찾는 타입 | 찾는 멤버 |
| --- | --- | --- |
| `can_chg.ready` | `can_chg_t` | `ready` |
| `can_chg_p->ready` | `can_chg_t` | `ready` |
| `can_chg[0].ready` | `can_chg_t` | `ready` |
| `device.status` | `device_t` | `status` |
| `sensor_p->value` | `sensor_t` | `value` |

`_t` 타입이 없으면 변수명과 같은 타입·구조체 태그도 찾는다. typedef 별칭도 연결한다.
타입이 없거나 그 타입 안에 멤버가 없으면 REVIEW로 표시하며, 다른 타입의 같은 이름 멤버로 연결하지 않는다.
타입 이름으로 추정한 비교에는 `변수 선언 연결 없이 변수명으로 타입 추정`과 선택한 타입·근거 파일을 표시한다.

동일한 타입 이름의 정의가 여러 파일에 있으면 해당 멤버의 타입과 범위가 일치하는 경우 비교할 수 있다.
서로 다른 범위, 일반 멤버와 비트필드 혼재, 일부 후보의 멤버 누락은 REVIEW로 표시한다.
실제 전처리 조건, 알 수 없는 비트 폭·부호 등으로 범위를 계산할 수 없는 경우에도 REVIEW를 유지한다.

일반 멤버는 `일반 멤버 (SKIP)`로 표시하고 집계·필터·보고서에 포함한다.
비트필드의 기존 범위는 MATCH/MISMATCH로 비교하고, 공란 범위는 CREATE로 표시한다.
비교만 실행할 때에는 원본 Excel을 변경하지 않는다.

## 검증 기록

- 회귀 테스트 `build/flexible-catalog-tests.log`
- 실제 Excel COM 및 OpenPyXL 비교 `build/flexible-catalog-verification/result.json`
- 배포 빌드 `build/flexible-catalog-build.log`
- 배포 EXE 실행 확인 `build/flexible-catalog-smoke.json`

회사 원본 파일을 수정하지 않고 합성 C/H와 A/B 양식 Excel로 검증한다.

검증 결과는 전체 264개 중 252개 통과, 선택 실행 조건에 따른 12개 건너뜀이다.
실제 Excel COM/OpenPyXL에서 include 없는 선언 및 변수 선언 없는 타입 추정을 모두 확인했다.
양식 A/B의 결과 셀 주소와 원본 해시가 유지됐고, 배포 EXE의 GUI 및 COM 모듈 로드도 통과했다.
