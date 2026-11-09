# StructMembers 배포 변경 사항

2026-11-09. 실행 파일은 `dist/RangeValidator-StructMembers/RangeValidator-StructMembers.exe`.
이전 DeclarationReview, DigitSheets 배포 폴더는 유지한다. 실행 파일과 `_internal` 폴더를 함께 보관한다.

## 첨부 사진에서 확인한 선언

사진에는 `typedef STRUCT { ... } can_chg_t;` 뒤에
`extern can_chg_t can_chg;`와 `extern const can_chg_t * const can_chg_p;`가 있다.
구조체 안에는 `uint obc_ready:1;` 비트필드와 `u2 qcs_maxcurr;` 일반 멤버가 함께 있다.
이 선언 순서 및 `const` 포인터 형태는 지원 대상이다.
사진만으로 `STRUCT`, `uint`의 실제 정의나 파일 전체의 전처리 조건은 확인할 수 없다.

## 수정한 동작

- 등록된 파일의 따옴표 include 경로에서 `#define STRUCT struct`와 기본 정수 타입 매크로를 선언 순서에 맞게 읽는다.
  해당 정의가 조건부이거나 include 파일을 등록하지 않았다면 임의로 확정하지 않는다.
- 파일 전체를 감싸는 `#if !defined(NAME)` 및 `#if !defined NAME`, `#define NAME 1` 형태의 include guard를 인식한다.
  앞에 `#pragma once`가 오는 경우도 지원한다. 실제 기능 선택 조건은 그대로 조건부로 남긴다.
- 조건부 타입·변수 연결에서 구조체 후보가 하나인 경우 멤버 종류의 근거를 유지한다.
  함수 또는 변수의 조건부 상태가 비트필드 범위 확정을 허용하지는 않는다.
- 확인된 일반 멤버를 결과에서 숨기지 않고 `일반 멤버 (SKIP)`로 표시한다.
  일반 멤버 필터, 집계와 보고서 표시를 제공하며 범위 자동 수정 대상에는 넣지 않는다.
- 조건부 선언에서 일반 멤버 종류만 확인한 경우 사유에 `빌드 조건 미확정`을 표시한다.
  비트필드임을 확인했으나 조건부로 남는 경우 REVIEW 사유에 비트필드임을 표시한다.
- 조건에 따라 타입이나 같은 이름 멤버의 종류가 달라지는 경우에는 REVIEW를 유지한다.

## 검증 및 한계

사진과 같은 선언 형태를 사용하는 합성 C/H 및 Excel 자료로 검증한다.
실제 회사 소스와 전체 상세설계서를 재분석한 결과는 아니다.
Clang을 비교 엔진에 새로 연결하거나 프로젝트의 빌드 매크로를 자동 추정하지 않는다.
사용할 C/H에는 구조체, 변수 선언뿐 아니라 `STRUCT`, `uint` 등의 정의가 들어 있는 파일도 등록해야 한다.
사용자 프로젝트에 남는 CONDITIONAL은 해당 선언과 함수 원형을 둘러싼 조건 및 빌드 설정을 확인해야 한다.

회귀 테스트 기록은 `build/struct-members-tests.log`, 합성 Excel COM/OpenPyXL 비교 기록은
`build/struct-members-verification/result.json`, 빌드 기록은 `build/struct-members-build.log`에 남긴다.

검증 결과는 전체 246개 중 234개 통과, 선택 실행 조건에 따른 12개 건너뜀이다.
별도로 실제 Excel COM/OpenPyXL의 A/B 표 읽기와 조건부 판정, 원본 해시 보존 검증을 통과했다.
배포 EXE의 GUI 및 Python COM 모듈 로드도 `build/struct-members-smoke.json`에서 확인했다.
