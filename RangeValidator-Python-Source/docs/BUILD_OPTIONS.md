# DirectCompare fix8 빌드옵션

## 사용 순서

1. 기존처럼 C/H 파일과 상세설계서 XLSX를 등록합니다.
2. 파일 등록 화면 아래에서 **빌드옵션 XML 선택**을 눌러 Eclipse CDT의 설정 내보내기 XML을 선택합니다.
3. 표시된 XML 경로와 매크로 개수를 확인한 뒤 **비교 실행**을 누릅니다.
4. 결과의 사유에서 적용 XML, 해시, C 문맥과 원본 선언 위치를 확인합니다. 보고서 Summary에도 설정 정보가 기록됩니다.

**적용 해제**를 누르면 기존 분석 방식으로 돌아갑니다. 옵션을 변경하거나 해제하면 이전 비교 결과가 무효화되므로 다시 비교해야 합니다. XML은 비교 시작 시 다시 읽습니다. 비교 결과에는 그때 읽은 값과 해시가 고정됩니다.

## 해석 범위

- Eclipse `cdtprojectproperties`의 `Macros` 및 `IncludePaths`에서 `C Source File` 항목을 읽습니다. Assembly/C++ 설정은 섞지 않습니다.
- `#if`, `#ifdef`, `#ifndef`, `#elif`, `#else`, 중첩 조건, 숫자·논리·비트 연산, 함수형 매크로, 소스의 `#define`/`#undef`, include 순서를 처리합니다.
- C 파일마다 독립적인 매크로 상태로 전처리합니다. 같은 헤더도 C 파일의 설정에 따라 다른 정의가 선택될 수 있습니다.
- include 경로는 로컬 경로와 XML 경로를 사용합니다. 찾지 못하면 등록한 C/H 중 요청한 경로 접미사와 유일하게 일치하는 헤더를 찾습니다. 중복 후보는 임의로 선택하지 않습니다.
- XML 상대 include 경로는 XML 파일이 있는 폴더 기준입니다. Eclipse workspace 경로나 `${...}` 변수는 실제 경로로 바꿔야 합니다.
- 빈 매크로 값은 **정의됨 + 빈 치환문**입니다. `defined(NAME)`/`#ifdef NAME`에는 사용할 수 있지만 임의로 숫자 1을 부여하지 않습니다. `#if NAME`처럼 숫자가 필요한 경우 XML에 실제 값을 적어야 합니다.
- 설정과 헤더에서 정의되지 않은 일반 매크로는 C 전처리 규칙에 따라 미정의로 취급합니다. 선택한 XML은 해당 분석에 사용할 프로젝트 매크로 집합이어야 합니다.
- 대상 컴파일러 전용 매크로가 조건에 필요한데 값이 없으면 REVIEW로 처리합니다. 호스트 Windows/Python의 매크로를 대상 임베디드 컴파일러 설정으로 대입하지 않습니다.

## REVIEW 처리

헤더 누락, 모호한 include 경로, 잘못된 조건식, 필요한 컴파일러 매크로 누락, 활성 `#error` 또는 지원하지 않는 전처리 지시문이 있으면 해당 C/H 문맥의 분석을 확정하지 않습니다. 관련 행은 `BUILD_OPTIONS_UNRESOLVED` 사유와 함께 REVIEW가 되며 자동 수정 후보에서 제외됩니다. 파일 문맥을 결정할 수 없고 실패한 C 파일이 있으면 해당 행도 REVIEW로 남깁니다. 독립적으로 정상 분석된 C 파일의 행은 계속 비교합니다.

`#pragma once`와 include guard를 처리합니다. 메모리 배치만 바꾸는 `#pragma pack`/`section`은 값 범위 계산에 사용하지 않습니다. 그 밖의 pragma나 컴파일러 확장은 별도 지원이 필요할 수 있습니다.

범위는 **선택된 C 비트필드 선언의 폭과 부호**로 계산합니다. 일반 멤버의 `// 12bit` 같은 주석을 비트필드 선언으로 바꾸지 않습니다. 코드의 표현 가능한 범위와 통신 규격·설계서가 제한하는 논리 범위는 서로 다를 수 있습니다.

이 버전은 빌드 구성 하나를 선택합니다. 파일별 컴파일 옵션, 컴파일러 내장 헤더/매크로의 자동 수집, `compile_commands.json`, Eclipse `.cproject` 전체 프로젝트 설정 가져오기는 지원하지 않습니다. 필요한 프로젝트 헤더 및 매크로를 등록해야 합니다.

## 검증 범위

제공된 SX3 XML의 25개 매크로를 읽는 검증과 동일한 문제 형태의 합성 C/H·XLSX로 GUI, 선언·정의 연결, 일반 멤버/비트필드 분기, 범위 비교, 오류 차단, 보고서 기록을 검증합니다. 원본 회사 C/H의 실제 분기 조건은 제공되지 않았으므로 사진 속 전체 프로젝트에 대한 해결 완료를 의미하지 않습니다.

fix7의 배열 인덱스 범위 무시와 배열 포인터 함수 인자 처리는 유지합니다. XML을 선택하지 않은 기존 비교 경로도 회귀 테스트합니다. 실제 회사 Excel/DRM 통합은 이번 검증에 포함하지 않습니다.

## 재빌드 및 재현

Windows CPython 3.13 x64에서 다음 명령을 사용합니다.

```powershell
python -m pip install -r requirements-lock.txt -r requirements-build-lock.txt
python run_gui.py --selftest
python run_gui.py --build-options-test build-options-smoke.json
python tools/run_tests.py --gui
python -m PyInstaller --distpath dist --workpath build/fix8 tools/RangeValidator-DirectCompare-final-Windows-x64.spec
```

빌드옵션 smoke test에는 선택적으로 세 번째 인자로 실제 XML 경로를 전달할 수 있습니다. 이 검증은 `OPT_PRESSURE_SENSOR`가 정의된 예시 XML과 임시 합성 소스/Excel 파일을 사용합니다.

전처리기는 BSD 라이선스의 pcpp 1.30을 사용합니다. 라이선스는 `docs/dependency-licenses`에 포함되어 있습니다.
