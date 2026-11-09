# Clang-1 비교 검증

## 도입 판단

현재 배포에서 Clang 채택을 보류한다. Windows x64에서 승인된 libclang과 Python 바인딩을 함께 로드하고 합성 정답과 일치하는지 확인하는 것이 다음 판단 조건이다. 비교 도구는 독립 실행하며 Excel 판정이나 Writer로 결과를 전달하는 경로가 없다.

## 환경과 기준선

- Windows AMD64, 기존 빌드 환경의 CPython 3.13.9 x64, PyInstaller 6.22.3을 사용했다. PATH의 `python.exe`는 실행할 수 없는 Windows 앱 별칭이었고 실제 실행 파일은 `C:\Dev\RangeValidator-Python\.build-python\python.exe`였다.
- PATH와 Program Files, 로컬 프로그램 및 패키지 경로, 작업 저장소와 기존 개발 폴더의 조회 가능한 범위에서 `clang.exe`와 `libclang.dll`을 찾지 못했다. 현재 빌드 Python에서 `clang.cindex`가 없었다. 실제 Clang 버전, 바인딩 버전, DLL 호환성 및 로딩 성능은 미검증이다. 사내에서 사용할 수 있는 설치 채널도 확인되지 않았다. 사내 승인된 오프라인 패키지 또는 승인된 별도 검증 환경이 필요하다.
- 시작 커밋 `c93f37983bb3480b8fd4a28392a35f1a67d04fa0`에는 Excel-6 코드와 초기 Clang 모듈 및 합성 생성기가 함께 있었다. 시작 당시 추적 파일 변경은 없었으며 기존 미추적 `dist/RangeValidator-Excel6-Baseline.zip`을 보존했다. Excel-6 전용 커밋을 식별할 수 없어 이 시작 커밋을 재빌드 출처로 기록했다.
- 기존 개발 빌드의 `--onedir --windowed`와 `gui_entry.py`를 사용해 [별도 기준선 폴더](../dist/RangeValidator-Excel6-Baseline-Rebuild/)에 빌드했다. 1,011개 파일, 49,315,899바이트이며 EXE SHA-256은 `EF3BA7271200CD99393892FFB00796E01E861256232560DCE49574AF42839DFB`다. 빌드 완료와 파일 구성을 확인했다. 이 재빌드 EXE의 실행 검증은 미실시다. [빌드 로그](../build/clang1-baseline-build.log)
- 기존 기준선 EXE와 ZIP의 작업 전후 SHA-256은 각각 `D6840E140D71305FE5864C0A83C0ABA4040C3E0BAB855AEEBD48363FAEF67B94`, `47684AE969E4B8A6A4BBC0227D6054A44F5DE40B2BC1EB42BFBD10DD7A9147AE`로 일치했다.

## AST와 비교 방법

- `rangevalidator/clang_probe.py`는 번역 단위별 AST, 구조체 정의 위치와 전방 선언, typedef의 canonical 선언, 중첩 구조체, 멤버 폭·signedness·범위, 전역·지역·매개변수 선언, 포인터·배열·한정자 층을 JSON으로 직렬화할 수 있게 수집한다. 실제 C 멤버 접근은 참조한 변수와 멤버 선언을 함께 저장한다. 합성 정답은 각 선언의 파일과 줄을 직접 지정하며, AST 내부 선언 ID에는 열도 포함한다.
- 빌드 구성 확인 여부와 대상 타깃을 명시한다. 오류가 있는 복구 AST와 빌드 구성이 미확인인 AST는 신뢰할 수 없는 것으로 표시하고 범위를 비운다. 파일 누락·미정의 식별자·미지원 옵션과 진단을 남긴다. 전방 선언만 있는 레코드와 폭 0 패딩은 각각 타입 또는 범위 미확정으로 기록한다. signedness는 canonical 정수 타입과 enum 기반 타입에서 추출한다. 이는 지정된 Clang 타깃의 해석이며, 사내 컴파일러의 plain int 또는 enum 비트필드 의미는 별도 대조가 필요하다.
- `arguments` 형식 `compile_commands.json`에서 파일별 옵션과 작업 디렉터리를 보존한다. 복수 변형 및 쉘 문자열은 거부한다. Windows x64/x86, C11/C99, `-I`, `-isystem`, `-D`, `-U`, `#if`, `#ifdef`, 매크로 비트폭, `cinst`의 정의 유무를 합성 사례로 준비했다. 회사 컴파일러 전용 플래그는 의미를 확인한 후 변환해야 한다. [Clang Compilation Database 규격](https://clang.llvm.org/docs/JSONCompilationDatabase.html)
- 기존 측은 실제 비교에 사용하는 `build_bitfield_catalog`와 단일 멤버 판정 함수를 호출한다. 중첩 경로 전체를 표현할 수 없으면 미확정으로 계산한다. 동일 이름의 멤버는 원본 구조체 위치를 대조했다. 전방 선언만 있는 변수는 미확정이 정답인 사례로 별도 집계한다.
- [기능 정답](../build/clang1-verification/functional/truth.json), [10개 파일 정답](../build/clang1-verification/stress10/truth.json), [100개 파일 정답](../build/clang1-verification/stress100/truth.json)은 C/H와 별도 파일이다. 기능 데이터에는 별도 typedef, 여러 헤더에 걸친 별칭, shadowing, 같은 멤버명, 포인터/배열/한정자, 중첩 경로, signed·unsigned·plain int·typedef 정수·enum과 유효하지 않은 폭, 누락 헤더, 잘못된 typedef, 타깃 및 조건 불일치를 포함한다.

## 기존 분석기 측정

| 합성 입력 | 발견 구조체 / 비트필드 / 별칭 | 전역 / 지역 / 매개변수 | 범위 정답 / 미확정 | 변수 타입 연결 정답 / 미확정 | 시간 | 최대 작업 집합 |
|---|---:|---:|---:|---:|---:|---:|
| 기능 16파일 | 26 / 41 / 9 | 20 / 2 / 1 | 17 / 10 | 19 / 4 | 0.142초 | 27.09 MiB |
| 스트레스 10파일 | 5 / 109 / 5 | 6 / 4 / 4 | 108 / 4 | 17 / 0 | 0.385초 | 27.09 MiB |
| 스트레스 100파일 | 50 / 199 / 50 | 51 / 49 / 49 | 198 / 49 | 197 / 0 | 1.613초 | 27.87 MiB |

발견 수는 스위트 전체의 고유 선언 인벤토리이며 오류 검증용 파일의 선언도 포함한다. 정답 대조는 번역 단위별로 하므로 반복 include의 선언이 여러 번 평가될 수 있다. 기능 스위트의 정상 구성에서 구조체 정의 12건이 맞았고 typedef 연결 9건이 맞았으며 2건은 미확정이었다. 멤버 접근 17건 중 범위와 원본 선언이 맞은 경우 9건, 일반 멤버 제외 2건, 미확정 6건이었다. 스트레스 범위 미확정은 전처리 분기에 따른 폭이다. 확정된 구조체 타입의 잘못된 연결은 세 스위트에서 관측되지 않았다. 폭 0인 이름 없는 비트필드 1건은 기존 인벤토리에서 누락됐다.

시간은 각 스위트의 독립 프로세스에서 수집과 정답 대조를 한 번 수행한 값이다. 최대 메모리는 Python 런타임을 포함한 Windows 프로세스 최대 작업 집합이다. Clang 측은 세 스위트 모두 `DEPENDENCY_UNAVAILABLE`이므로 구조체·typedef·접근 정확성, 진단, 시간 및 메모리 비교는 미측정이다. [기계 판독용 요약](../build/clang1-verification/summary.json) · [기능 케이스별 대조](../build/clang1-verification/functional/legacy.json)

## DLL 및 배포 검토

libclang C API에는 선언 커서, canonical type, 비트필드 폭을 조회할 인터페이스가 있다. 이 수준으로 사내 코드에 필요한 선언 또는 전처리 정보가 충분하지 않다면 C++ LibTooling의 AST와 CompilationDatabase를 검토할 수 있다. LibTooling은 Clang C++ API와 빌드 버전 관리가 추가된다. [Libclang](https://clang.llvm.org/docs/LibClang.html) · [LibTooling](https://clang.llvm.org/docs/LibTooling.html)

승인 환경에서 동일 릴리스의 x64 DLL과 바인딩을 먼저 대조한다. 전체 DLL 경로를 `--dll`로 전달하면 호환성 검사를 켠 채 `Index.create()`로 로딩을 확인하고 실제 DLL 버전·크기·바인딩 정보를 기록한다. PyInstaller one-folder에는 `clang.cindex`, `libclang.dll`, 해당 DLL의 PE import에 필요한 런타임, 필요한 경우 Clang 리소스 헤더를 포함할 수 있다. 번들 기준 절대 경로를 `Config.set_library_file()`에 전달하고 종속 DLL이 별도 디렉터리에 있으면 `os.add_dll_directory()` 적용을 검증한다. 리소스 헤더 위치는 `-resource-dir`로 제어할 수 있다. 외부 Windows SDK나 업체 헤더가 필요한 번역 단위에는 해당 경로도 공급해야 한다. 이 구성으로 별도 Python이나 LLVM 설치 없이 동작하는지 깨끗한 Windows x64에서 검증한다. 이번에 실측한 49,315,899바이트는 기준선 크기이며 Clang의 DLL·리소스·종속 런타임 증가분은 미측정이다. [PyInstaller 바이너리 포함](https://pyinstaller.org/en/stable/spec-files.html#adding-binary-files)

LLVM 배포본의 Apache-2.0 with LLVM-exception 라이선스와 포함 컴포넌트의 고지 파일을 확인한다. 승인할 DLL, 바인딩, 런타임과 리소스의 정확한 배포본을 확정한 후 고지 및 재배포 조건을 각각 검토한다. [LLVM 라이선스 정책](https://llvm.org/docs/DeveloperPolicy.html#copyright-license-and-patents)

## 재현과 Clang-2 진입 조건

전체 회귀는 GUI 선택 테스트를 켜고 실행했다. 총 228개 중 216개 통과, 12개 건너뜀, 실패·오류 0개였다. 기존 224개는 213개 통과와 실제 Excel 선택 테스트 11개 건너뜀으로 기준과 일치했다. 신규 4개는 3개 통과했고 실제 AST 검증 1개는 `clang.cindex`가 없어 실행하지 못했다. [회귀 로그](../build/clang1-regression.log)

```powershell
python -B tools/compare_clang_parser.py build/clang1-approved --dll C:\approved\LLVM\bin\libclang.dll
```

출력 경로는 새 디렉터리로 지정한다. Clang을 실행할 수 없는 경우에도 기존 분석 결과와 독립 정답은 저장되며 종료 코드는 1이다. 로드 오류와 정답 불일치도 종료 코드 1이며 JSON 상태로 구분한다. `RV_LIBCLANG_DLL`에 승인 DLL을 설정하면 `test_clang_probe.py`의 AST 테스트에서도 같은 바이너리를 사용한다.

다음 단계에서는 승인된 DLL과 바인딩에서 기능·10파일·100파일 정답을 재실행하고, 사내 컴파일러의 옵션 매핑, 시스템 헤더, `cinst`, plain int 및 enum의 타깃 의미를 확인한다. 같은 선언이 빌드 변형별로 다른 레이아웃을 가질 때의 식별과 병합 규칙을 정하고, 실험용 one-folder에서 DLL 의존성과 크기 증가를 측정한다. 이 검증 결과로 Clang-2에서 읽기 전용 보조 분석을 진행할지 판단한다.
