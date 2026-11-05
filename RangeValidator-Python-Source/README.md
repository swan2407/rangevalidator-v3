> **V3.1 (2026-11-04)**: 느린 읽기·PARTIAL_READ·INPUT_CONTENT_CHANGED·SOURCE_CONTEXT_REVIEW 로 모든 행이 막히던 문제 수정, C 분석기 재작성(typedef 사전·매크로 비트폭·중첩 구조체·.c 변수 선언), **[계산 범위 전체 내보내기]** 추가. 원인·수정 내용은 [docs/V3.1_변경_분석.md](docs/V3.1_변경_분석.md).

# RangeValidator Python V3.1

한국어 Tkinter/ttk GUI에서 C/H 선언과 A/B 상세설계서를 비교하고,
확정된 불일치 셀을 Excel COM으로 수정해 별도 파일에 저장합니다.
V1/V2는 수정하지 않았으며 V2 기능을 사용하지 않습니다.

**배포 ZIP:** `dist/RangeValidator-Python-Windows-x64.zip`

ZIP 전체를 새 폴더에 풀고 `RangeValidator-Python.exe`를 실행하세요.
Python/pip/MSYS2 설치는 필요하지 않습니다. COM에는 설치·인증된 Excel과
회사 승인 DRM 클라이언트/정책이 별도로 필요합니다.

- [회사 PC 실행 안내](docs/WINDOWS_GUIDE_KO.md)
- [검증 결과](docs/TEST_RESULTS.md)
- [V1 비교 및 차이](docs/V1_DIFFERENCES.md)
- [DRM 환경 추가 검증](docs/DRM_VALIDATION.md)
- [배포 해시/빌드 기록](docs/DISTRIBUTION.json)
- [PATH 격리 실행 기록](docs/DISTRIBUTION_VALIDATION.json)

실제 회사 DRM 파일의 읽기/거부/출력 보호 유지는 **미검증**입니다.
저장 상태 `SAVED_VERIFIED`는 수정한 셀의 값 검증 결과입니다. DRM 적용 여부는 별도로 확인해야 합니다.

## 개발 실행

GUI, 테스트, 패키지 설치와 EXE 빌드에는 Tcl/Tk가 포함된 **Windows CPython 3.13 x64**를 사용합니다.
아래 경로는 5-2단계에서 확인한 기존 개발 환경입니다. 다른 PC에서는 해당 PC의 Windows CPython 경로를 지정하세요.

```powershell
Set-Location C:\Range\RangeValidator-Python-Source
$rvPython = 'C:\Dev\RangeValidator-Python\.build-python\python.exe'
& $rvPython -m pip check
& $rvPython -B -c "import openpyxl, pythoncom, win32com.client, rangevalidator.gui, rangevalidator.writer, rangevalidator.report"
& $rvPython -B -m unittest discover -s tests
& $rvPython gui_entry.py
& $rvPython -m rangevalidator batch Detail.xlsx --header types.h
```

5-2단계 확인 버전은 Python 3.13.9, pip 25.2, openpyxl 3.1.5, et-xmlfile 2.0.0, pywin32 311,
Pillow 13.0.0, PyInstaller 6.22.3입니다. MSYS2 UCRT64 Python에는 pip와 pywin32가 없었고,
확인한 오프라인 pywin32/Pillow 휠은 Windows CPython 3.13 x64용입니다.
기존 embeddable `.runtime`에는 Tkinter와 pip가 없습니다.

사내 정책이 허용하는 신규 개발 환경에서는 승인된 로컬 휠 저장소에서 같은 인터프리터로 의존성을 설치합니다.

```powershell
$rvWheelhouse = 'C:\Dev\RangeValidator-Python\wheelhouse'
& $rvPython -m pip install --no-index --find-links $rvWheelhouse -r requirements-lock.txt -r requirements-build-lock.txt
& $rvPython -m pip check
```

EXE 패키징 스크립트 `tools/package_windows.py`는 자신을 실행한 Python으로 PyInstaller를 호출합니다.
실제 Excel과 Tk GUI 통합 테스트는 별도 승인 환경에서 각각 `RV_TEST_EXCEL=1`, `RV_TEST_GUI=1`을 사용합니다.

`analyze/read/compare` 초기 CLI와 초기 모델 API는 78개 회귀 호환을 위해 유지합니다.
생산용 GUI와 `batch` CLI는 `workflow/engine`의 C/H 카탈로그 기반 판정을 사용합니다.
CLI는 JSON을 stdout에만 출력하며 입력 경로로 리디렉션하지 마세요.

## 구조

```text
rangevalidator/
  analysis.py       기존 토큰 C 분석/비트필드 범위
  models.py         기존 공통 Workbook/Sheet/Cell
  readers.py        openpyxl, STA COM, 조건부 자동 fallback
  common.py         Common 선언·절·메타데이터
  engine.py         다중 타입/문서 문맥, A/B 추출, V1 판정
  workflow.py       파일별 비교 상태와 세션 무결성
  writer.py         선택 확인·잠금·재분석·COM SaveAs·재개봉 검사
  preservation.py   네이티브 XLSX 값/서식/이미지 보존 근거
  report.py         Summary/Changes/Review/Errors/Results
  diagnostics.py    마스킹 구조화 단계 로그
  gui.py            한국어 4개 탭, 작업 스레드·페이지 표시
  selfcheck.py      명시적으로 실행하는 합성 배포 검사
 gui_entry.py
 tests/              기존 78개 + 새 엔진/GUI/Writer 회귀
 tools/              V1 oracle, 테스트/배포 검사/빌드
 docs/               검증 기록/지원 경계/라이선스
 wheelhouse/         오프라인 휠과 해시
 dist/               EXE onedir + ZIP + 체크섬
```

어떤 입력 원본도 저장하지 않습니다. 입력 해시/mtime, 현재 값, 수식/병합,
분석 근거를 저장 전에 재확인합니다. 일반 정수 범위/AST/CFG/Alias Analysis/
함수 I/O 분석은 구현하지 않았습니다. 수식과 병합 데이터는 자동 수정하지 않습니다.
