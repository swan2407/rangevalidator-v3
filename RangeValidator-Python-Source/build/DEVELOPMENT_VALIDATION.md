# 6-1 개발 EXE 검증 기록

- Windows CPython C:DevRangeValidator-Python.build-pythonpython.exe, 3.13.9 x64
- PyInstaller 6.22.3, --onedir --windowed, gui_entry.py
- 전체 테스트 196개, 통과 180개, 실패 0개, 오류 0개, 기존 옵션 건너뜀 16개
- 동결 EXE 직접 실행, GUI 창과 C/H 파일·폴더 입력, Excel 입력, Common 입력 없음, 0% 진행률 확인
- 동결 EXE --integration-test 실행 종료 코드 0
- 합성 MISMATCH 1개 실제 Excel COM 저장, 저장본 재열기 검증, 원본 보존 통과
- 보고서 Summary/Mismatch/Review/Errors/Results 확인
- Mismatch/Review A2 고정 및 전체 AutoFilter 확인
- 검증 전후 Excel 프로세스 없음. 기존 사용자 Excel 창이 없는 환경에서 검증했으므로 기존 사용자 창과의 공존 검증은 남음
- CREATE와 MANUAL_REVIEW의 동결 EXE 동시 저장 및 실제 GUI 입력 전 과정은 미검증
- Python 미설치 별도 Windows PC, 읽기 전용 설치 경로, 회사 DRM 환경은 미검증
- MSYS2 Python 사용 안 함. 신규 패키지 설치 및 hidden import 추가 없음
- 기본 hooks가 Tcl/Tk, pythoncom, pywintypes, win32com, openpyxl, Pillow를 수집
- 빌드 경고에는 플랫폼 선택 모듈(posix/pwd/grp 등), 선택 의존성(lxml/numpy/defusedxml/olefile), win32com.gen_py 등이 있음
- 실제 수행한 GUI 및 COM 저장 경로에서 모듈/DLL 누락 오류 없음. 모든 선택 경로를 실행했다는 의미는 아님
- 소스 기능 변경 없음. tools/build_development.py만 추가
