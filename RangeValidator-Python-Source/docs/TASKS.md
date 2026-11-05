# Tasks — GUI/EXE milestone

- [x] 기존 78개 테스트 먼저 실행: 78/78 통과.
- [x] V1 실제 Common/Reader/Resolver/Comparator/테스트 조사, V1/V2 수정 없음.
- [x] Common 절/변수/범위 메타데이터, 다중 헤더와 선언 충돌, A/B·다중 문서/시트.
- [x] 원래 파서·Workbook 모델 유지, V1 호환 판정 및 안전한 MISMATCH 후보.
- [x] 조건부 openpyxl→COM 자동 읽기, STA/보안 플래그/Find/Value2/안전한 종료.
- [x] 선택 대상만 입력 잠금/해시/재분석/셀 재확인 후 COM SaveAs, 새 경로 및 재검증.
- [x] 네이티브 값/수식/적용 서식/행열/시트/병합/이미지 보존 검사.
- [x] COM 전용 메모리 메타데이터 검사, 보호 유지 미검증 상태 분리.
- [x] 한국어 4개 탭 GUI, 작업 스레드, 필터/페이지/선택·인코딩·진행·마스킹 로그.
- [x] Summary/Changes/Review/Errors/Results 보고서와 일부 파일 실패 처리.
- [x] V1 18개 합성 문서/270행 일치, 초기 차이 2건 검토·문서화.
- [x] 최종 106/106 회귀, 실제 일반 Excel Reader/Writer 및 GUI 포함.
- [x] PyInstaller x64 onedir EXE/ZIP, 런타임/pywin32/고지/체크섬 포함.
- [x] ZIP 추출 해시·Windows-only PATH GUI와 실제 COM 저장 자체검사.
- [ ] 실제 회사 DRM allow/deny 및 출력/보고서 보호 유지 검증.
- [ ] 물리적 클린 PC/EDR/Office 버전/실제 잠김·ACL·긴 UNC·대규모 성능 검증.
- [ ] 지원 밖 Excel 객체/범위 밖 COM 서식·이미지 payload 보존, hung COM supervision.

일반 정수 범위, V2 AST, CFG, Alias Analysis, 함수 I/O는 추가하지 않았습니다.
검증 결과/배포 해시: TEST_RESULTS.md, 지원 경계: WINDOWS_GUIDE_KO.md.


## 진단 로그 한정 수정 (2026-11-03)

- [x] 기존 소스/EXE 진단 탭 및 누락 단계 확인, 기존 배포 해시/실행 증거 보관.
- [x] 진단 보기 버튼·별도 창·읽기 실패 자동 표시·실패 후 진단/결과 유지.
- [x] 단계별 성공/실패·HRESULT/Win32·마지막 성공·Close/Quit/종료 및 마스킹 보완.
- [x] Reader 비진단 AST 동일, 엔진/Writer/보고서/공통 모델 해시 불변.
- [x] 관련 27개 테스트만 통과. 전체/분석/Writer/보고서 회귀는 실행하지 않음.
- [x] onedir 재패키징, 새 ZIP 추출·Windows-only PATH 진단 전용 EXE 검사 통과.
- [ ] 실제 회사 DRM 읽기·정책 거부 및 별도 클린 PC 확인.

최신 진단 감사와 해시: DIAGNOSTIC_AUDIT.md. 이전 106개 전체 회귀는 과거 기록입니다.


## 실제 DRM 대량 읽기 실패 보고에 대한 코드 수정 (2026-11-03)

- [x] Input changed 문자열/tuple 판정/예외 전파 원인 조사, 미확정 환경 원인 구분.
- [x] 입력별 변경·메타데이터 변경·조회 불가 분리, 근거 유지 및 기존 저장 차단 유지.
- [x] Reader 실패를 파일 ERROR로 분리, 부분 성공·미추출·미비교 상태와 GUI 집계.
- [x] Find 단계/예외 타입/HRESULT/분류/파일·시트 번호, bounded read-only 재시도와 병합 캐시.
- [x] HRESULT 없는 Python 예외/예산/불일치 Find를 COM 거부와 구분.
- [x] 관련 41개 테스트 통과. 전체/Writer/보고서 테스트는 실행하지 않음.
- [x] 새 버전 onedir EXE/ZIP, 추출본 Windows-only PATH 정상 비교+파일 ERROR 분리 확인.
- [x] 파서·엔진·Common·Writer·보고서 및 Workbook 모델 유지, V1/V2 접근/수정 없음.
- [ ] 실제 회사 Find 예외 타입/검색 단계/분류와 입력별 변경 flag 재확인.
- [ ] 실제 회사 DRM/UNC에서 읽기·retry 회복·대량 성능은 별도 미검증.

최신 근거/해시: DRM_BATCH_AUDIT.md 및 DRM_BATCH_* 기록.


## Reference Excel COM Reader rebuild (2026-11-03)

- [x] Read supplied excel_io.py without modifying V1/V2; port its Find/End/Value2 flow into V3.
- [x] Remove mandatory first_column bounds consistency; reuse one owned Excel per comparison thread/session.
- [x] Preserve Workbook model, parser/engine/GUI/Writer/report interfaces and source implementations.
- [x] Fail closed on uncertain sampling, caps, limits or incomplete extraction; mask diagnostic paths/data.
- [x] Compare reference Writer behavior; retain all existing Writer safety policies.
- [x] Final related-only tests: 58 passed, no failures/errors/skips; actual Excel 75000-cell read and A/B/Common engine parity passed.
- [x] New Windows x64 EXE/ZIP created; extracted-file hashes and Windows-only PATH GUI/COM/batch probes passed. See DISTRIBUTION.json and COM_REFERENCE_DISTRIBUTION_VALIDATION.json.
- [ ] Company-local actual DRM allow/deny, output/report protection and clean PC verification.

Detailed behavior and company-PC recheck: COM_READER_REBUILD.md. No full test-suite execution is authorized for this task.

- [x] Sheet/cell read errors are isolated; readable data is compared, failures appear as ERROR, all partial-workbook fixes are blocked. Related tests: 60 passed.
- [x] Partial-read Windows x64 EXE/ZIP rebuilt; ZIP extraction, hashes, GUI and runtime/batch probes passed. See COM_REFERENCE_DISTRIBUTION_VALIDATION.json.

- [x] COM ?? ??? ?? excel_io.py? Find/UsedRange/End ??? ???? ??; V3 ????????? ?? ??. ?? ??? 60/60 ? ?? EXE ?? ??.

## Detailed diagnostics (2026-11-04)

- [x] Include sheet names and cell coordinates in COM read failures; include sheet, cell, and engine reason for each REVIEW in diagnostic log.
- [x] Keep workbook paths and cell contents out of the default diagnostic log.
- [ ] Actual company DRM diagnostics verification remains pending.


## Header-only GUI input validation (2026-11-04)

- [x] Common Parameter is optional in GUI; C headers and detailed-design workbooks remain required.
- [x] Header-only workflow regression passed; GUI package rebuilt.

