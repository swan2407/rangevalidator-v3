"""Write completion evidence only after tests, package and isolated run agree."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
load=lambda name:json.loads((ROOT/'docs'/name).read_text(encoding='utf-8-sig'))
t=load('TEST_RUN.json');d=load('DISTRIBUTION.json');v=load('DISTRIBUTION_VALIDATION.json');b=load('BASELINE_78.json');oracle=load('V1_COMMON_DIFFERENTIAL.json')
current={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'rangevalidator').glob('*.py')}
assert t['successful'] and t['tests']==106 and not t['skipped'] and not t['errors'] and not t['failures']
assert b['tests']==78 and b['successful']
assert t['source_sha256']==d['source_sha256']==v['source_sha256']==current
assert not oracle['differences'] and oracle['reference_files_unchanged'] and oracle['git_state_unchanged']
assert v['results']['smoke-test']['gui'] and v['results']['integration-test']['verified']==1
assert hashlib.sha256(Path(d['exe']).read_bytes()).hexdigest()==d['exe_sha256']
assert hashlib.sha256(Path(d['zip']).read_bytes()).hexdigest()==d['zip_sha256']
assert all(hashlib.sha256(Path(p).read_bytes()).hexdigest()==h for p,h in oracle['source_sha256'].items())
summary=f'''# 최종 검증 결과 — 2026-11-03

- 기존 V3 baseline: **78/78 통과**, 실제 Excel 포함, {b['elapsed_seconds']}초.
- 최종 회귀: **{t['tests']}/{t['tests']} 통과**, 실패 0 / 오류 0 / 건너뜀 0,
  {t['elapsed_seconds']}초. `.build-python/python.exe tools/run_tests.py --excel --gui`.
- 컴파일 문법 확인: compileall. 전체 테스트 로그 TEST_RUN.log, 환경/소스 해시 TEST_RUN.json.
- V1 현재 소스 비교: **18개 문서 / 270행 판정 일치**, 차이 0.
  참조 소스 해시와 추적 Git 상태 불변. V1_COMMON_DIFFERENTIAL.json 참조.
- GUI 테스트 5개: 실행·비동기 비교·선택/무효화·로그 자동 표시·보고서·종료·페이지 접근.
- 실제 일반 Excel 테스트: 초기 Reader/수명 관리 5개 및 신규 통합 4개.
  openpyxl/COM 판정 일치, 2개 A/B 문서의 unsigned/signed/빈 셀 등 **8개 승인 셀**
  SaveAs와 재개봉 MATCH, 원본 해시/mtime, 값/수식/셀 서식/행열/시트/병합/이미지
  데이터·위치 보존, 일부 파일 실패 격리를 확인했습니다.
- COM 전용 보존 경로는 native manifest 실패를 모의하여 실제 COM 읽기/저장을
  실행했습니다. DRM_CONFIRM_REQUIRED/UNVERIFIED가 유지됨을 확인했으며,
  이 테스트는 실제 DRM 검증이 아닙니다.
- 잘못된 대상, REVIEW, 중복, 입력/세션 변경은 수정 차단. Common 절/선언 충돌,
  다중 헤더, A/B·다중 시트/문서, 수식/병합/숨김, 파일 오류, CP949 등을 검사했습니다.
- EXE: PyInstaller {d['pyinstaller']} Windows x64 onedir, {d['files']}개 ZIP 파일.
- 배포 ZIP 전 항목 해시 대조 및 새 폴더 추출 검사 통과.
- Windows/System32만 PATH에 남기고 Python 환경변수를 제거한 **ZIP 추출 EXE**:
  네 탭 GUI 실행·종료, pywin32 패키지 내부 로드, 합성 비교→COM SaveAs→원본 보존→
  보고서 자체검사 통과. 외부 Python/MSYS2 로드 없음. DISTRIBUTION_VALIDATION.json 참조.
- 테스트·배포·현재 소스 해시 일치. 기존 V1/V2 소스 수정 없음. V2 접근/포팅 없음.

## 배포 파일

EXE: `{d['exe']}`

SHA-256: `{d['exe_sha256']}`

ZIP: `{d['zip']}`

SHA-256: `{d['zip_sha256']}`

배포 루트와 ZIP 내부에 SHA256SUMS.txt가 있습니다. 라이선스/저작권 고지와
사용안내가 ZIP에 포함됩니다. 빌드 로그: build/PYINSTALLER_BUILD.log.

## 검증 범위의 한계

실제 회사 DRM 읽기·거부·SaveAs 보호 유지, 물리적 클린 PC, EDR 승인,
실제 ACL/파일 잠김, 긴 UNC 경로, 임의 템플릿/모든 Excel 객체, 매우 큰 파일의
성능/메모리, 정책 모달·정지 COM 자동 복구는 미검증 또는 미지원입니다.
보존 지원 밖 객체는 REVIEW로 저장을 차단할 수 있습니다. COM 전용 보존은
Find 내용 영역과 그림 메타데이터에 한정되므로 전체 그림 바이트/영역 밖 서식
보존을 인증하지 않습니다. 실제 DRM 출력은 보호 유지 확인 전 완료가 아닙니다.

초기 개발 중 발견한 uint 순서/COM Address/출력 인코딩 문제에 이어,
이번 단계의 테마·행열 XML 정규화와 GUI 컨트롤 초기화 문제를 수정하고
관련 회귀를 통과했습니다. 임의 XML 차이를 무시하지 않고 실제 적용 서식과
독립적인 값/이미지 검사를 사용합니다. 실패한 중간 시험 배포물은
build/release-history에 보관했고 최종 ZIP은 재빌드·재검증했습니다.
'''
(ROOT/'docs/TEST_RESULTS.md').write_text(summary,encoding='utf-8')
(ROOT/'docs/TASKS.md').write_text('''# Tasks — GUI/EXE milestone

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
''',encoding='utf-8')
print(json.dumps(dict(tests=t['tests'],source_package_test_hashes_equal=True,v1_rows=oracle['rows'],zip_sha256=d['zip_sha256'],exe_sha256=d['exe_sha256']),indent=2))
