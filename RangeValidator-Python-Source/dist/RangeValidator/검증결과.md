# 6-2 실제 Excel COM 통합 검증 결과

재빌드한 One-folder EXE의 --full-integration-test 명시적 모드로 수행했습니다.
최종 실행은 build/acceptance-6-2-final.json에 기록했으며 종료 코드는 0입니다.
실행 작업 폴더는 C:Range로 EXE 설치 폴더와 다릅니다.
검증한 EXE는 dist/RangeValidator/RangeValidator.exe와 바이트 단위로 같은 빌드입니다.

## 실제 확인

- 실제 Microsoft Excel COM Reader로 합성 상세설계서 두 개를 읽음
- 한글 파일명과 1.합성검증 / 2.추가검증 시트 처리
- MATCH 4 / MISMATCH 2 / CREATE 2 / REVIEW 7
- 파일·시트별 탐색, 판정 필터, 변수 검색은 실제 Tk 앱 콜백으로 확인
- 분석·저장 완료 100% 확인
- AUTO_MISMATCH 2 / AUTO_CREATE 2 / MANUAL_REVIEW 2 저장·재열기 검증 완료
- 추가 수동 요청 중 수식 셀 2개와 병합 비앵커 1개는 SKIPPED
- 선택하지 않은 REVIEW와 MATCH는 유지
- 원본 파일 두 개의 바이트가 수정 전후 동일
- 수정본 두 개를 별도의 COM Reader로 다시 열어 값·수식 여부를 확인
- 출력 경로 충돌 파일은 SAVE_FAILED, 다른 파일은 SAVED_VERIFIED로 처리
- Summary / Mismatch / Review / Errors / Results와 이력 6건·제외 3건 확인
- Mismatch와 Review의 A2 고정 및 전체 데이터 AutoFilter 확인
- 종료 후 RangeValidator 및 Excel 프로세스 잔류 없음
- 전체 단위 테스트 196개, 통과 180개, 실패·오류 0개, 기존 옵션 건너뜀 16개

## 발견 오류와 최소 수정

첫 실행의 실제 저장은 완료됐으나 이후 독립 read_com 재검증에서
Excel 정리 실패가 발생했습니다. Reader가 마지막 com_sheet 프록시를 정리
시점까지 보유하는 경로를 확인했습니다.
com_values.py에서 통합 문서를 닫기 전에 마지막 시트 참조를 해제하도록 했고,
전체 테스트와 동일 EXE 통합 점검을 재실행해 통과했습니다.
Range/Value2/Formula 데이터 획득과 저장 정책에는 변경이 없습니다.

gui_entry.py에는 명시적 --full-integration-test 진입점과 실패 기록만 추가했습니다.
rangevalidator/development_probe.py는 합성 데이터 생성과 실제 COM 검증을 수행하는 개발 점검 코드입니다.
tools/run_frozen_integration.py --excel로 명시적으로 실행할 수 있습니다.
재실행 때 기존 결과를 보존하도록 새로운 수신 JSON 경로를 사용합니다.

## 직접 GUI 조작과 자동 점검의 구분

기존 EXE 입력 화면과 헤더 선택은 직접 관측했습니다.
파일 대화상자 조작에서 자동화 도구의 포커스/캐시 대상이 일치하지 않아
수동 파일 선택·REVIEW 입력 대화상자의 전 과정은 미검증입니다.
최종 점검은 동결 EXE 안의 실제 Tk 큐·작업 스레드·탐색·필터 콜백을 사용했고,
수동 요청 모델에는 테스트가 명시적 값을 전달했습니다. 사람의 직접 입력 검증과 구분합니다.

## 추가 확인이 필요한 환경

- 실제 사용자 Excel 창과의 공존. 테스트 시작 전 사용자 Excel 프로세스는 없었음
- 읽기 전용 설치 폴더 권한에서의 실행
- Python 미설치 별도 PC
- 회사 DRM 적용·COM 제한·출력 정책
- 최종 배포 로그 비활성화 및 One-file 빌드

회사 실제 자료와 DRM 파일은 사용하지 않았고 보안 정책을 변경하지 않았습니다.
