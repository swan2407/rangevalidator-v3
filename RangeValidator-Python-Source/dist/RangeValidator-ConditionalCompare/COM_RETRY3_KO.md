# COMRetry3 배포 변경 사항

2026-11-09. 실행 파일은 `dist/RangeValidator-COMRetry3/RangeValidator-COMRetry3.exe`.
실행 파일과 `_internal` 폴더를 함께 보관한다. FlexibleCatalog의 구조체 비교 기능을 포함한다.

## 읽기 재시도

`0x80020009`가 발생한 읽기 호출을 최초 시도 포함 최대 3회 시도한다.
첫 실패 후 0.1초, 두 번째 실패 후 0.25초를 기다리고 같은 스레드에서 다시 호출한다.
중간에 성공하면 읽기를 계속한다. 3회 모두 실패한 시트는 기존 방식대로 PARTIAL_READ로 남기고 다음 시트를 읽는다.
Find 등의 대체 읽기 경로가 성공하면 실패한 기본 호출만으로 PARTIAL로 판정하지 않는다.

셀 값·수식뿐 아니라 셀 범위 취득, 시트 수·이름·표시 상태, 병합 상태·주소,
UsedRange 및 검색 좌표 읽기에도 제한 재시도를 적용한다.
파일 열기, 저장, 닫기, Excel 종료는 이 읽기 재시도 대상으로 넣지 않는다.

재시도 로그에는 1/3, 2/3, 3/3 시도 정보와 HRESULT가 남는다.
재시도 후 성공 로그는 일반 반복 로그 생략 대상에서 제외한다.
최종 실패에는 실패 API, 내부 HRESULT와 마지막 시도 횟수를 유지한다.

재시도는 일시적 오류에서 복구할 기회를 제공하며, 지속적으로 실패하는 파일의 원인을 자동 해결하지는 않는다.

## 검증 기록

- 전체 회귀 테스트 `build/com-retry3-tests.log`
- 실제 Excel 읽기 중 합성 `pywintypes.com_error`를 주입한 복구 검증 `build/com-retry3-verification/result.json`
- 배포 빌드 `build/com-retry3-build.log`
- 배포 EXE 실행 확인 `build/com-retry3-smoke.json`

전체 회귀 테스트 269개 중 257개 통과, 선택 실행 조건에 따른 12개 건너뜀.
실제 Excel 읽기 중 주입한 오류가 세 번째 시도에서 복구되는 경우와,
지속 실패 시 세 번에서 중단하고 다음 시트를 읽는 경우를 확인했다. 합성 원본 Excel의 해시는 유지됐다.
