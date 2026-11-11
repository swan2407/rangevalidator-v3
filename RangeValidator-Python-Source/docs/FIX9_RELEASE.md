# DirectCompare fix9 (2026-11-11)

- 엑셀 Range/Scope의 숫자 매크로를 등록된 C/H 정의로 해석합니다. NO, YES, SET 같은 이름에 값을 고정하지 않습니다. 빌드옵션 XML 사용 시 해당 C 문맥의 활성 정의를 사용합니다.
- 엑셀과 C 파일명이 달라도 잘린 긴 함수명을 연결합니다. 해당 함수의 매개변수·지역변수를 우선하며, 함수 후보가 여러 개이면 임의로 고르지 않습니다.
- 함수 포인터, 미등록 타입, 복잡한 배열 등 일부 인자의 타입을 해석하지 못해도, 타입이 확인된 다른 변수의 멤버는 비교합니다. 해석이 불확실한 인자 자체는 REVIEW이며, 확인된 선언·정의 타입 충돌도 REVIEW입니다.
- 일반 멤버는 SKIP / 일반 멤버로 표시합니다. 비트필드 범위 비교에서 수치가 다르면 MISMATCH입니다.

Windows ZIP을 새 폴더에 풀고 `RangeValidator-DirectCompare-fix9-Windows-x64.exe`를 실행하세요. Python 설치는 필요 없습니다.

소스 실행: `py -3.13 run_gui.py`

소스 의존성 설치: `py -3.13 -m pip install -r requirements-lock.txt -r requirements-build-lock.txt`

재빌드: `py -3.13 -m PyInstaller --distpath dist --workpath build/directcompare-fix9 tools/RangeValidator-DirectCompare-final-Windows-x64.spec`

배포 점검: `RangeValidator-DirectCompare-fix9-Windows-x64.exe --release-test 결과.json`

이전 버전 번호가 적힌 docs의 검증 기록은 해당 버전의 과거 기록입니다. 이번 배포의 검증 결과는 함께 제공되는 VALIDATION.json을 확인하세요.
