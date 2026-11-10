"""RangeValidator-DirectCompare 패치판 실행 진입점.

배포본(exe) 안의 gui_entry 와 같은 동작에 `import rangevalidator_patch` 한 줄을 앞세운 것입니다.
    py -3.13 run_gui.py                 GUI 실행
    py -3.13 run_gui.py --selftest      패치 자가 테스트 (selftest_patch.py)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import rangevalidator_patch  # noqa: E402,F401  ← 원본 모듈을 import 한 뒤 함수들을 교체한다


def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--build-options-test':
        from rangevalidator.build_options_smoke import run
        run(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        return
    if len(sys.argv) > 1 and sys.argv[1] == '--selftest':
        import selftest_patch
        raise SystemExit(selftest_patch.main())
    if len(sys.argv) > 1 and sys.argv[1] == '--batch-read-test':
        from selftest_patch import run_gui_smoke
        run_gui_smoke(sys.argv[2])
        return
    from gui_release import main as gui_main
    gui_main()


if __name__ == '__main__':
    main()
