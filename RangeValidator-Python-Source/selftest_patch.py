# -*- coding: utf-8 -*-
"""패치 자가 테스트 — 임시 폴더에 C/H 와 상세설계서(xlsx)를 만들어 비교 결과를 기대값과 대조한다.

    py -3.13 selftest_patch.py            (또는  py -3.13 run_gui.py --selftest)
openpyxl 만 있으면 된다(Excel/COM 불필요). 각 줄: PASS/FAIL 이름 [실제 상태들]
"""
import shutil
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import rangevalidator_patch  # noqa: E402,F401
from openpyxl import Workbook  # noqa: E402
from rangevalidator.workflow import run_comparison  # noqa: E402

PRE = "typedef unsigned char u1;\ntypedef unsigned short u2;\n"
DEV = PRE + "typedef struct { u1 active:1; u1 mode:3; u2 level:5; } Dev_t;"
ROWS = [("g_dev.active", "u1", "0~1"), ("g_dev.mode", "u1", "0~7"), ("g_dev.level", "u2", "0~31")]


def guard(name, body):
    return f"#ifndef {name}\n#define {name}\n{body}\n#endif\n"


def arrow(rows):
    return [(n.replace('.', '->'), t, r) for n, t, r in rows]


def files(types=DEV, decl="extern Dev_t g_dev;", c='#include "g.h"\nvoid Motor_Run(void){ }\n', extra=None):
    f = {"types.h": guard("TYPES_H", types), "g.h": guard("G_H", '#include "types.h"\n' + decl), "motor.c": c}
    if extra:
        f.update(extra)
    return f


def build(root, sources, rows, sheet="1. Motor_Run", xlsx="motor_Detail.xlsx", nsheets=1, stray=None):
    root = Path(root)
    paths = []
    for rel, text in sources.items():
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding="utf-8")
        paths.append(str(p))
    w = Workbook()
    cover = w.active
    cover.title = "표지"
    cover["A1"] = "상세설계서"
    if stray:
        cover.cell(*stray, "x")
    for k in range(nsheets):
        s = w.create_sheet(sheet if nsheets == 1 else f"{k + 1}. Motor_Run")
        s["A1"] = "motor.c"
        s["A2"] = "Global(Input)"
        for col, lab in enumerate(["Name", "Type", "Range"], 1):
            s.cell(4, col, lab)
        for r, (n, t, rg) in enumerate(rows, 5):
            s.cell(r, 1, n)
            s.cell(r, 2, t)
            s.cell(r, 3, rg)
    d = root / xlsx
    w.save(d)
    return paths, str(d)


CASES = []


def case(name, sources, rows, expect, **kw):
    CASES.append((name, sources, rows, expect, kw))


M3 = ['MATCH'] * 3
# --- 3번: 선언 수집 ---------------------------------------------------------------------------
case("AUTOSAR VAR() 선언 (Compiler.h 없음)", files(decl="extern VAR(Dev_t, MOTOR_VAR) g_dev;"), ROWS, M3)
case("AUTOSAR CONST() 선언", files(decl="extern CONST(Dev_t, MOTOR_CONST) g_dev;"), ROWS, M3)
case("AUTOSAR P2VAR() 포인터 선언 (->)", files(decl="extern P2VAR(Dev_t, AUTOMATIC, APPL) g_dev;"), arrow(ROWS), M3)
COMP = guard("COMPILER_H", "#define AUTOMATIC\n#define VAR(type, memclass) type\n#define CONST(type, memclass) const type\n#define FUNC(rettype, memclass) rettype\n")
case("Compiler.h 의 VAR()/FUNC() 전개", files(decl='#include "Compiler.h"\nextern VAR(Dev_t, AUTOMATIC) g_dev;',
                                           c='#include "g.h"\nFUNC(void, CODE) Motor_Run(void){ }\n', extra={"Compiler.h": COMP}), ROWS, M3)
case("선언 뒤 __attribute__((section))", files(decl='extern Dev_t g_dev __attribute__((section(".x")));'), ROWS, M3)
case("구조체 __attribute__((packed))", files(types=PRE + "typedef struct { u1 active:1; u1 mode:3; u2 level:5; } __attribute__((packed)) Dev_t;"), ROWS, M3)
case("#if 0 안의 열린 중괄호", files(decl="#if 0\nvoid f(void){\n#endif\nextern Dev_t g_dev;"), ROWS, M3)
case("#ifdef 가지별 함수 머리(중괄호 불균형)", files(decl="#ifdef A\nvoid f(int x){\n#else\nvoid f(void){\n#endif\n}\nextern Dev_t g_dev;"), ROWS, M3)
case("#if 0 … #else 선언 #endif", files(decl="#if 0\nextern Dev2_t g_dev;\n#else\nextern Dev_t g_dev;\n#endif"), ROWS, M3)
# --- 2번: 비트필드 인식 ------------------------------------------------------------------------
case("비트폭이 enum 상수", files(types=PRE + "enum { W_MODE = 3, W_LVL = 5 };\ntypedef struct { u1 active:1; u1 mode:W_MODE; u2 level:W_LVL; } Dev_t;"), ROWS, M3)
case("비트폭 enum 이 다른 헤더", files(types=PRE + '#include "w.h"\ntypedef struct { u1 active:1; u1 mode:W_MODE; u2 level:5; } Dev_t;',
                              extra={"w.h": guard("W_H", "typedef enum { W_MODE = 3 } W_e;")}), ROWS, M3)
case("기반 타입 uint8/uint16 미정의 (관용 폴백)", files(types="typedef struct { uint8 active:1; uint8 mode:3; uint16 level:5; } Dev_t;"),
     [("g_dev.active", "uint8", "0~1"), ("g_dev.mode", "uint8", "0~7"), ("g_dev.level", "uint16", "0~31")], M3)
case("enum 타입 비트필드", files(types=PRE + "typedef enum { A, B } E_t;\ntypedef struct { E_t active:1; u1 mode:3; u2 level:5; } Dev_t;"),
     [("g_dev.active", "E_t", "0~1"), ("g_dev.mode", "u1", "0~7"), ("g_dev.level", "u2", "0~31")], M3)
case("typedef union 레지스터 g.bit.x", files(types=PRE + "typedef union { u2 all; struct { u1 active:1; u1 mode:3; u2 level:5; } bit; } Dev_t;"),
     [("g_dev.bit.active", "u1", "0~1"), ("g_dev.bit.mode", "u1", "0~7"), ("g_dev.all", "u2", "0~65535")], ['MATCH', 'MATCH', 'SKIP'])
NEST = PRE + "typedef struct { u1 active:1; u1 mode:3; u2 level:5; } Cfg_t;\ntypedef struct { u1 id; Cfg_t cfg; Cfg_t *pcfg; Cfg_t arr[2]; } Dev_t;"
case("중첩 구조체 경로 a.b.x / a.p->x / a.arr[i].x", files(types=NEST),
     [("g_dev.cfg.active", "u1", "0~1"), ("g_dev.pcfg->mode", "u1", "0~7"), ("g_dev.arr[1].level", "u2", "0~31")], M3)
case("2차원 배열 변수 g[i][j].x", files(decl="extern Dev_t g_dev[2][3];"), [("g_dev[i][j].active", "u1", "0~1")], ['MATCH'])
# --- 기존 동작 유지(회귀) ------------------------------------------------------------------------
case("회귀: typedef 별칭 체인", files(decl="typedef Dev_t Dev_c_t;\nextern Dev_c_t g_dev;"), ROWS, M3)
case("회귀: 포인터 typedef (Dev_p)", files(decl="typedef Dev_t *Dev_p;\nextern Dev_p g_dev;"), arrow(ROWS), M3)
case("회귀: const 구조체 인스턴스", files(decl="extern const Dev_t g_dev;", c='#include "g.h"\nconst Dev_t g_dev = {1,2,3};\nvoid Motor_Run(void){ }\n'), ROWS, M3)
case("회귀: 함수 매개변수 포인터", files(decl="", c='#include "g.h"\nvoid Motor_Run(Dev_t *p){ }\n'), arrow([("p.active", "u1", "0~1")]), ['MATCH'])
case("회귀: 지역 static 변수", files(decl="", c='#include "g.h"\nvoid Motor_Run(void){ static Dev_t l; }\n'), [("l.active", "u1", "0~1")], ['MATCH'])
case("회귀: 범위 틀림 → MISMATCH", files(), [("g_dev.mode", "u1", "0~3")], ['MISMATCH'])
case("회귀: #ifdef 양쪽 선언 충돌 → REVIEW 유지", files(types=DEV + "\ntypedef struct { u1 active:2; u1 mode:3; u2 level:5; } Dev2_t;",
                                              decl="#ifdef A\nextern Dev2_t g_dev;\n#else\nextern Dev_t g_dev;\n#endif"), ROWS[:1], ['REVIEW'])
case("회귀: 없는 멤버 → REVIEW", files(), [("g_dev.nothing", "u1", "0~1")], ['REVIEW'])
# --- 함수 찾기 ----------------------------------------------------------------------------------
TWO = {"motor/motor_types.h": guard("MT_H", DEV), "motor/motor.c": '#include "motor_types.h"\nstatic Dev_t g_dev;\nvoid Motor_Run(void){}\n',
       "fan/fan_types.h": guard("FT_H", PRE + "typedef struct { u1 active:2; u1 mode:4; u2 level:6; } Dev_t;"),
       "fan/fan.c": '#include "fan_types.h"\nstatic Dev_t g_dev;\nvoid Fan_Run(void){}\n'}
case("xlsx 이름이 다른 모듈이어도 함수 파일 우선", TWO, ROWS, M3, xlsx="fan_Detail.xlsx")
case("시트명 '3.2 Motor_Run' 에서 함수명 추출", TWO, ROWS, M3, sheet="3.2 Motor_Run", xlsx="Detail.xlsx")
case("시트명 '3. 모터 구동 (Motor_Run)'", TWO, ROWS, M3, sheet="3. 모터 구동 (Motor_Run)", xlsx="Detail.xlsx")
# --- 사용자 REVIEW 사례 -----------------------------------------------------------------------
case("NVM/사용자 매크로: 반복 include, 비활성 undef, section pragma",
     files(types=DEV + "\n#define MY_RAM\n#define NVM MY_RAM\n",
           decl='#include "types.h"\n#if 0\n#undef NVM\n#endif\n#pragma section farbss = DATA\nextern NVM Dev_t g_dev;'),
     ROWS, M3)
case("별도 include 폴더의 STRUCT/NCM 별칭 연결",
     {'include/types.h': guard('TYPES_H', PRE + '\n#define STRUCT struct\n#define NCM\n'
                              'typedef STRUCT { u1 active:1; u1 mode:3; u2 level:5; } Dev_t;'),
      'motor.c': '#include "types.h"\n#pragma section farbss = DATA\nNCM Dev_t g_dev;\nvoid Motor_Run(void) {}'},
     ROWS, M3)

NCM_TYPES = '''typedef union {
    struct { unsigned int state:5; } T;
    unsigned int raw;
} ncm_id2_union_t;
extern NCM ncm_id2_union_t ncm_id2;
'''
NCM_ROWS = [('ncm_id2.T.state', 'unsigned int', '0~31')]
case('NCM 정의 미수집, 헤더 extern과 C 정의 및 중첩 union',
     {'NCM_Data.h': NCM_TYPES, 'Ncm_Data.c': '#include "NCM_Data.h"\nNCM ncm_id2_union_t ncm_id2;'},
     NCM_ROWS, ['MATCH'], xlsx='Ncm_Data_Detail.xlsx', reason_contains={0: "'NCM' 선언 수식어"})
case('include 미수집, 다른 이름의 설계서에서도 동일한 NCM 타입 연결',
     {'NCM_Data.h': NCM_TYPES, 'Ncm_Data.c': 'NCM ncm_id2_union_t ncm_id2;'},
     NCM_ROWS, ['MATCH'])
case('NCM 선언의 실제 타입 충돌은 REVIEW',
     {'NCM_Data.h': NCM_TYPES, 'Ncm_Data.c': '#include "NCM_Data.h"\n'
      'typedef struct { unsigned int state:2; } Other;\nNCM Other ncm_id2;'},
     NCM_ROWS, ['REVIEW'], xlsx='Ncm_Data_Detail.xlsx', reason_contains={0: '타입 후보'})

CHARGING_TYPES = ('typedef struct { unsigned int state:5; } rec_data_charging_t;\n'
                  'extern NVM rec_data_charging_t rec_data_charging;')
CHARGING_ROWS = [('rec_data_charging.state', 'unsigned int', '0~31')]
case('NVM 정의 미수집, 헤더 extern과 C 정의의 일반 구조체',
     {'Charging.h': CHARGING_TYPES, 'Charging.c': '#include "Charging.h"\n'
      'NVM rec_data_charging_t rec_data_charging;'},
     CHARGING_ROWS, ['MATCH'], xlsx='Charging_Detail.xlsx', reason_contains={0: "'NVM' 선언 수식어"})
case('NVM 매크로와 include 없이 헤더와 C의 동일 타입 연결',
     {'Charging.h': CHARGING_TYPES, 'Charging.c': 'NVM rec_data_charging_t rec_data_charging;'},
     CHARGING_ROWS, ['MATCH'])

GENERIC_TYPES = 'typedef struct { unsigned int state:3; } b;\nextern a b c;'
GENERIC_ROWS = [('c.state', 'unsigned int', '0~7')]
case('extern a b c와 a b c, 등록된 b 타입으로 연결',
     {'Example.h': GENERIC_TYPES, 'Example.c': '#include "Example.h"\na b c;'},
     GENERIC_ROWS, ['MATCH'], xlsx='Example_Detail.xlsx', reason_contains={0: "'a' 선언 수식어로 해석, 타입 'b'"})
case('일반 접두어의 헤더와 C가 include 없이 동일한 타입으로 연결',
     {'Example.h': GENERIC_TYPES, 'Example.c': 'a b c;'}, GENERIC_ROWS, ['MATCH'])
case('타입으로 정의된 접두어의 선언은 임의로 비교하지 않음',
     {'Example.h': 'typedef unsigned int a;\n' + GENERIC_TYPES,
      'Example.c': '#include "Example.h"\na b c;'}, GENERIC_ROWS, ['REVIEW'])
case('등록된 타입 변경 매크로는 일반 접두어로 추정하지 않음',
     {'Example.h': GENERIC_TYPES, 'Example.c': '#include "Example.h"\na b c;',
      'Config.h': '#define a unsigned int\n'}, GENERIC_ROWS, ['REVIEW'])
case('일반 접두어가 같아도 실제 타입이 다른 선언은 REVIEW',
     {'Example.h': GENERIC_TYPES, 'Example.c': '#include "Example.h"\n'
      'typedef struct { unsigned int state:2; } other;\na other c;'},
     GENERIC_ROWS, ['REVIEW'], xlsx='Example_Detail.xlsx', reason_contains={0: '타입 후보'})

REC = PRE + """#define ITEM_COUNT 4
typedef struct { u1 flag:1; u1 ordinary; } Item_t;
typedef union {
    struct { u1 bit:1; } status;
    Item_t items[ITEM_COUNT];
    Item_t literal[4];
    Item_t *pointer;
    u1 data[4];
} Rec_t;
"""
REC_ROWS = [('rec->items[i].flag', 'u1', '0~1'), ('rec->status.bit', 'u1', '0~1'),
            ('rec->items[i].ordinary', 'u1', '0~255'), ('rec->pointer->flag', 'u1', '0~1'),
            ('rec->pointer.flag', 'u1', '0~1'), ('rec->literal[4].flag', 'u1', '0~1'),
            ('rec->data[4]', 'u1', '0~255')]
case("포인터 매개변수와 중첩 union, 배열 및 접근 검증",
     files(types=REC, decl='', c='#include "g.h"\nvoid Motor_Run(Rec_t *rec) {}\n'),
     REC_ROWS, ['MATCH', 'MATCH', 'SKIP', 'MATCH', 'REVIEW', 'REVIEW', 'REVIEW'],
     reason_contains={4: "'pointer'의 멤버에는 '->'로 접근", 5: "'literal[4]'는 배열 범위를 벗어났습니다 (0~3)",
                      6: '[4]가 원소 개수 표기인지 확인'})

TRUNCATED = 'chg_remaining_time_qcs_100m'
TRUNCATED_C = f'#include "g.h"\nvoid {TRUNCATED}s(void) {{ static Dev_t ls_rct_qcs_100ms; }}\n'
LOCAL_ROWS = [('ls_rct_qcs_100ms.active', 'u1', '0~1')]
case("잘린 시트 함수명으로 지역 선언 연결", files(decl='', c=TRUNCATED_C), LOCAL_ROWS, ['MATCH'],
     sheet=f'12. {TRUNCATED}', reason_contains={0: f"→ '{TRUNCATED}s'"})
case("잘린 시트 함수명의 후보가 둘이면 REVIEW",
     files(decl='', c=TRUNCATED_C + f'void {TRUNCATED}ax(void) {{ static Dev_t ls_rct_qcs_100ms; }}\n'),
     LOCAL_ROWS, ['REVIEW'], sheet=f'12. {TRUNCATED}')

CAN_TYPES = PRE + "typedef int Result_t; typedef struct { u1 dch_mode:1; } CanChg_t;"
CAN_ROWS = [('can_chg_p->dch_mode', 'u1', '0~1')]
CAN_DECL = "Result_t Motor_Run(void);\nextern CanChg_t *can_chg_p;"
case("can_chg_p: 동등한 반환 typedef와 함수 정의 연결",
     files(CAN_TYPES, CAN_DECL, '#include "g.h"\nint Motor_Run(void) { return 0; }'),
     CAN_ROWS, ['MATCH'])
case("can_chg_p: 실제로 충돌하는 함수 반환 타입의 후보 표시",
     files(CAN_TYPES, CAN_DECL, '#include "g.h"\nlong Motor_Run(void) { return 0; }'),
     CAN_ROWS, ['REVIEW'], reason_contains={0: "함수 'Motor_Run'의 선언과 정의를 연결하지 못했습니다."})
CAN_DUPLICATE = files(CAN_TYPES, 'extern CanChg_t *can_chg_p;',
                      '#include "g.h"\nvoid Motor_Run(void) {}',
                      extra={'legacy.c': 'static void Motor_Run(void) {}'})
case("can_chg_p: 파일 근거로 동명 함수 구분", CAN_DUPLICATE, CAN_ROWS, ['MATCH'])
case("can_chg_p: 파일 근거 없이 동명 함수가 여럿이면 REVIEW",
     CAN_DUPLICATE, CAN_ROWS, ['REVIEW'], xlsx='Detail.xlsx',
     reason_contains={0: 'legacy.c:1 (정의)'})

SOC_TYPES = PRE + "typedef struct { u1 mode:3; } soc_v_t;"
SOC_ROWS = [('soc_v.mode', 'u1', '0~7')]
case("soc_v: 헤더의 extern과 C 정의 연결, 다른 파일의 동명 변수 제외",
     files(SOC_TYPES, 'extern soc_v_t soc_v;', '#include "g.h"\nsoc_v_t soc_v;\nvoid Motor_Run(void) {}\n',
           extra={'legacy.c': 'static struct { unsigned char mode:2; } soc_v;'}),
     SOC_ROWS, ['MATCH'])

LONG_PREFIX = ''.join(f'static void filler_{i}(void) {{\n    unsigned value = {i};\n    (void)value;\n}}\n'
                      for i in range(850))
LONG_BODY = '    step += 1;\n' * 3100
LONG_C = ('#include "g.h"\n' + LONG_PREFIX +
          'void Motor_Run(Rec_t *rec) {\n    unsigned step = 0;\n' + LONG_BODY +
          '    static Rec_t late_local;\n    (void)step;\n}\n')
LATE_LINE = next(i for i, line in enumerate(LONG_C.splitlines(), 1) if 'static Rec_t late_local' in line)
case("6,500줄 C 파일: 3,000줄 이후 함수 매개변수와 함수 내부의 뒤쪽 지역 선언",
     files(REC, '', LONG_C),
     [('rec->items[i].flag', 'u1', '0~1'), ('late_local.status.bit', 'u1', '0~1')],
     ['MATCH', 'MATCH'], reason_contains={1: f'C 선언 motor.c:{LATE_LINE}'})
# --- 1번: 읽기 ----------------------------------------------------------------------------------
case("표지 시트의 멀리 떨어진 셀 (예전: 파일 READ_ERROR)", files(), ROWS, M3, stray=(5000, 60))
case("함수 시트 120장 × 800행 (예전: 셀 예산 초과)", files(), [("g_dev.mode", "u1", "0~7")] * 800, ['MATCH'] * 800 * 120, nsheets=120)


def main():
    passed = failed = 0
    started = time.perf_counter()
    for name, sources, rows, expect, kw in CASES:
        tmp = tempfile.mkdtemp(prefix="rv_selftest_")
        try:
            options = dict(kw)
            reason_contains = options.pop('reason_contains', {})
            paths, xlsx = build(tmp, sources, rows, **options)
            session = run_comparison("", paths, [xlsx], "openpyxl", "utf-8-sig")
            got = [r.status for r in session.results]
            if session.errors:
                got = [f"ERROR:{e.get('stage')}" for e in session.errors]
            ok = got == expect and all(
                part in session.results[index].reason for index, part in reason_contains.items())
            shown = got if len(got) <= 6 else f"{len(got)}건 {sorted(set(got))}"
            print(f"{'PASS' if ok else 'FAIL'}  {name:46} {shown}")
            if not ok:
                for r in session.results[:3]:
                    print(f"        {r.cell} {r.status} | {r.reason[:140]}")
            passed += ok
            failed += not ok
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n{passed} PASS / {failed} FAIL  ({time.perf_counter() - started:.1f}s)  patch={rangevalidator_patch.PATCH_VERSION}")
    return 1 if failed else 0


def run_gui_smoke(receipt):
    """Verify the release GUI through a mixed-result comparison and a failed job."""
    import json
    import tkinter as tk
    from gui_release import App

    root = tk.Tk()
    root.withdraw()
    try:
        app = App(root)
        assert [app.notebook.tab(tab, 'text') for tab in app.notebook.tabs()] == [
            '파일 등록', '비교 결과', '수정 및 저장']
        assert not any(isinstance(widget, tk.ttk.Button) and widget.cget('text') == '진단 로그 보기'
                       for widget in app.progress.master.master.winfo_children())

        def wait_for_result():
            deadline = time.monotonic() + 120
            while app.busy and time.monotonic() < deadline:
                root.update()
                time.sleep(.02)
            assert not app.busy, 'GUI worker did not finish'
            assert app.elapsed_text.get().startswith('소요 시간 ')
            assert not any(isinstance(widget, tk.Toplevel) for widget in root.winfo_children())

        with tempfile.TemporaryDirectory(prefix='rv_gui_smoke_') as tmp:
            source, good = build(tmp, files(), ROWS)
            bad = Path(tmp) / 'broken.xlsx'
            bad.write_bytes(b'Synthetic corrupt workbook')
            app.headers, app.details = source, [good, str(bad)]
            app.compare()
            assert app.elapsed_text.get().startswith('경과 시간 ')
            wait_for_result()
            assert app.notebook.select() == str(app.tabs[1])
            assert [file.status == 'ERROR' for file in app.session.files] == [False, True]
            assert [row.status for row in app.session.results] == M3
            assert '오류 파일 1개' in app.status.get()
            completed_time = app.elapsed_text.get()
            root.update()
            assert app.elapsed_text.get() == completed_time

            app.headers, app.details = [*source, good], [good]
            app.compare()
            wait_for_result()
            assert '비교를 완료하지 못했습니다' in app.status.get()

        Path(receipt).write_text(json.dumps({
            'tabs': [app.notebook.tab(tab, 'text') for tab in app.notebook.tabs()],
            'log_button_removed': True, 'log_popup_removed': True,
            'comparison_elapsed': completed_time, 'failed_comparison_elapsed': app.elapsed_text.get(),
            'mixed_batch_preserved': True,
        }, ensure_ascii=False, indent=2), encoding='utf-8')
    finally:
        root.destroy()


if __name__ == "__main__":
    raise SystemExit(main())
