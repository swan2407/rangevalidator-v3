"""Verify the release's macro and function linkage fixes inside the frozen EXE."""
import hashlib
import json
from pathlib import Path
import tempfile
import traceback


def run(receipt):
    from openpyxl import Workbook
    import rangevalidator_patch
    from .build_options import load_build_options
    from .workflow import run_comparison

    result = {'successful': False, 'version': rangevalidator_patch.PATCH_VERSION}
    try:
        with tempfile.TemporaryDirectory(prefix='rv_fix9_') as temporary:
            root = Path(temporary)
            header = root / 'types.h'
            header.write_text('''#ifndef TYPES_H
#define TYPES_H
#define NO 0U
#define YES 1U
#define SET 1U
typedef unsigned int uint;
typedef struct { uint BpsCheckSum:8; uint BpsRollingCounter:8; } BpsDataStateCycle_t;
typedef struct { uint bit:3; uint tickCount; } RficRecoveryContext_t;
typedef void (*StateFuncPtr_t)(RficRecoveryContext_t*);
typedef struct { uint welded:2; uint flag:1; } relay_t;
#endif
''', encoding='utf-8')
            bps = root / 'pressure_sensor.c'
            bps.write_text('#include "types.h"\nBpsDataStateCycle_t bps_data;\n'
                'void pressure_diagnosis_process_can_data(const BpsDataStateCycle_t* bps_data) {}', encoding='utf-8')
            recovery = root / 'rf_recovery.c'
            signature = 'static void StateTransition(RficRecoveryContext_t* ctx, StateFuncPtr_t newFuncPtr)'
            recovery.write_text('#include "types.h"\nRficRecoveryContext_t *newFuncPtr;\n' + signature + ';\n' + signature + ' {}', encoding='utf-8')
            relay = root / 'relay.c'
            relay.write_text('#include "types.h"\nrelay_t relay_diag;\nvoid Main(void) {}', encoding='utf-8')

            def book(filename, title, entries):
                path = root / filename
                w = Workbook(); sheet = w.active; sheet.title = title
                sheet.append(['Name', 'Type', 'Range'])
                for name, value in entries:
                    sheet.append([name, '', value])
                w.save(path); w.close()
                return path

            books = [book('unrelated_bps_design.xlsx', ('1.pressure_diagnosis_process_can_data')[:31], [
                ('bps_data->BpsCheckSum', '0~255'), ('bps_data->BpsRollingCounter', '0~4294967295')]),
                book('unrelated_recovery_design.xlsx', '1.StateTransition', [
                    ('ctx->bit', '0~255'), ('ctx->tickCount', '0~4294967295'), ('newFuncPtr->bit', '0~7')]),
                book('relay.xlsx', '1.Main', [
                    ('relay_diag.welded', 'NO'), ('relay_diag.welded', 'yes'),
                    ('relay_diag.welded', 'set'), ('relay_diag.flag', 'NO~YES')])]
            before = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in [header, bps, recovery, relay, *books]}
            options_path = root / 'build.xml'
            options_path.write_text('<cdtprojectproperties><section name="settings.Macros">'
                '<language name="C Source File"/></section></cdtprojectproperties>', encoding='utf-8')
            checks = []
            for label, options in [('without_xml', None), ('with_xml', load_build_options(options_path))]:
                session = run_comparison(None, [header, bps, recovery, relay], books,
                                         backend='openpyxl', build_options=options)
                assert not session.errors, session.errors
                rows = session.results
                actual = [(r.status, r.expected) for r in rows]
                expected = [('MATCH', '0~255'), ('MISMATCH', '0~255'), ('MISMATCH', '0~7'),
                    ('SKIP', ''), ('REVIEW', ''), ('MISMATCH', '0~3'), ('MISMATCH', '0~3'),
                    ('MISMATCH', '0~3'), ('MATCH', '0~1')]
                assert actual == expected, [(r.expression, r.status, r.expected, r.reason) for r in rows]
                assert '일반 멤버' in rows[3].reason
                assert "'NO' → 0" in rows[5].reason
                assert "'yes' → 1" in rows[6].reason
                assert "'set' → 1" in rows[7].reason
                assert not rows[4].candidate
                checks.append({'context': label, 'rows': actual})
            assert before == {p: hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in before}
            result.update(successful=True, checks=checks, originals_unchanged=True)
    except Exception:
        result['error'] = traceback.format_exc()
        raise
    finally:
        Path(receipt).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
