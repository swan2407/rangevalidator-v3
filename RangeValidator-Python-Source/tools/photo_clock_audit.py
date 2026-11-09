"""Reproduce the four clock rows shown in the user's 2026-11-09 photos.

The struct below is transcribed from visible lines 2334-2352, not a generated substitute.
The photos do not show uint's definition, surrounding directives, variable declaration,
sheet header, or the full existing range. Those are explicitly separated as scenarios.
"""
from dataclasses import asdict
import argparse
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

PHOTO_STRUCT = '''typedef struct {
    uint hu_clock_crc1val:16;         //2 byte
    uint hu_clock_alvcnt1val:8;
    uint rsv00:8;                    //4byte
    uint rsv04:16;                   //6byte
    uint rsv06:16;                   //8byte
    uint rsv08:6;
    uint hu_timeformat:2;
    uint hu_displayhour:8; //10byte
    uint hu_displayminute:8;
    uint rsv24:8;//12byte
    uint rsv32:2;
    uint hu_displaymonth:4;
    uint hu_displayyear1:2;
    uint hu_displayyear2:8;//14byte
    uint hu_displayday:8;
    uint hu_am_pm_info:2;
    uint rsv122:6;//16byte
} can_hu_clock_01_t; /* BMSCR-85 */
'''
PHOTO_ROWS = [('hu_displayday', 8, '0~255', 31, 2349),
              ('hu_displaymonth', 4, '0~15', 32, 2346),
              ('hu_displayyear1', 2, '0~3', 33, 2347),
              ('hu_displayyear2', 8, '0~255', 34, 2348)]
SCENARIOS = ('photo_only', 'known_uint_no_variable', 'conditional_variable',
             'conditional_structure', 'conditional_include', 'conditional_function',
             'conditional_uint', 'conflicting_structure')


def make_case(root, scenario):
    from openpyxl import Workbook
    root.mkdir(parents=True, exist_ok=True)
    preamble = '' if scenario == 'photo_only' else 'typedef unsigned int uint;\n'
    if scenario == 'conditional_uint':
        preamble = '#if FEATURE\ntypedef unsigned int uint;\n#endif\n'
    prefix = preamble + '\n' * (2333 - preamble.count('\n'))
    if scenario in ('conditional_structure', 'conflicting_structure'):
        prefix = prefix[:-1] + '#if FEATURE\n'
    header_text = prefix + PHOTO_STRUCT
    if scenario == 'conditional_structure':
        header_text += '#endif\n'
    elif scenario == 'conflicting_structure':
        header_text += '#else\n' + PHOTO_STRUCT.replace('hu_displayday:8', 'hu_displayday:4') + '#endif\n'
    declaration = 'extern can_hu_clock_01_t can_hu_clock_01;\n'
    if scenario == 'conditional_variable':
        header_text += '#if FEATURE\n' + declaration + '#endif\n'
    elif scenario not in ('photo_only', 'known_uint_no_variable'):
        header_text += declaration
    header = root / 'canfd_comm.h'
    header.write_text(header_text, encoding='utf-8')
    source_text = '#include "canfd_comm.h"\nvoid Update(void) {}\n'
    if scenario == 'conditional_include':
        source_text = '#if FEATURE\n#include "canfd_comm.h"\n#endif\nvoid Update(void) {}\n'
    elif scenario == 'conditional_function':
        source_text = '#include "canfd_comm.h"\n#if FEATURE\nvoid Update(void) {}\n#endif\n'
    source = root / 'photo_clock.c'
    source.write_text(source_text, encoding='utf-8')
    book = Workbook()
    sheet = book.active
    sheet.title = '1.Update'
    for column, value in enumerate(('Input/Output', 'Property', 'Name', 'Data Type', 'Scope'), 4):
        sheet.cell(30, column, value)
    for member, _, _, row, _ in PHOTO_ROWS:
        sheet.cell(row, 4, 'Input')
        sheet.cell(row, 5, 'Global')
        sheet.cell(row, 6, f'can_hu_clock_01.{member}')
        sheet.cell(row, 7, 'uint')
        # The screenshot only displays "0~4294...". This full value is a declared test input.
        sheet.cell(row, 8, '0~4294967295')
    path = root / 'photo_clock.xlsx'
    book.save(path)
    book.close()
    return [header, source], path


def audit(root, phase, backend='openpyxl'):
    from rangevalidator.engine import catalog_headers
    from rangevalidator.workflow import run_comparison
    output = dict(phase=phase, backend=backend,
        evidence='Struct and four member widths transcribed from photos; surrounding conditions are scenarios.',
        assumptions=['uint declaration is not visible; scenarios separate absent/unsigned/conditional definitions.',
                     'Actual surrounding #if and variable declaration are not visible.',
                     'Workbook headers and 0~4294967295 are reproduction inputs, not fully visible photo contents.'],
        cases={})
    for scenario in SCENARIOS:
        paths, path = make_case(root / scenario, scenario)
        catalog = catalog_headers(paths).bitfields
        binding = catalog.lookup_variable(paths[1], 'can_hu_clock_01')
        function = catalog.lookup_function(paths[1], 'Update')
        structs = catalog.type_candidates.get('can_hu_clock_01_t', ())
        original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        session = run_comparison(None, paths, [path], backend=backend)
        output['cases'][scenario] = dict(function_status=function.status,
            binding=asdict(binding), structures=[asdict(d) for d in structs],
            workbook_hash_before=original_hash,
            workbook_hash_after=hashlib.sha256(path.read_bytes()).hexdigest(),
            errors=session.errors, rows=[asdict(r) for r in session.results])
        print(scenario, [(r.cell, r.status, r.expected) for r in session.results])
    root.mkdir(parents=True, exist_ok=True)
    target = root / f'{phase}-{backend}.json'
    target.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding='utf-8')
    return output


def validate(output):
    """Expectations come from the photographed literal widths, not catalog output."""
    for scenario, case in output['cases'].items():
        assert not case['errors'], (scenario, case['errors'])
        assert case['workbook_hash_before'] == case['workbook_hash_after'], scenario
        assert len(case['rows']) == 4, (scenario, case['rows'])
        for result, (member, _, expected, row, _) in zip(case['rows'], PHOTO_ROWS):
            assert result['expression'] == f'can_hu_clock_01.{member}', (scenario, result)
            assert result['cell'] == f'H{row}', (scenario, result)
            conflict = scenario == 'conflicting_structure' and member == 'hu_displayday'
            assert (result['status'], result['expected'], result['candidate']) == (
                ('REVIEW', '', False) if conflict else ('MISMATCH', expected, True)), (scenario, result)
    print(f"Verified {len(output['cases'])} scenarios, 32 rows, original workbooks unchanged.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--phase', required=True)
    parser.add_argument('--backend', default='openpyxl', choices=('openpyxl', 'com'))
    parser.add_argument('--validate', action='store_true')
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1] / 'build/photo-clock-audit')
    args = parser.parse_args()
    output = audit(args.root, args.phase, args.backend)
    if args.validate:
        validate(output)
