"""RangeValidator-DirectCompare 수정 패치 (v1)

원본 소스 없이 배포본(.pyc)의 동작을 분석해 만든 **함수 단위 교체 패치**입니다.
`import rangevalidator_patch` 한 줄로 아래 모듈의 함수/상수가 교체됩니다.
각 교체 함수의 docstring 에 원본과 달라진 점을 적어 두었으니, 원본 소스에 옮길 때 그 부분만 반영하면 됩니다.

교체 대상
  rangevalidator.models.ReaderLimits          셀 예산 기본값 (1번 읽기 오류)
  rangevalidator.readers.read_openpyxl        함수 시트만 읽기 · 실제 있는 셀만 수집 · 예산을 '실제 셀 수'로 (1번)
  rangevalidator.com_values._last_cell/_merge_map/MAX_COLS   COM 경로 한도 완화 (1번, DRM 파일)
  rangevalidator.bitfield_catalog._prepare_unit              #if 0 · 가지별 중괄호 불균형 · __attribute__ 제거 · 함수형 매크로 수집 (3번)
  rangevalidator.bitfield_catalog._expand_declaration_macros 함수형 매크로 전개 + AUTOSAR Compiler 추상화 매크로 내장 (3번)
  rangevalidator.bitfield_catalog._collect_prepared          union 정의 수집, enum 상수 → 비트폭 계산용 값 등록, enum typedef 등록 (2번)
  rangevalidator.bitfield_catalog._Registry.resolve          enum 타입 비트필드 · 관용 정수 타입(uint8 등) 폴백 (2번)
  rangevalidator.bitfield_catalog.build_bitfield_catalog     멤버 타입 표(member_types) 추가 → 중첩 경로 지원 기반
  rangevalidator.bitfield_catalog.BitfieldCatalog.lookup_function  파일 힌트가 틀려도 유일한 정의면 사용 · 시트명에서 함수명 추출 보강
  rangevalidator.engine.MEMBER_ACCESS / _match_binding / _member_catalog_match   중첩 멤버 경로(a.b.c, a[i].b.c, p->b.c) 지원
  rangevalidator.engine._catalog_match        변수 범위 파일: 함수 정의 파일 > xlsx 이름 추정 (다른 모듈 변수와 비교하던 위험 제거)
  rangevalidator.engine._input_source         시트마다 소스 전체를 훑던 추정 결과 캐시 (성능)
"""
from __future__ import annotations

import re
from collections import defaultdict
from copy import copy
from dataclasses import replace
from pathlib import Path

from rangevalidator import analysis, bitfield_catalog as bc, engine as eng, models, readers as rd, com_values as cv
from rangevalidator.analysis import (BASE_PROFILES, CONVENTIONAL_PROFILES, Diagnostic, Token, canonical_keyword,
                                     identifier, matching_names, tokenize)
from rangevalidator.bitfield_catalog import Origin, TypeLink, _Declaration, _Unit, _CMT_STR, _DIRECTIVE, _INCLUDE, _guard

PATCH_VERSION = "DirectCompare-fix-2"

# =============================================================================
# 1번. Excel 읽기 오류 — 셀 예산 · 메모리
# =============================================================================
# 원본: ReaderLimits(max_sheets=512, max_cells=250000, max_text=16000000)
#   - openpyxl 경로: 모든 시트(표지·목차 포함)의 "내용 사각형(min~max 행×열)" 넓이를 파일 전체로 누적 → 25만 초과 시
#     ValueError → ReadFailure('openpyxl.extract') → COM 폴백 없이 파일 전체 READ_ERROR.
#   - COM 경로(DRM): 함수 시트별 사각형을 누적, 남은 예산(remaining)을 넘는 시트부터 SHEET_READ_ERROR → 뒤쪽 시트는
#     데이터가 적어도 오류. 시트 1장 병합 스캔 예산(200000)도 사각형 기준.
# 패치: 예산은 '실제 값이 있는 셀 수' 기준, 기본값 5,000,000. 사각형 넓이는 더 이상 예산에 넣지 않음.
_ORIGINAL_READER_LIMITS = (models.ReaderLimits.max_sheets, models.ReaderLimits.max_cells, models.ReaderLimits.max_text)
models.ReaderLimits.max_cells = 5_000_000
models.ReaderLimits.__init__.__defaults__ = (models.ReaderLimits.max_sheets, models.ReaderLimits.max_cells, models.ReaderLimits.max_text)


def _refresh_default_limits(*functions):
    """readers/com_values 의 함수들은 `limits=ReaderLimits()` 를 **import 시점의 인스턴스**로 기본값에 담고 있어
    클래스 기본값만 바꿔서는 효과가 없다 → 기본값 튜플 안의 ReaderLimits 인스턴스를 새 값으로 바꾼다."""
    for function in functions:
        defaults = getattr(function, '__defaults__', None)
        if defaults:
            function.__defaults__ = tuple(models.ReaderLimits() if isinstance(d, models.ReaderLimits) else d for d in defaults)


_refresh_default_limits(rd.read_auto, rd.read_com, rd.read_workbook, cv.read_workbook_values)


def read_openpyxl(path, limits=None):
    """원본 readers.read_openpyxl 교체.

    달라진 점
      1. 비교 대상(함수 시트: 이름에 숫자가 있는 시트)만 셀을 읽는다. 다른 시트는 이름·상태만 가진 빈 Sheet 로 남긴다
         (compare_book 은 함수 시트만 보고, writer 는 COM 으로 원본을 직접 열므로 영향 없음).
      2. 사각형(iter_rows)으로 훑지 않고 실제 존재하는 셀(ws._cells)만 모은다 → 멀리 떨어진 셀 하나 때문에
         수십만 개의 빈 Cell 을 만들지 않는다. 값이 None 인 셀은 병합 영역 안쪽 셀만 보관(COM 경로와 동일).
      3. 셀 예산(limits.max_cells)은 '값이 있는 셀 수' 기준으로 센다.
      4. 값(캐시) 통합문서는 read_only 스트리밍으로 열고 수식 셀의 캐시 값만 가져온다 → 메모리 절반.
         (원본은 formula/value 두 통합문서를 모두 전체 로드: 셀당 약 400B × 2)
    결과 구조(Workbook/Sheet/Cell, bounds, merged_ranges, read_timings)는 원본과 같다.
    """
    from time import perf_counter
    from datetime import date, datetime, time, timedelta
    limits = limits or models.ReaderLimits()
    stage, formula_book, value_book = 'path', None, None
    cleanup, result, failure = [], None, None
    try:
        rd.check_limits(limits)
        p = rd.input_path(path)
        stage = 'dependency.openpyxl'
        from openpyxl import load_workbook
        from openpyxl.utils.datetime import to_excel, time_to_days, timedelta_to_days
        stage = 'openpyxl.open'
        tick = perf_counter()
        formula_book = load_workbook(p, read_only=False, data_only=False, keep_links=False)
        open_seconds = perf_counter() - tick
        stage = 'openpyxl.extract'
        if len(formula_book.worksheets) > limits.max_sheets:
            raise ValueError('Sheet budget exceeded')
        result = models.Workbook(str(p), 'openpyxl')
        result.source_sheet_count = len(formula_book.worksheets)
        result.read_timings['open_seconds'] = open_seconds
        tick = perf_counter()
        total = text = 0
        formula_cells = {}          # 시트명 → {(row, col): Cell}  수식 셀 (캐시 값은 뒤에서 채움)
        for ws in formula_book.worksheets:
            ranges = list(ws.merged_cells.ranges)
            sheet = models.Sheet(ws.title, ws.sheet_state, merged_ranges=sorted(map(str, ranges)))
            if not eng.is_function_sheet(ws.title):
                result.sheets.append(sheet)         # 비교 대상 아님: 이름만 보관
                continue
            active = [c for c in ws._cells.values() if c.value is not None]
            bounds = rd.bounds_for(active)
            sheet.bounds = bounds
            if bounds:
                total += len(active)
                if total > limits.max_cells:
                    raise ValueError('Content cell budget exceeded')
                wanted = {}
                for c in ws._cells.values():
                    region, anchor = rd._merge(c, ranges)
                    if c.value is None and not region:
                        continue
                    is_formula = c.data_type == 'f'
                    formula = c.value if is_formula and isinstance(c.value, str) else None
                    if is_formula and formula is None:
                        formula = getattr(c.value, 'text', None) or '<non-scalar formula>'
                    value = None if is_formula else c.value
                    if isinstance(value, (datetime, date)):
                        value = to_excel(value, formula_book.epoch)
                    elif isinstance(value, time):
                        value = time_to_days(value)
                    elif isinstance(value, timedelta):
                        value = timedelta_to_days(value)
                    kind = 'error' if c.data_type == 'e' else ('formula' if is_formula else 'value')
                    cell = models.Cell(c.row, c.column, c.coordinate, value, is_formula, formula, region, anchor, kind)
                    if isinstance(value, str):
                        text += len(value)
                    if formula:
                        text += len(formula)
                    if text > limits.max_text:
                        raise ValueError('Text budget exceeded')
                    sheet.cells[cell.address] = cell
                    if is_formula:
                        wanted[(c.row, c.column)] = (sheet, cell.address)
                if wanted:
                    formula_cells[ws.title] = wanted
            result.sheets.append(sheet)
        if formula_cells:                        # 수식 셀의 캐시 값만 read_only 스트리밍으로 보충
            value_book = load_workbook(p, read_only=True, data_only=True, keep_links=False)
            for title, wanted in formula_cells.items():
                vs = value_book[title]
                rows = sorted({r for r, _ in wanted})
                for row in vs.iter_rows(min_row=rows[0], max_row=rows[-1]):
                    for vc in row:
                        hit = wanted.get((getattr(vc, 'row', None), getattr(vc, 'column', None)))
                        if hit is None:
                            continue
                        value = vc.value
                        if isinstance(value, (datetime, date)):
                            value = to_excel(value, formula_book.epoch)
                        elif isinstance(value, time):
                            value = time_to_days(value)
                        elif isinstance(value, timedelta):
                            value = timedelta_to_days(value)
                        target_sheet, address = hit
                        target_sheet.cells[address] = replace(target_sheet.cells[address], value=value)
                        if isinstance(value, str):
                            text += len(value)
        result.read_timings['body_read_seconds'] = perf_counter() - tick
    except Exception as exc:                     # noqa: BLE001 — 원본과 같이 ReadFailure 로 감싼다
        failure = models.ReadFailure(stage, exc)
    finally:
        for label, book in (('values.close', value_book), ('formulas.close', formula_book)):
            if book is not None:
                try:
                    book.close()
                except Exception as exc:        # noqa: BLE001
                    cleanup.append(models.ReadFailure(label, exc).as_dict())
    if failure or cleanup:
        if not failure:
            failure = models.ReadFailure('cleanup', RuntimeError('Workbook cleanup failed'))
        failure.cleanup_errors.extend(cleanup)
        raise failure
    return result


rd.read_openpyxl = read_openpyxl

# COM 경로(DRM 파일) ----------------------------------------------------------
# 원본: _last_cell(strict=True) 가 마지막 열 > 256 이면 'Content extends beyond column 256' 로 시트 읽기 실패,
#       _merge_map 은 (행×열) > 200000 이면 'Merge scan budget' 으로 시트 읽기 실패,
#       _read_blocks 는 (데이터 행×열) 가 남은 파일 예산(limits.max_cells − 누적)을 넘으면 실패.
# 패치: 열 한도 1024 로 올리고 넘는 열은 잘라 읽음(실패 대신), 병합 스캔 예산 500만, 파일 예산은 ReaderLimits 로 500만.
cv.MAX_COLS = 1024
_original_last_cell = cv._last_cell
_original_merge_map = cv._merge_map


def _last_cell(sheet, diag, strict=False):
    """원본 _last_cell: strict=True 호출(→ _copy_sheet) 도 비-엄격으로 바꿔, 열이 MAX_COLS 를 넘으면 실패 대신 MAX_COLS 까지만 읽는다."""
    row, col = _original_last_cell(sheet, diag, strict=False)
    return row, min(col, cv.MAX_COLS)


def _merge_map(ws, last_row, last_col, diag, max_cells=200000):
    """원본 _merge_map 의 병합 스캔 예산(200000 → 5,000,000)."""
    return _original_merge_map(ws, last_row, last_col, diag, max_cells=max(max_cells, 5_000_000))


_original_copy_sheet = cv._copy_sheet


def _copy_sheet(ws, limits, remaining=None, diag=None, merge_limit=200000):
    """원본 _copy_sheet 를 호출하되, 통합문서 누적 예산에 더하는 '셀 수' 를 (데이터 행×열 사각형) 대신 **실제 저장한 셀 수**로
    돌려준다 → 멀리 떨어진 셀 하나가 있는 시트가 뒤쪽 시트들의 예산을 잠식하지 않는다."""
    sheet, _cells, chars, errors = _original_copy_sheet(ws, limits, remaining, diag, merge_limit)
    return sheet, len(sheet.cells), chars, errors


cv._last_cell = _last_cell
cv._merge_map = _merge_map
cv._copy_sheet = _copy_sheet

# =============================================================================
# 3번. 선언 수집 — _prepare_unit / 매크로 전개
# =============================================================================
# GCC/MSVC/컴파일러 확장 키워드: 선언자 앞뒤에 붙어 선언 수집을 막는다. 텍스트 단계에서 제거(줄 수 유지).
_ATTRIBUTE_WORDS = ('__attribute__', '__attribute', '__declspec', '__asm__', '__asm', '_Alignas', '__alignof__')
_BARE_WORDS = frozenset({'__packed', '__extension__', '__inline', '__inline__', '__forceinline', '__near', '__far',
                         '__huge', '__sfr', '__bit', '__xdata', '__pdata', '__idata', '__data', '__code', '__interrupt',
                         '__cdecl', '__stdcall', '__fastcall', '__restrict', '__restrict__'})
_WORD = re.compile(r'[A-Za-z_]\w*')


_EXT_WORD = re.compile(r'\b(' + '|'.join(map(re.escape, _ATTRIBUTE_WORDS)) + r')\b')
_BARE_WORD = re.compile(r'\b(' + '|'.join(map(re.escape, sorted(_BARE_WORDS))) + r')\b')


def _strip_extensions(source):
    """`__attribute__((...))`, `__declspec(...)`, `__asm__(...)` 와 `__packed` 같은 확장 키워드를 공백으로 바꾼다.

    줄바꿈은 그대로 두어 줄 번호가 유지된다. 키워드가 없는 파일은 그대로 돌려준다(대용량 소스에서 비용 0에 가깝게).
    """
    if not _EXT_WORD.search(source) and not _BARE_WORD.search(source):
        return source
    out, pos, n = [], 0, len(source)
    for m in _EXT_WORD.finditer(source):
        if m.start() < pos:
            continue
        out.append(source[pos:m.start()])
        j = m.end()
        while j < n and source[j].isspace():
            j += 1
        end = m.end()
        if j < n and source[j] == '(':
            depth, k = 0, j
            while k < n:
                if source[k] == '(':
                    depth += 1
                elif source[k] == ')':
                    depth -= 1
                    if depth == 0:
                        break
                k += 1
            if depth == 0:
                end = k + 1
        removed = source[m.start():end]
        out.append(' ' + '\n' * removed.count('\n'))
        pos = end
    out.append(source[pos:])
    return _BARE_WORD.sub(' ', ''.join(out))


def _brace_delta(lines, start, stop):
    """lines[start:stop] 의 중괄호 순증가량. 리터럴 안의 중괄호는 세지 않는다."""
    delta = 0
    for line in lines[start:stop]:
        i, n = 0, len(line)
        while i < n:
            c = line[i]
            if c in '"\'':
                i += 1
                while i < n and line[i] != c:
                    i += 2 if line[i] == '\\' else 1
            elif c == '{':
                delta += 1
            elif c == '}':
                delta -= 1
            i += 1
    return delta


def _prune_conditional_branches(lines, directives):
    """`#if 0` 가지를 비우고, 한 조건 그룹의 어느 가지든 중괄호를 열거나 닫으면(함수 머리/블록이 가지별로 다른 '구조적' 그룹)
    첫 가지만 남기고 나머지 가지의 코드 줄을 비운다. 가지 전부가 중괄호 균형 0 인 보통의 선언 그룹은 그대로 둔다(후보 비교 유지).

    원본은 모든 가지를 그대로 파싱해, 한 가지에만 `{` 가 있으면 그 뒤 파일 전체가 함수 본문으로 취급되어
    선언이 수집되지 않았다(DECL_NOT_COLLECTED). 지시문 줄 자체는 남겨 둔다(조건부 표시에 쓰임).
    반환: (지워진 줄 번호 집합, #else 가지가 '항상 참' 이 된 #if 0 그룹의 지시문 줄 번호 집합)
    """
    groups, stack = [], []
    for line, command, body in directives:
        if command in ('if', 'ifdef', 'ifndef'):
            stack.append({'start': line, 'marks': [(line, command, body)], 'end': None})
        elif command in ('elif', 'else') and stack:
            stack[-1]['marks'].append((line, command, body))
        elif command == 'endif' and stack:
            group = stack.pop()
            group['end'] = line
            groups.append(group)
    cleared, always_else = set(), set()

    def clear(start, stop):
        cleared.update(range(start, stop))
        for ln in range(start, stop):
            if lines[ln - 1].strip() and not _DIRECTIVE.fullmatch(lines[ln - 1]):
                lines[ln - 1] = ''

    for group in sorted(groups, key=lambda g: g['end'] - g['start']):        # 안쪽 그룹부터
        marks = group['marks'] + [(group['end'], 'endif', '')]
        branches = [(marks[k][0] + 1, marks[k + 1][0]) for k in range(len(marks) - 1)]   # 1-based [start, stop)
        first_cmd, first_body = marks[0][1], marks[0][2].strip()
        dead_first = first_cmd == 'if' and re.fullmatch(r'\(*\s*0+[uUlL]*\s*\)*', first_body) is not None
        live = branches
        if dead_first:
            clear(*branches[0])                                   # #if 0 가지는 컴파일되지 않는 코드
            live = branches[1:]
            if len(branches) == 2 and marks[1][1] == 'else':
                always_else.add(group['start'])
        if len(live) > 1:
            deltas = [_brace_delta(lines, b[0] - 1, b[1] - 1) for b in live]
            if any(deltas):                                       # 어느 가지든 중괄호를 열거나 닫으면(구조적 가지) → 첫 가지만
                for start, stop in live[1:]:
                    clear(start, stop)
    return cleared, always_else


def _function_macro(expr):
    """`#define NAME(a, b) body` 의 `(a, b) body` 부분 → ('<fn>', params, body_words) / 형식이 아니면 None."""
    m = re.match(r'\(\s*([^)]*)\)\s*(.*)$', expr, re.S)
    if not m:
        return None
    params = tuple(p.strip() for p in m.group(1).split(',') if p.strip())
    if any(not identifier(p) for p in params):
        return None
    body = m.group(2).strip()
    if '#' in body:                           # 문자열화 / 토큰 붙이기 매크로는 전개하지 않음
        return None
    words = tuple(t.text for t in tokenize(body, '<macro>', []))
    return ('<fn>', params, words)


class _MacroEvents(list):
    def __init__(self, guard, once):
        super().__init__()
        self.guard = guard
        self.once = once


def _prepare_unit(source, file, files):
    """원본 bitfield_catalog._prepare_unit 교체.

    달라진 점
      1. 주석 제거 뒤 `__attribute__((...))`, `__declspec(...)`, `__packed` 등 컴파일러 확장을 제거(_strip_extensions).
      2. `#if 0` 가지는 버리고, 한 조건 그룹의 가지끼리 중괄호 균형이 다르면 첫 가지만 남긴다(_prune_conditional_branches).
         `#if 0 … #else X #endif` 의 X 는 조건부로 표시하지 않는다(항상 컴파일되는 코드).
      3. 함수형 매크로 `#define NAME(a,b) body` 를 macro_events 에 ('<fn>', params, body) 값으로 기록한다
         (원본은 객체형 매크로 중 기본형 철자만 기록). unit.macros 에는 원본처럼 target='' 로 등록.
    나머지(guard 판정, 조건부 줄 집합, include 처리, 반환값 형식)는 원본과 같다.
    """
    unit = _Unit()
    lines = source.lstrip('﻿').splitlines()
    for i in range(len(lines) - 2, -1, -1):
        if lines[i].endswith('\\'):
            lines[i] = lines[i][:-1] + lines[i + 1]
            lines[i + 1] = ''
    source = _CMT_STR.sub(lambda m: m[0] if m[0][0] in '"\'' else ' ' + '\n' * m[0].count('\n'), '\n'.join(lines))
    source = _strip_extensions(source)
    lines = source.split('\n')
    directives = []
    for i, line in enumerate(lines, 1):
        match = _DIRECTIVE.fullmatch(line)
        if match:
            directives.append((i, match[1].casefold(), match[2]))
    guard = _guard(directives, lines)
    guard_name = None
    if guard is not None:
        guard_index = next(index for index, (line, _, _) in enumerate(directives) if line == guard)
        guard_name = directives[guard_index + 1][2].split()[0]
    discarded, always_else = _prune_conditional_branches(lines, directives)
    macro_events = _MacroEvents(guard_name, any(
        command == 'pragma' and body.strip().casefold() == 'once'
        for _, command, body in directives[:1]))
    conditional, conditions = set(), []
    by_line = {line: (command, body.strip()) for line, command, body in directives}
    dead_groups = []                                   # #if 0 그룹 추적용 스택 (True = 이 그룹은 #if 0)
    for i, line in enumerate(lines, 1):
        if i not in by_line:
            if any(conditions):
                conditional.add(i)
            continue
        command, body = by_line[i]
        lines[i - 1] = ''
        if command in ('if', 'ifdef', 'ifndef'):
            conditions.append(i != guard)
            dead_groups.append(i in always_else)
        elif command in ('else', 'elif') and conditions and dead_groups and dead_groups[-1]:
            conditions[-1] = False                     # #if 0 의 else 가지 = 항상 컴파일
            dead_groups[-1] = False
        elif command == 'endif':
            if conditions:
                conditions.pop()
            if dead_groups:
                dead_groups.pop()
        elif i in discarded:
            continue
        elif command in ('define', 'undef'):
            match = re.match(r'([A-Za-z_]\w*)(.*)', body)
            if match:
                name, expr = match.groups()
                replacement = None
                words = [canonical_keyword(w) for w in expr.split()]
                if command == 'define' and not any(conditions) and not expr.startswith('('):
                    if (name == guard_name or not words or ' '.join(words) in BASE_PROFILES
                            or len(words) == 1 and (identifier(words[0]) or words[0] in
                                                   ('struct', 'union', 'enum', 'const', 'volatile'))):
                        replacement = tuple(words)
                if command == 'define' and expr.startswith('('):
                    replacement = _function_macro(expr)
                macro_events.append((i, name, replacement))
                target = expr.strip() if command == 'define' and not expr.startswith('(') else ''
                unit.macros[name].append(_Declaration(target, Origin(file, i), any(conditions)))
        if command == 'include':
            local = None
            match = _INCLUDE.fullmatch(body)
            if match:
                local = str((Path(file).parent / match[1]).resolve())
                folded = _CASEFOLD_FILES.get(id(files))
                if folded is None or folded[0] is not files:
                    by_name = defaultdict(list)
                    for path in files:
                        by_name[Path(path).name.casefold()].append(path)
                    folded = (files, {p.casefold(): p for p in files}, by_name)
                    _CASEFOLD_FILES.clear()
                    _CASEFOLD_FILES[id(files)] = folded
                local = folded[1].get(local.casefold(), local)
                include = Path(match[1])
                if local not in files and not include.is_absolute() and '..' not in include.parts:
                    suffix = '/' + include.as_posix().casefold()
                    matches = [path for path in folded[2].get(include.name.casefold(), ())
                               if Path(path).as_posix().casefold().endswith(suffix)]
                    if len(matches) == 1:
                        local = matches[0]
                if local in files and unit.includes.get(local, True):
                    unit.includes[local] = any(conditions)
            macro_events.append((i, None, None if any(conditions) else local))
        elif command in ('include_next', 'import') or (
                command == 'pragma' and not re.match(r'^(?:once|section|pack)\b', body, re.IGNORECASE)):
            macro_events.append((i, None, None))
    return unit, '\n'.join(lines), conditional, macro_events


# AUTOSAR Compiler Abstraction (Compiler.h) — 입력 폴더에 Compiler.h 가 없을 때만 쓰는 내장 정의
AUTOSAR_MACROS = {
    'FUNC': (('rettype', 'memclass'), ('rettype',)),
    'FUNC_P2CONST': (('rettype', 'ptrclass', 'memclass'), ('const', 'rettype', '*')),
    'FUNC_P2VAR': (('rettype', 'ptrclass', 'memclass'), ('rettype', '*')),
    'P2VAR': (('ptrtype', 'memclass', 'ptrclass'), ('ptrtype', '*')),
    'P2CONST': (('ptrtype', 'memclass', 'ptrclass'), ('const', 'ptrtype', '*')),
    'CONSTP2VAR': (('ptrtype', 'memclass', 'ptrclass'), ('ptrtype', '*', 'const')),
    'CONSTP2CONST': (('ptrtype', 'memclass', 'ptrclass'), ('const', 'ptrtype', '*', 'const')),
    'P2P2VAR': (('ptrtype', 'memclass', 'ptrclass'), ('ptrtype', '*', '*')),
    'P2P2CONST': (('ptrtype', 'memclass', 'ptrclass'), ('const', 'ptrtype', '*', '*')),
    'P2FUNC': (('rettype', 'ptrclass', 'fctname'), ('rettype', '(', '*', 'fctname', ')')),
    'CONSTP2FUNC': (('rettype', 'ptrclass', 'fctname'), ('rettype', '(', '*', 'const', 'fctname', ')')),
    'CONST': (('type', 'memclass'), ('const', 'type')),
    'VAR': (('type', 'memclass'), ('type',)),
}
_ASSUMED_MACRO_NOTE = {}       # file → set(name)  (내장 정의를 쓴 파일별 매크로 이름, 진단용)
_DEFINED_CACHE = {}            # id(inputs) → (inputs, 입력 전체에서 정의된 매크로 이름 집합)
_CASEFOLD_FILES = {}           # id(files) → (files, {casefold path: path})


def _call_arguments(tokens, start):
    """tokens[start] 가 '(' 일 때 짝 ')' 까지의 인자 목록(토큰 리스트들)과 ')' 위치. 짝이 없으면 (None, start)."""
    depth, args, current = 0, [], []
    for k in range(start, len(tokens)):
        text = tokens[k].text
        if text == '(':
            depth += 1
            if depth == 1:
                continue
        elif text == ')':
            depth -= 1
            if depth == 0:
                args.append(current)
                return [a for a in args if a or len(args) > 1], k
        elif text == ',' and depth == 1:
            args.append(current)
            current = []
            continue
        current.append(tokens[k])
    return None, start


def _expand_declaration_macros(tokens, file, events, inputs):
    """원본 bitfield_catalog._expand_declaration_macros 교체.

    달라진 점: 값이 ('<fn>', params, body) 인 함수형 매크로를 `NAME ( 인자, … )` 꼴에서 매개변수 치환해 전개한다
    (중첩 8단계까지). 입력 파일 어디에도 정의가 없는 AUTOSAR Compiler 추상화 매크로(VAR, CONST, P2VAR, FUNC …)는
    내장 정의(AUTOSAR_MACROS)로 전개한다. 객체형 별칭을 따라가고, 가드된 헤더를 다시 include해도 활성 정의를 유지한다.
    """
    expanded, active, processing = [], {}, {file}
    once_seen = {file} if events.once else set()
    cache = _DEFINED_CACHE.get(id(inputs))
    if cache is None or cache[0] is not inputs:
        defined_anywhere = {name for file_events in inputs.values() for _, name, _value in file_events if name is not None}
        _DEFINED_CACHE.clear()
        _DEFINED_CACHE[id(inputs)] = (inputs, defined_anywhere)
    else:
        defined_anywhere = cache[1]

    def apply(name, value, depth):
        if name is not None:
            if value is None:
                active.pop(name, None)
            else:
                active[name] = value
            return
        if value not in inputs or depth >= 64:
            active.clear()
            return
        included = inputs[value]
        if value in once_seen or included.guard in active:
            return
        if value in processing:
            active.clear()
            return
        if included.once:
            once_seen.add(value)
        processing.add(value)
        try:
            for _, included_name, included_value in included:
                apply(included_name, included_value, depth + 1)
        finally:
            processing.remove(value)

    def expand_tokens(items, depth):
        out, i = [], 0
        while i < len(items):
            token = items[i]
            replacement = active.get(token.text)
            builtin = None
            if replacement is None and token.text in AUTOSAR_MACROS and token.text not in defined_anywhere:
                builtin = AUTOSAR_MACROS[token.text]
            if (isinstance(replacement, tuple) and replacement[:1] == ('<fn>',)) or builtin:
                params, body = (replacement[1], replacement[2]) if builtin is None else builtin
                if i + 1 < len(items) and items[i + 1].text == '(':
                    args, close = _call_arguments(items, i + 1)
                    if args is not None and len(args) == len(params) and depth < 8:
                        mapping = dict(zip(params, args))
                        produced = []
                        for word in body:
                            if word in mapping:
                                produced.extend(Token(t.text, token.line) for t in mapping[word])
                            else:
                                produced.append(Token(word, token.line))
                        if builtin is not None:
                            _ASSUMED_MACRO_NOTE.setdefault(file, set()).add(token.text)
                        out.extend(expand_tokens(produced, depth + 1))
                        i = close + 1
                        continue
                out.append(token)
                i += 1
                continue
            if replacement is not None:
                if len(replacement) == 1 and replacement[0] != token.text and depth < 8:
                    out.extend(expand_tokens([Token(replacement[0], token.line)], depth + 1))
                else:
                    out.extend(Token(text, token.line) for text in replacement)
            else:
                out.append(token)
            i += 1
        return out

    events = iter(events)
    event = next(events, None)
    pending = []
    for token in tokens:
        while event is not None and event[0] < token.line:
            if pending:
                expanded.extend(expand_tokens(pending, 0))
                pending = []
            apply(event[1], event[2], 0)
            event = next(events, None)
        pending.append(token)
    if pending:
        expanded.extend(expand_tokens(pending, 0))
    return expanded


def _collect_enums(tokens, conditional, file, unit):
    """정리된 소스에서 `enum [tag] { A, B = 3, C }` 를 찾아 열거 상수를 비트폭 계산용 값으로 unit.macros 에 등록하고,
    `typedef enum {...} E_t;` / `enum tag` 이름을 unit.types 에 target='enum' 으로 등록한다.

    값은 식 문자열로 저장하고(명시 값은 그 식, 다음 상수는 '(이전상수)+1'), 실제 계산은 _Registry.width 의 매크로
    전개가 맡는다 → 매크로·다른 상수를 참조하는 값도 기존 규칙대로 계산된다.
    """
    i, n = 0, len(tokens)
    while i < n:
        if tokens[i].text != 'enum':
            i += 1
            continue
        line = tokens[i].line
        j = i + 1
        tag = ''
        if j < n and identifier(tokens[j].text):
            tag = tokens[j].text
            j += 1
        if j >= n or tokens[j].text != '{':
            i = j
            continue
        close = analysis._Parser.matching(tokens, j)
        if close < 0:
            break
        uncertain = line in conditional
        previous = None
        for part in analysis._Parser.split(tokens[j + 1:close], ','):
            words = [t.text for t in part]
            if not words or not identifier(words[0]):
                continue
            name = words[0]
            if '=' in words:
                expr = ' '.join(words[words.index('=') + 1:])
            elif previous is None:
                expr = '0'
            else:
                expr = f'({previous}) + 1'
            unit.macros[name].append(_Declaration(expr, Origin(file, part[0].line), uncertain))
            previous = name
        if tag:
            unit.types['enum ' + tag].append(_Declaration('enum', Origin(file, line), uncertain))
        # typedef enum {...} NAME, *PNAME;
        k = close + 1
        typedef = i > 0 and tokens[i - 1].text == 'typedef'
        if typedef:
            names = []
            while k < n and tokens[k].text != ';':
                if identifier(tokens[k].text):
                    names.append(tokens[k].text)
                k += 1
            for name in names:
                if not unit.types.get(name):
                    unit.types[name].append(_Declaration('enum', Origin(file, line), uncertain))
        i = close + 1


class _MemberCollector(bc._Collector):
    """Reuse the C object declarator parser for ordinary struct and union members."""

    def parse_members(self, st, body, kind):
        st.member_declarations = {}
        for tokens in self.split(body, ';'):
            if not tokens:
                continue
            words = [token.text for token in tokens]
            origin = Origin(self.file, tokens[0].line)
            offset = next((i for i, word in enumerate(words) if word not in analysis.QUALIFIERS), len(words))
            if '{' in words and offset < len(words) and words[offset] in ('struct', 'union'):
                inner, end, _tag = self.parse_aggregate(tokens, offset, False, origin.line)
                if inner is None:
                    continue
                self.register_struct(inner)
                if end == len(tokens):
                    st.fields.extend(inner.fields)
                    self.m.fields.extend(inner.fields)
                    st.ordinary.extend(inner.ordinary)
                    st.members.update(inner.members)
                    st.nested.update(inner.nested)
                    st.member_declarations.update(inner.member_declarations)
                    continue
                declarations = tuple(bc._variable_declarations(
                    self, words[:offset], tokens[end:], origin, False, inner.name))
                st.nested.update((d.name, inner) for d in declarations)
            elif ':' not in words:
                declarations = tuple(bc._variable_declarations(self, None, tokens, origin, False))
                if not declarations:
                    super().parse_members(st, tokens, kind)
            else:
                super().parse_members(st, tokens, kind)
                continue
            for declaration in declarations:
                name = declaration.name
                if name in st.members:
                    self.review(origin.line, f"duplicate member '{name}'")
                st.ordinary.append(name)
                st.members[name] = declaration.spelling
                st.member_declarations[name] = declaration


def _collect_prepared_core(prepared, file, inputs):
    """원본 bitfield_catalog._collect_prepared 교체(선언 분류 루프).

    달라진 점: `typedef union {…} U_t;` / `union tag {…} var;` 를 struct 와 똑같이 정의로 수집한다.
    (원본은 kind == 'struct' 만 처리하고 union 몸체가 있는 typedef 는 'Unsupported aggregate typedef' 로 버림 →
     `typedef union { u2 all; struct {…} bit; } Reg_t;` 레지스터 맵 변수가 모두 '선언 연결이 불확실함 (UNRESOLVED)')
    """
    unit, source, conditional, macro_events = prepared
    tokens = tokenize(source, file, unit.reviews)
    if macro_events:
        tokens = bc._expand_declaration_macros(tokens, file, macro_events, inputs)
    unit._patch_tokens = tokens                      # enum 수집에 재사용 (토큰화 2회 방지)
    parser = _MemberCollector(tokens, file, analysis.SourceModel())

    def function(head, body):
        try:
            declaration = bc._collect_function(parser, head, body, conditional)
        except (IndexError, ValueError, RecursionError) as exc:
            unit.reviews.append(Diagnostic(file, head[0].line if head else 1, f'Function skipped: {exc}'))
            return
        if declaration is not None:
            unit.functions.append(declaration)

    for tokens in parser.declarations(function):
        try:
            words = [token.text for token in tokens]
            origin = Origin(file, tokens[0].line)
            uncertain = any(token.line in conditional for token in tokens)
            if uncertain:
                unit.conditional_declarations.add(origin)
            typedef = words[0] == 'typedef'
            offset = next((i for i in range(int(typedef), len(words)) if words[i] not in analysis.QUALIFIERS), len(words))
            kind = words[offset]
            after_kind = offset + 1
            if after_kind < len(words) and identifier(words[after_kind]):
                after_kind += 1
            if kind in ('struct', 'union') and words[after_kind:after_kind + 1] == ['{']:
                nested_start = len(parser.m.types)
                st, end, tag = parser.parse_aggregate(tokens, offset, typedef, origin.line)
                if st is None:
                    continue
                definition = bc._definition_id(Origin(file, st.line), unit.definitions)
                names = []
                if typedef:
                    names = bc._typedefs(parser, tokens[end:], definition, origin, uncertain, unit)
                if names:
                    parser.rename(st, names[0])
                unit.definitions[definition] = (st, uncertain)
                for nested in parser.m.types[nested_start:]:
                    nested_origin = bc._definition_id(Origin(file, nested.line), unit.definitions)
                    unit.definitions[nested_origin] = (nested, uncertain)
                if tag:
                    unit.types[f'{kind} {tag}'].append(_Declaration(definition, origin, uncertain))
                if not typedef:
                    bc._collect_variables(parser, words[:offset], tokens[end:], origin, uncertain, unit,
                                          f'{kind} {tag}' if tag else st.name, definition)
            elif kind in ('struct', 'union') and len(words) == 2 and identifier(words[1]):
                spelling = f'{kind} {words[1]}'
                unit.types[spelling].append(_Declaration(spelling, origin, uncertain))
            elif typedef and '{' in words:
                unit.reviews.append(Diagnostic(file, origin.line, 'Unsupported aggregate typedef'))
            elif typedef:
                bc._collect_typedef(parser, tokens[1:], origin, uncertain, unit)
            else:
                prototype = bc._collect_function(parser, tokens, None, conditional)
                if prototype is not None:
                    unit.functions.append(prototype)
                else:
                    bc._collect_variables(parser, None, tokens, origin, uncertain, unit)
        except (IndexError, ValueError, RecursionError) as exc:
            unit.reviews.append(Diagnostic(file, origin.line, f'Declaration skipped: {exc}'))
    unit.reviews.extend(parser.m.reviews)
    return unit


def _collect_prepared(prepared, file, inputs):
    """_collect_prepared_core 뒤에 enum 수집(_collect_enums)을 덧붙인다. 내장 AUTOSAR 매크로를 쓴 파일에는 INFO 진단을 남긴다."""
    unit = _collect_prepared_core(prepared, file, inputs)
    tokens = unit.__dict__.pop('_patch_tokens', None)
    try:
        if tokens and any(t.text == 'enum' for t in tokens):
            _collect_enums(tokens, prepared[2], file, unit)
    except (IndexError, ValueError, RecursionError) as exc:
        unit.reviews.append(Diagnostic(file, 1, f'enum collection skipped: {exc}'))
    names = _ASSUMED_MACRO_NOTE.pop(file, None)
    if names:
        unit.reviews.append(Diagnostic(file, 1, f"AUTOSAR Compiler 매크로를 내장 정의로 전개함 (Compiler.h 미수집): {', '.join(sorted(names))}",
                                       'AUTOSAR_MACRO_ASSUMED', 'INFO'))
    return unit


bc._prepare_unit = _prepare_unit
bc._expand_declaration_macros = _expand_declaration_macros
bc._collect_prepared = _collect_prepared

# =============================================================================
# 2번. 타입 해석 — enum 비트필드 · 관용 정수 타입 폴백
# =============================================================================
_CANONICAL_INTEGER = {(False, 8): 'unsigned char', (True, 8): 'signed char', (False, 16): 'unsigned short', (True, 16): 'short',
                      (False, 32): 'unsigned int', (True, 32): 'signed int', (False, 64): 'unsigned long long', (True, 64): 'long long',
                      (False, 1): '_Bool'}
_original_resolve = bc._Registry.resolve


def _registry_resolve(self, file, name, root=None, seen=frozenset()):
    """원본 _Registry.resolve 를 감싼다.

    달라진 점
      1. 철자가 'enum' 또는 'enum tag' 이면 → INTEGER(unsigned int) 로 해석 + 가정 메모
         (열거형 비트필드의 부호는 구현 정의이나 양수 열거자만 있는 흔한 경우 unsigned 와 같음. 메모가 비고에 남음).
      2. 원본이 UNRESOLVED 이고 선언 근거(evidence)도 없으며 이름이 관용 정수 typedef(uint8, uint8_t, U16, BOOL …)이면
         → 관용 폭·부호로 해석 + 가정 메모. (원본은 'uint' 하나만 이렇게 처리)
    """
    spelled = ' '.join(name.split()) if isinstance(name, str) else name
    if isinstance(spelled, str) and (spelled == 'enum' or spelled.startswith('enum ')):
        return TypeLink('INTEGER', integer_type='unsigned int',
                        assumptions=(f"열거형({spelled}) 비트필드: unsigned 로 가정",))
    link = _original_resolve(self, file, name, root, seen)
    if link.status == 'UNRESOLVED' and not link.evidence and isinstance(spelled, str) and spelled in CONVENTIONAL_PROFILES \
            and spelled not in BASE_PROFILES:
        profile = CONVENTIONAL_PROFILES[spelled]
        base = _CANONICAL_INTEGER.get(profile)
        if base:
            return TypeLink('INTEGER', integer_type=base,
                            assumptions=(f"'{spelled}' typedef 미수집: 관용 타입 {base} 적용",))
    return link


bc._Registry.resolve = _registry_resolve

# =============================================================================
# 중첩 멤버 경로 지원 — 멤버 타입 표
# =============================================================================
_LAST_REGISTRY = {}
_original_bind_functions = bc._bind_functions


def _bind_functions_capture(registry, globals_by_scope):
    _LAST_REGISTRY['registry'] = registry
    groups = defaultdict(list)
    for unit in registry.units.values():
        for function in unit.functions:
            groups[function.name].append(function)
    replacements = {}
    for functions in groups.values():
        if len({f.return_type for f in functions}) < 2 or all(f.definition for f in functions):
            continue
        identities = {}
        for function in functions:
            origin = function.origin
            tokens = [Token(word, origin.line) for word in (*function.return_type, '_rv_return')]
            declaration = next(bc._variable_declarations(bc._Collector, None, tokens, origin, False), None)
            if declaration is None:
                break
            link = registry.resolve(origin.file, declaration.spelling)
            if link.status not in ('STRUCT', 'INTEGER', 'SCALAR'):
                break
            identities[function] = bc._type_identity(bc._qualify(link, declaration.qualifiers))
        else:
            # The original binder compares return-type spellings before it checks parameter
            # compatibility. Give equivalent typedefs the same spelling for that check.
            type_ids = {}
            for function, identity in identities.items():
                type_id = type_ids.setdefault(identity, len(type_ids))
                replacements[id(function)] = replace(function, return_type=(f'_rv_type_{type_id}',))
    if replacements:
        registry = copy(registry)
        registry.units = {
            file: replace(unit, functions=[replacements.get(id(f), f) for f in unit.functions])
            if any(id(f) in replacements for f in unit.functions) else unit
            for file, unit in registry.units.items()
        }
    return _original_bind_functions(registry, globals_by_scope)


bc._bind_functions = _bind_functions_capture
_original_build = bc.build_bitfield_catalog


def _member_table(registry, catalog):
    """정의별 멤버 → (타입 철자, 타입 정의, 포인터 깊이, 배열 차원) 표.

    analysis.Struct.members / nested 에 있는 정보를 카탈로그 쪽으로 옮긴 것. 이름 있는 타입은 registry.resolve 로,
    익명 중첩 struct/union 은 nested 의 Struct 객체로 정의를 찾는다.
    """
    by_struct = {id(st): origin for origin, (st, _c) in registry.definitions.items()}
    table = {}
    for origin, (st, _conditional) in registry.definitions.items():
        row = {}
        for member, spelling in st.members.items():
            if not isinstance(spelling, str):
                continue
            declaration = getattr(st, 'member_declarations', {}).get(member)
            depth = declaration.pointer_depth if declaration else 0
            arrays = declaration.arrays if declaration else ()
            base = declaration.spelling if declaration else spelling
            target = None
            nested = st.nested.get(member)
            if nested is not None:
                target = by_struct.get(id(nested))
            elif base and base not in ('function pointer', 'struct', 'union', 'enum') and not base.startswith('<'):
                try:
                    link = registry.resolve(origin.file, base)
                except RecursionError:
                    link = None
                if link is not None and link.definition in catalog.definitions:
                    target = link.definition
                    depth += link.pointer_depth
            row[member] = (base, target, depth, arrays)
        table[origin] = row
    return table


def build_bitfield_catalog(paths, encoding='utf-8-sig'):
    """원본 build_bitfield_catalog 를 호출한 뒤 catalog.member_types(정의별 멤버 타입 표)를 덧붙인다."""
    catalog = _original_build(paths, encoding)
    registry = _LAST_REGISTRY.pop('registry', None)
    catalog.member_types = _member_table(registry, catalog) if registry is not None else {}
    catalog.function_declarations = {}
    if registry is not None:
        for unit in registry.units.values():
            for function in unit.functions:
                catalog.function_declarations.setdefault(function.name, []).append(function)
    return catalog


bc.build_bitfield_catalog = build_bitfield_catalog
if hasattr(eng, 'build_bitfield_catalog'):
    eng.build_bitfield_catalog = build_bitfield_catalog

# 함수 찾기 ----------------------------------------------------------------------
_original_lookup_function = bc.BitfieldCatalog.lookup_function


def lookup_function(self, file, name):
    """원본 lookup_function 을 감싼다.

    달라진 점
      1. 파일 힌트(xlsx 파일명에서 추정한 .c)로 거르면 후보가 0 이 되어 AMBIGUOUS 가 되던 경우 → 힌트 없이 다시 찾아
         유일한 정의면 그것을 쓴다. (원본: 유일한 Motor_Run 정의가 있어도 xlsx 이름이 다른 모듈이면 '함수 정의를 확정할 수 없음')
      2. 시트 이름 전체가 함수명이 아닐 때('3.2 Motor_Run', '3. 모터 구동 (Motor_Run)', 'Motor_Run()') 시트 이름 속
         식별자 토큰 중 함수 이름과 일치하는 것이 하나면 그 함수로 찾는다.
    """
    scope = _original_lookup_function(self, file, name)
    if scope.status not in ('MISSING', 'AMBIGUOUS'):
        return scope
    if file is not None:
        retry = _original_lookup_function(self, None, name)
        if retry.status not in ('MISSING', 'AMBIGUOUS'):
            return retry
    if scope.status == 'MISSING':
        words = [w for w in _WORD.findall(name or '') if w not in ('Function', 'function')]
        tries = list(words)
        for size in (2, 3, 4):                                  # 'Motor Run' → 'Motor_Run'
            tries += ['_'.join(words[k:k + size]) for k in range(len(words) - size + 1)]
        hits = list(dict.fromkeys(spelling for word in dict.fromkeys(tries) if word != name
                                  for spelling in matching_names(self.functions, word)))
        if len(hits) == 1:
            found = _original_lookup_function(self, file, hits[0])
            if found.status in ('MISSING', 'AMBIGUOUS') and file is not None:
                found = _original_lookup_function(self, None, hits[0])
            if found.status not in ('MISSING', 'AMBIGUOUS'):
                return found
    return scope


bc.BitfieldCatalog.lookup_function = lookup_function

# =============================================================================
# engine — 중첩 멤버 경로(a.b.c, a[i].b.c, p->b.c)
# =============================================================================
# 원본 MEMBER_ACCESS: base[index]?(.|->)member  한 단계만 허용 → 그 외는 MEMBER_PATH_UNRESOLVED.
# 패치: member 그룹이 나머지 경로 전체를 받는다. _catalog_match 는 그대로 두고, _match_binding / _member_catalog_match 가
#       경로를 따라 내려가 마지막 멤버의 비트필드를 비교한다.
eng.MEMBER_ACCESS = re.compile(
    r'(?P<base>[A-Za-z_]\w*)\s*(?:\[\s*(?P<index>\w+)\s*\]\s*(?:\[\s*\w+\s*\]\s*)*)?(?P<operator>\.|->)\s*'
    r'(?P<member>[A-Za-z_]\w*(?:\s*\[\s*\w+\s*\])*(?:\s*(?:\.|->)\s*[A-Za-z_]\w*(?:\s*\[\s*\w+\s*\])*)*)')
_SEGMENT = re.compile(r'(->|\.)?\s*([A-Za-z_]\w*)\s*((?:\[\s*\w+\s*\]\s*)*)')


_SUBSCRIPT = re.compile(r'\[\s*(\w+)\s*\]')


def _split_path(member):
    return [(op, name, _SUBSCRIPT.findall(subs)) for op, name, subs in _SEGMENT.findall(member)]


def _c_index_literal(text):
    match = bc._INTEGER.fullmatch(text)
    if not match or len(match[1]) > 24:
        return None
    digits = match[1]
    base = 0 if digits.lower().startswith(('0x', '0b')) else 8 if digits.startswith('0') else 10
    return int(digits, base)


def _access_issue(name, indices, depth, arrays, operator=None):
    for index, size in zip(indices, arrays):
        value, bound = _c_index_literal(index), _c_index_literal(size)
        if value is not None and bound is not None and value >= bound:
            if bound == 0:
                return f"배열 '{name}'에 접근할 수 있는 원소가 없습니다."
            reason = f"'{name}[{index}]'는 배열 범위를 벗어났습니다 (0~{bound - 1})."
            if value == bound:
                reason += f" [{index}]가 원소 개수 표기인지 확인해 주세요."
            return reason
    remaining = len(arrays) - len(indices)
    depth += min(0, remaining)
    if depth < 0:
        return f"'{name}'에 붙은 배열 번호가 너무 많습니다."
    if operator is not None:
        if remaining > 0:
            return f"'{name}'의 멤버에 접근하려면 배열 번호가 필요합니다."
        if depth > 1:
            return f"'{name}'은 포인터를 한 번 더 풀어야 멤버에 접근할 수 있습니다."
        expected = '->' if depth else '.'
        if operator != expected:
            return f"'{name}'의 멤버에는 '{expected}'로 접근해야 합니다."
    return ''


def _descend(bitfields, candidates, path):
    """candidates(origin→CatalogStruct) 에서 중간 경로 조각들을 따라 내려간다. (후보들, 실패 사유) 반환."""
    table = getattr(bitfields, 'member_types', {}) or {}
    current = dict(candidates)
    for index, (_, name, subscripts) in enumerate(path[:-1]):
        following = {}
        for origin, definition in current.items():
            row = table.get(origin, {})
            names = matching_names(row, name)
            if len(names) != 1:
                reason = (f"'{definition.name}'에서 중간 멤버 '{name}'을 찾지 못했습니다." if not names
                          else f"대소문자만 다른 멤버 후보가 여러 개 있음: {', '.join(sorted(names))}")
                return {}, reason
            spelling, target, depth, arrays = row[names[0]]
            issue = _access_issue(name, subscripts, depth, arrays, path[index + 1][0])
            if issue:
                return {}, issue
            if target is None or target not in bitfields.definitions:
                return {}, f"중간 멤버 '{name}'의 타입 '{spelling}' 정의를 찾지 못했습니다."
            following[target] = bitfields.definitions[target]
        current = following
    return current, ''


def _compare_member_path(bitfields, candidates, path, reason):
    if len(path) > 1:
        candidates, failure = _descend(bitfields, candidates, path)
        if not candidates:
            return 'REVIEW', f"{failure} / {reason}", None
    _, member, indices = path[-1]
    if indices:
        for origin in candidates:
            row = bitfields.member_types.get(origin, {})
            names = matching_names(row, member)
            if len(names) == 1:
                _, _, depth, arrays = row[names[0]]
                issue = _access_issue(member, indices, depth, arrays)
                if issue:
                    return 'REVIEW', f"{issue} / {reason}", None
    return eng._compare_member_candidates(list(candidates.values()), member, reason)


def _origin_summary(declarations, label):
    declarations = tuple(declarations)
    paths_by_name = defaultdict(set)
    for declaration in declarations:
        file = declaration.origin.file
        paths_by_name[Path(file).name.casefold()].add(file)
    locations = []
    for declaration in declarations:
        file, line = declaration.origin.file, declaration.origin.line
        name = Path(file).name
        location = name if len(paths_by_name[name.casefold()]) == 1 else file
        locations.append(f"{location}:{line} ({label(declaration)})")
    locations = list(dict.fromkeys(locations))
    if not locations:
        return ''
    return ' / ' + ', '.join(locations[:2]) + (f' 외 {len(locations) - 2}곳' if len(locations) > 2 else '')


def _match_binding(bitfields, access, binding):
    """변수 선언 및 경로의 각 단계에서 배열·포인터 형태를 확인한 뒤 끝 멤버를 비교한다."""
    name = access['base']
    path = _split_path(access['member'])
    head = access.string[:access.start('operator')] if hasattr(access, 'string') else ''
    indices = _SUBSCRIPT.findall(head) if head else ([access['index']] if access['index'] else [])

    link = binding.type_link
    definition = link.definition if link else None
    if binding.status in ('INTEGER', 'SCALAR') or (link and link.integer_type and not definition):
        return 'SKIP', '', None
    if binding.status not in ('STRUCT', 'CONDITIONAL'):
        details = _origin_summary(binding.declarations, lambda d: f"{d.storage} {d.spelling}".strip())
        if binding.status == 'AMBIGUOUS':
            reason = f"변수 '{name}'의 타입 후보를 하나로 정할 수 없습니다."
        elif binding.status == 'UNRESOLVED':
            reason = f"변수 '{name}'의 타입 정의를 찾지 못했습니다."
        else:
            reason = f"변수 '{name}'의 선언을 확인할 수 없습니다."
        return 'REVIEW', reason + details, None
    first = binding.declarations[0]
    reason = f"C 선언 {Path(first.origin.file).name}:{first.origin.line} ({first.spelling})"
    candidates = {definition: bitfields.definitions[definition]} if definition in bitfields.definitions else {}
    if not candidates:
        for declaration in binding.declarations:
            if declaration.target in bitfields.definitions:
                record = bitfields.definitions[declaration.target]
                candidates[record.origin] = record
                continue
            spelling = declaration.spelling.removeprefix('struct ')
            names = matching_names(bitfields.type_candidates, spelling)
            if not names:
                return 'REVIEW', f"'{name}'의 선언 타입 '{spelling}' 후보를 찾지 못함 / {reason}", None
            for spelling in names:
                candidates.update((d.origin, d) for d in bitfields.type_candidates[spelling])
    for declaration in binding.declarations:
        arrays = declaration.arrays
        depth = link.pointer_depth if definition else declaration.pointer_depth
        if declaration.kind == 'PARAMETER' and len(arrays) == 1:
            depth += 1
            arrays = ()
        issue = _access_issue(name, indices, depth, arrays, access['operator'])
        if issue:
            return 'REVIEW', f"{issue} / {reason}", None
    if len(path) > 1:
        reason += f" / 중첩 경로 {access['member']}"
    return _compare_member_path(bitfields, candidates, path, reason)


def _member_catalog_match(bitfields, access, missing_reason):
    """원본 engine._member_catalog_match 교체(선언을 못 찾아 변수 이름으로 타입을 추정하는 경로). 중첩 경로만 추가."""
    base = access['base']
    path = _split_path(access['member'])
    member = path[-1][1]
    stems = [base]
    if access['operator'] == '->' and base.casefold().endswith('_p'):
        stems.insert(0, base[:-2])
    attempts = list(dict.fromkeys(s for stem in stems for s in (stem + '_t', stem)))
    candidates, type_names = {}, ()
    for spelling in attempts:
        type_names = matching_names(bitfields.type_candidates, spelling)
        if type_names:
            for tn in type_names:
                candidates.update((d.origin, d) for d in bitfields.type_candidates[tn])
            break
    if not candidates:
        return 'REVIEW', missing_reason + f" / 변수명에 대응하는 구조체 타입을 찾지 못함 ({', '.join(attempts)})", None
    locations = [f"{d.name} ({Path(d.origin.file).name}:{d.origin.line})" for d in candidates.values()]
    shown = ', '.join(locations[:5]) + (f" 외 {len(locations) - 5}개" if len(locations) > 5 else '')
    reason = (f"변수 선언 연결 없이 변수명으로 타입 추정: '{base}' → {', '.join(type_names)} / "
              f"해당 구조체의 '{member}' 비교 / {shown}")
    return _compare_member_path(bitfields, candidates, path, reason)


def _truncated_local_scope(bitfields, source_file, function_name, variable):
    # Excel 시트 제목에는 31자 제한이 있다. 긴 접두어, 같은 C 파일, 정확한
    # 지역 선언이 모두 확인될 때에만 잘린 제목을 복원한다.
    if not source_file or not function_name or len(function_name) < 24 or not identifier(function_name):
        return None
    candidates = [(name, scope) for name, scopes in bitfields.functions.items()
                  if name != function_name and name.startswith(function_name)
                  for scope in scopes if scope.file == source_file]
    if len(candidates) == 1:
        name, scope = candidates[0]
        binding = scope.variables.get(variable)
        if scope.status == 'RESOLVED' and binding and binding.status == 'STRUCT':
            return name, scope
    return None


def _catalog_match(catalog, access, source_file, function, function_name=''):
    """원본 engine._catalog_match 교체.

    달라진 점 한 가지: 변수 범위(scope)로 쓸 파일을 고를 때 **함수 정의가 있는 파일을 xlsx 파일명에서 추정한 .c 보다
    우선**한다. 원본은 `source_file or function.file` 라서 xlsx 이름이 다른 모듈(fan_Detail.xlsx)이면 Motor_Run 시트의
    변수를 fan.c 의 같은 이름 static 변수와 비교해 버렸다(틀린 범위로 MISMATCH 판정 → 덮어쓰기 위험).
    """
    bitfields = catalog.bitfields
    name, member = access['base'], access['member']
    if function.status == 'MISSING':
        resolved = _truncated_local_scope(bitfields, source_file, function_name, name)
        if resolved:
            full_name, scope = resolved
            status, reason, field = _catalog_match(catalog, access, source_file, scope, function_name=full_name)
            return status, f"{reason} / 시트 함수명 '{function_name}' → '{full_name}'", field
    if function.alternatives:
        scopes = function.alternatives
        local_scopes = [scope for scope in scopes if matching_names(scope.variables, name)]
        if local_scopes and len(local_scopes) != len(scopes):
            shared = bitfields.lookup_variable(function.file or source_file, name)
            if shared.status == 'MISSING' and bitfields.lookup_variable(None, name).status == 'MISSING':
                scopes = local_scopes
        matches = [_catalog_match(catalog, access, source_file, scope, function_name=function_name) for scope in scopes]
        return eng._agree_member_matches(matches, name, member)
    if function.status not in ('MISSING', 'RESOLVED', 'CONDITIONAL'):
        declarations = bitfields.function_declarations
        choices = [d for spelling in matching_names(declarations, function_name or '')
                   for d in declarations[spelling]]
        definitions_here = [d for d in choices if d.definition and d.origin.file == function.file]
        if len(definitions_here) == 1:
            visible = bitfields.source_scopes.get(function.file, {})
            choices = definitions_here + [d for d in choices if not d.definition and d.origin.file in visible
                                          and (not d.internal or d.origin.file == function.file)]
        definitions = [d for d in choices if d.definition]
        if len(definitions) > 1:
            reason = f"함수 '{function_name}'의 정의가 여러 곳에 있습니다."
            choices = definitions
        elif len(definitions) == 1 and len(choices) > 1:
            reason = f"함수 '{function_name}'의 선언과 정의를 연결하지 못했습니다."
        else:
            reason = f"함수 '{function_name}'의 정의를 확인할 수 없습니다."
        details = _origin_summary(choices, lambda d: '정의' if d.definition else '선언')
        return 'REVIEW', reason + details, None
    source_file = function.file or source_file or None
    binding = bitfields.lookup_variable(source_file, name, function.variables)
    if binding.status == 'MISSING' and source_file is not None:
        binding = bitfields.lookup_variable(None, name)
    if binding.status == 'MISSING':
        return eng._member_catalog_match(bitfields, access,
                                         eng._missing_declaration_reason(bitfields, name, source_file, function_name))
    match = eng._match_binding(bitfields, access, binding)
    if binding.declarations and function.variables.get(binding.declarations[0].name) is binding \
            and all(d.conditional for d in binding.declarations):
        fallback = bitfields.lookup_variable(source_file, name)
        if fallback.status != 'MISSING':
            return eng._agree_member_matches([match, eng._match_binding(bitfields, access, fallback)], name, member)
    return match


eng._match_binding = _match_binding
eng._member_catalog_match = _member_catalog_match
eng._catalog_match = _catalog_match

# 성능: _input_source(xlsx 이름 → .c 파일 추정) 는 시트마다 소스 파일 전체를 훑는다(3,000 파일 × 120 시트 = 수십만 Path 연산).
# 같은 카탈로그·같은 힌트는 결과가 같으므로 캐시한다.
_original_input_source = eng._input_source
_INPUT_SOURCE_CACHE = {}


def _input_source(bitfields, hint):
    key = (id(bitfields), hint)
    cached = _INPUT_SOURCE_CACHE.get(key)
    if cached is not None and cached[0] is bitfields:
        return cached[1]
    result = _original_input_source(bitfields, hint)
    if len(_INPUT_SOURCE_CACHE) > 4096:
        _INPUT_SOURCE_CACHE.clear()
    _INPUT_SOURCE_CACHE[key] = (bitfields, result)
    return result


eng._input_source = _input_source
