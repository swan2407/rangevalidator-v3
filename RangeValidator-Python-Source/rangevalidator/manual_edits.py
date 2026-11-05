"""User-approved edits kept separate from immutable analysis results."""
from dataclasses import dataclass

from .engine import is_function_sheet, parse_interval

AUTO_MISMATCH = 'AUTO_MISMATCH'
AUTO_CREATE = 'AUTO_CREATE'
MANUAL_REVIEW = 'MANUAL_REVIEW'


@dataclass(frozen=True)
class ReviewRequest:
    row_id: str
    value: str
    selected: bool = False


@dataclass(frozen=True)
class ManualEdit:
    id: str
    file: str
    sheet: str
    cell: str
    expression: str
    existing: object
    expected: str
    reason: str
    status: str = 'REVIEW'
    kind: str = MANUAL_REVIEW

    @classmethod
    def from_result(cls, row, value):
        return cls(row.id, row.file, row.sheet, row.cell, row.expression, row.existing, value, row.reason)


def is_review_target(row):
    return row.status == 'REVIEW' and bool(row.cell) and is_function_sheet(row.sheet)


def manual_value(value):
    """Normalize a supported ordered integer range for writing as text."""
    try:
        bounds = parse_interval(value) if isinstance(value, str) and len(value) <= 32767 else None
    except ValueError:
        bounds = None
    if bounds is None or bounds[0] > bounds[1]:
        raise ValueError('최솟값과 최댓값을 정수 범위(예: 0~7)로 입력하세요.')
    return f'{bounds[0]}~{bounds[1]}'
