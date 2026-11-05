from dataclasses import dataclass, field
from typing import Any
from .error_info import classify_error


@dataclass(frozen=True)
class Cell:
    row: int
    column: int
    address: str
    value: Any
    is_formula: bool = False
    formula: str | None = None
    merged_range: str | None = None
    merge_anchor: str | None = None
    kind: str = "value"


@dataclass
class Sheet:
    name: str
    state: str  # visible, hidden, veryHidden
    cells: dict[str, Cell] = field(default_factory=dict)
    merged_ranges: list[str] = field(default_factory=list)
    bounds: tuple[int, int, int, int] | None = None
    merge_scope: str = 'all'


@dataclass
class Workbook:
    path: str
    backend: str
    sheets: list[Sheet] = field(default_factory=list)
    partial: bool = False
    read_errors: list[dict] = field(default_factory=list)


@dataclass(frozen=True)
class ReaderLimits:
    max_sheets: int = 512
    max_cells: int = 250_000  # scanned content rectangle, all sheets combined
    max_text: int = 16_000_000


class ReadFailure(Exception):
    def __init__(self, stage, cause, cleanup_errors=()):
        self.stage = stage
        self.code = getattr(cause, "hresult", None)
        if self.code is None:
            self.code = getattr(cause, "winerror", None)
        if self.code is None:
            self.code = getattr(cause, "errno", None)
        if self.code is None and type(cause).__name__=='com_error' and cause.args and isinstance(cause.args[0],int):
            self.code=cause.args[0]
        self.exception_type = getattr(cause,'exception_type',type(cause).__module__+'.'+type(cause).__qualname__)
        info=getattr(cause,'excepinfo',None)
        self.inner_hresult = info[5] if isinstance(info,tuple) and len(info)>5 and isinstance(info[5],int) else getattr(cause,'inner_hresult',None)
        self.error_class=getattr(cause,'error_class',classify_error(self.code,self.exception_type))
        self.api = getattr(cause,'api',stage)
        self.backend = getattr(cause,'backend','') or ('com' if stage.startswith('com.') else 'openpyxl' if stage.startswith('openpyxl.') else '')
        self.context = dict(getattr(cause,'context',{}))
        self.detail = str(cause)
        self.cleanup_errors = list(cleanup_errors)
        super().__init__(f"{stage}: code={self.code!r}: {self.detail}")

    def as_dict(self):
        code_hex = f"0x{self.code & 0xffffffff:08X}" if isinstance(self.code, int) else None
        return dict(stage=self.stage, code=self.code, code_hex=code_hex,
                    detail=self.detail, cleanup_errors=self.cleanup_errors, exception_type=self.exception_type,
                    inner_hresult=self.inner_hresult,api=self.api,context=self.context,error_class=self.error_class)


def address(row, col):
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"
