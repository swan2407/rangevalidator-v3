"""Read-only index for browsing an existing comparison session."""
from collections import Counter
from dataclasses import dataclass, field

from .engine import is_function_sheet
from .workflow import canonical

STATES = ('MATCH', 'MISMATCH', 'REVIEW', 'CREATE', 'SKIP')


@dataclass
class SheetResults:
    rows: list = field(default_factory=list)
    counts: Counter = field(default_factory=Counter)
    status: str = ''
    reason: str = ''


@dataclass
class FileResults:
    path: str
    sheets: dict[str, SheetResults] = field(default_factory=dict)
    counts: Counter = field(default_factory=Counter)
    record: object = None


class ResultIndex:
    def __init__(self, session):
        self.files = {}
        self.rows = {}
        self.counts = Counter()
        if session is None:
            return

        paths = {}

        def file_for(path):
            if path not in paths:
                key = canonical(path)
                if key not in self.files:
                    self.files[key] = FileResults(path)
                paths[path] = self.files[key]
            return paths[path]

        def sheet_for(path, name):
            sheets = file_for(path).sheets
            if name not in sheets:
                sheets[name] = SheetResults()
            return sheets[name]

        for path in session.details:
            file_for(path)
        for record in session.files:
            file_for(record.file).record = record
        for path, name, status, reason in session.coverage:
            if is_function_sheet(name):
                entry = sheet_for(path, name)
                entry.status, entry.reason = status, reason
        for result in session.results:
            if not is_function_sheet(result.sheet):
                continue
            self.rows[result.id] = result
            entry = sheet_for(result.file, result.sheet)
            entry.rows.append(result)
            entry.counts[result.status] += 1
            file_for(result.file).counts[result.status] += 1
            self.counts[result.status] += 1
        for error in session.errors:
            name = error.get('sheet_name') or ''
            if error.get('error_kind') == 'SHEET_READ_ERROR' and is_function_sheet(name):
                entry = sheet_for(error['file'], name)
                entry.status = 'PARTIAL' if entry.rows else 'ERROR'
                entry.reason = error.get('reason', '')

    def visible(self, file, sheet, status, search):
        if file not in self.files or sheet not in self.files[file].sheets:
            return []
        query = search.strip().casefold()
        if status == '일반 멤버':
            status = 'SKIP'
        return [r for r in self.files[file].sheets[sheet].rows
                if (status == '전체' or r.status == status) and query in r.expression.casefold()]
