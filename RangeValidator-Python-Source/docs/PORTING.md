# V1 porting boundaries and differences

## Inspected references

Read V1 docs/REQUIREMENTS.md, ROADMAP.md and TASKS.md before implementation.
Primary source references: src/bitfield_parser.cpp, type_matcher.cpp,
variable_expression_parser.cpp, project_type_index.cpp, common_type_resolver.cpp;
tests/parser_tests.cpp, range_tests.cpp, matching_tests.cpp,
project_analysis_tests.cpp and excel_com_tests.py. No V2 analysis was used.
Reference hashes for the direct differential subset are in
V1_DIFFERENTIAL_RESULTS.json. The oracle builds solely inside the V3 build folder.

## Implemented C subset

Tokenizer masks comments/literals, tracks member lines, balances declaration
braces and skips function bodies for extraction (function presence blocks binding).
Anonymous/tagged typedef structs and named structs; multiline fields; ordinary
simple integer members distinguished; positive decimal literal widths. Mixed bad
and good fields retain extraction evidence but never certify bindings through a
source-wide REVIEW. Exact simple three-directive include guards are allowed.

Range profile is V1 Windows GCC 32-bit int, two's-complement signed representation:
explicit unsigned/unsigned int and signed/signed int, width 1..32. The only scalar
alias supported is one exact `typedef unsigned int uint;`, before use. Other built-in
integer bit-fields may be extracted, but their ranges remain unavailable/REVIEW.
Plain int/char signedness is never guessed. General typedef chains are unsupported.

Bindings: a single source, previously defined unique struct alias/tag, object or
one `*`, optional `extern`; exact obj.member or ptr->member. Duplicate types,
variables, ordinary/bit-field member collisions, wrong operators, unknown bases,
macros, width expressions, conditions and complex declarations give REVIEW.
No matching by final member alone. Ordinary members are SKIP and have no range.

## Excel and comparison

Both Readers return Workbook/Sheet/Cell with original row, column, A1 coordinate,
value, formula flag/text, merge area/anchor and visible/hidden/veryHidden state.
COM error values are normalized to Excel error text; dates/times from openpyxl
are normalized to Excel serial numbers for Value2 compatibility.

openpyxl loads twice (formula and cached-value views), uses normal in-memory mode
for merge metadata, never saves, and uses the 3.x sparse `_cells` index to avoid
scanning style-only distant cells. Budgets apply after workbook loading; this is
not a bounded-memory ZIP parser. Formula results can be None if the file has no
cached result. COM may calculate an in-memory result: never report equality of
formula cached values between backends as guaranteed.

COM only examines the Find content rectangle. `Sheet.merge_scope=content_bounds`
means merges intersecting scanned cells are recorded; wholly blank merges outside
that rectangle are not discovered. openpyxl uses `merge_scope=all`. Empty COM sheets
have no cells/bounds. This coverage distinction is explicit in the common model.
Default limits: 512 sheets, 250000 rectangle cells, 16000000 text characters. COM
reads cells individually; large-sheet performance is not a V1 throughput claim.

Additional conservative comparison discovers basic A Name/Type/Range and B
Name/Data type/Scope headers by content. Type must agree with the resolved
bit-field. Formula/error/merged headers or data and duplicate headers give REVIEW.
MATCH uses decimal/scientific numeric pairs separated by ~ or ... Wrong/reversed
pairs, scalar and blank are read-only AUTO_FIX_CANDIDATE proposals; free text is
REVIEW. Every result includes document coordinates and source field/declaration.
No workbook write or actual correction is performed.

## Recorded differences, not full parity

The 47-case differential compares V1 parser field metadata/ranges, binding gate,
and single-header matcher status against Python. All except two explicitly
recorded differences agree; this is not V1's full CTest suite.

- Ordinary member: legacy single-header matcher returns REVIEW; Python SKIP
  reflects the requested integer exclusion (and V1 Common resolver's SKIP rule).
- Invalid C keyword member (`unsigned int int:2`): V1 parser extracts a field but
  its binding model rejects it. Python rejects it during extraction. Both block
  matching as REVIEW; no new certified result is introduced.
- JSON proposal status AUTO_FIX_CANDIDATE corresponds to V1 certified MISMATCH;
  this implementation has no Writer or edit certification API.

Intentionally unported V1 workflows: Common Parameter declaration workbook and its
section/context rules, filename/source context resolution, multi-file include
visibility/linkage/static scopes, symbol/file diagnostic isolation, Batch mode,
full template compatibility/merged-header rules and V1 report/Writer engine.
C arrays/qualifiers from V1 Common spreadsheet declarations are not accepted as C
source declarations; the original V1 single-header C binding subset also rejects
them. Unsupported expressions remain REVIEW, not last-member fallback.
No AST, macro evaluation, enum/ordinary integer validation, nested pointers,
functions/local/return value flow, CFG or alias analysis.

## COM lifecycle and remaining environment limits

CoInitializeEx(STA) must succeed; MTA is reported without an unmatched uninitialize.
Snapshot existing process IDs before DispatchEx; verify a new PID before settings,
Open or Quit. Cannot prove ownership => fail without altering/quitting that app.
Close(False), Quit, proxy release, CoUninitialize and owned PID exit check (30s)
are independent cleanup steps. No process is killed. Cleanup failure discards data
and is reported separately. Native stat/open is not an access prerequisite for COM.

A hung COM call/modal policy prompt is not interruptible in this first implementation;
no worker timeout/cancellation supervisor. XLM/add-ins/DRM policy must be managed by
the approved Excel environment. DispatchEx failure before ownership can be proven
cannot guarantee cleanup of a partially created server. Actual DRM, denied ACL,
long UNC paths and Office-uninstalled clean-PC behavior remain unverified.

API references consulted (no project data transmitted):
- https://learn.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open
- https://learn.microsoft.com/en-us/office/vba/api/Excel.Range.Find
- https://learn.microsoft.com/en-us/office/vba/api/Excel.Range.Value2
- https://learn.microsoft.com/en-us/office/vba/api/Excel.Application.AutomationSecurity
- https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.reader.excel.html
