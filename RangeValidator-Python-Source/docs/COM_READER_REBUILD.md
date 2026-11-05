# Excel COM Reader rebuild ? 2026-11-03

Reference: user-provided C:/Dev/RangeValidator/excel_io.py, read only. Company documents were not used in development or transmitted.

## Ported behavior

V3 rangevalidator/com_values.py implements _excel, _open, _find_last, _sample_rows, _last_cell, _copy_sheet, read_workbook_values and close_excel around the reference flow. DispatchEx, thread-local reuse, ReadOnly=True, Notify=False, UpdateLinks=0, disabled events/alerts/screen updates and AutomationSecurity=3 remain. A comparison session owns one Excel instance across Common and detail workbooks, then closes it in finally. Standalone reads close their own instance. Strict STA/owned-process checks and failure reporting are retained.

Find now uses exactly the reference search keywords: What="*", After=A1, LookIn=xlFormulas, LookAt=xlPart, SearchOrder=rows/columns, SearchDirection=xlPrevious, MatchCase=False. Only last row and last column are searched. No first_column consistency precondition, MatchByte or SearchFormat arguments are added.

Fallback uses UsedRange endpoints only as bounded exploration hints, End(xlUp) over at most 256 columns, End(xlToLeft) over reference _sample_rows, and the reference final UsedRange column guess capped at 64. Values are obtained with rectangle Value2. Bulk Formula plus explicit HasFormula for strings beginning '=' preserves the V3 formula safety metadata. Merged addresses and hidden/veryHidden states are retained; merge restoration is attempted only up to 50000 cells. Merge ranges may extend beyond the last populated cell, as in the reference.

## Deliberate safety differences

Macro security failure remains fatal. Open returning None searches for a unique exact FullName, rather than the reference basename-only lookup. Close/Quit/STA/process-exit failures remain failures with independent evidence. No output removal or sheet deletion is imported.

The existing public Workbook/Sheet/Cell models and Reader signatures remain. Readers delegate COM work to the rebuilt adapter; openpyxl-first automatic fallback stays unchanged. Analysis, Common parsing, design parsing, engine, GUI, Writer, report and preservation implementation files are unchanged against the preceding distribution hashes.

## FindBoundsError and limits

Missing/failed Find is no longer an immediate FIND_BOUNDS_INCONSISTENT error. End fallback is attempted first. Successful fully explored fallback can read normally. FindBoundsError is retained for genuinely unconfirmed extent: sampled rows, columns capped by 256, UsedRange-only final guess, failed End probes or invalid coordinates. Failed Find cannot falsely certify an empty workbook.

Supported documents remain Common Parameter and A/B detailed-design formats. Every confirmed column is included, including A Name/Type/Range and B Name/Data type/Scope. Confirmed rectangles are never clipped. Defaults: 512 sheets, 250000 rectangle cells across the workbook, 16000000 text characters; merge restoration maximum 50000 cells per sheet. An over-budget rectangle, incomplete bulk matrix, unconfirmed fallback or merged sheet beyond the merge budget fails the whole file, discards partial models and blocks edits. A Common read failure blocks dependent comparison. These conservative failures can reject otherwise readable documents; they never represent truncated data as complete success. Formatting-only far UsedRange endpoints do not expand a successful Find rectangle.

Diagnostics show file/sheet numbers, Find results, UsedRange/End fallback stages, final rows/columns, limit failures and fixed omission/block reasons, merge restored/skipped, extraction outcome, exception type and HRESULT. Paths, names, values and arbitrary exception messages are excluded from default diagnostics.

## Writer comparison

Both implementations open the source read-only and save a new workbook through Excel. V3 additionally keeps certified selected MISMATCH only, immutable input evidence, locks, fresh comparison, formula/merged/duplicate/stale edit rejection, exclusive new output directories, reopened value checks and preservation verification. The reference's output deletion, sheet deletion and generic cell mutation are not ported. Writer source is unchanged. Its shared bounds helper now uses the new strict fallback routine; uncertain preservation bounds still fail closed. Writer/report test suites were not run in this task.

## Validation and company recheck

Only tools/run_com_reference_tests.py is the source-test allowlist; see COM_REFERENCE_TEST_RUN.json/.log for final counts. It includes direct reference function comparisons, fallback/first-column cases, A/B/Common and engine integration, merged/hidden sheets, large rectangles, limits, COM failures, cleanup and actual desktop Excel on sanitized workbooks. Full suite was not run. Package receipts and hashes are recorded in DISTRIBUTION.json and COM_REFERENCE_DISTRIBUTION_VALIDATION.json.

On the company PC: extract the entire new ZIP to a fresh local folder; verify the EXE hash using Get-FileHash -Algorithm SHA256; confirm desktop Excel and the normal DRM client can open the same inputs under the same account. Run the EXE, register Common, multiple headers and A/B documents, then compare. Check file success/error counts and diagnostic file/sheet numbers, com.bounds.find/end/usedrange/final, com.extract and com.process.exit. Confirm required Scope/Range columns and last known variable rows are present. Any bounds.blocked, bounds.limit or merge.skipped FAILED event must leave the file uneditable. Test both DRM-permitted and policy-denied files; denial must stay a file error without blocking unrelated successful files. Only then test a selected certified MISMATCH into a NEW output folder, reopen it under company policy, and verify values, formulas, merges, images, hidden sheets, formatting, source hash and protection. Keep company paths/data local; export only approved masked logs.

Actual company DRM acceptance, company output/report protection, physical clean-PC/EDR validation and arbitrary/hung workbook behavior remain unverified.


## Partial read behavior (2026-11-04)

If a sheet cannot be extracted, the Reader records a sheet ERROR and continues with other sheets. If rectangle Value2 or Formula extraction fails, it retries individual cells; failed coordinates are omitted and recorded with row/column, HRESULT and exception class, while readable cells remain available to the comparator. Failed sheet and cell content never enters the Workbook model as a guessed value.

A partial workbook is explicitly marked PARTIAL. Available rows retain MATCH/MISMATCH/REVIEW/SKIP results, but every auto-fix candidate for that workbook is disabled. GUI file status and Errors view show the read failures. This prevents an unread cell or skipped sheet from being presented as complete evidence.


## Proven-reader mechanism alignment (2026-11-04)

The COM reader now follows the supplied code when Find bounds are absent: Use UsedRange endpoints as hints, scan up to 256 columns with End(xlUp), sample rows with End(xlToLeft), then use min(UsedRange columns, 64) if the right edge is still unknown. It does not reject sampled rows, that 64-column estimate, or large content rectangles. COM initialization uses CoInitialize semantics. If Open returns None, the opened workbook is located by Name as in the reference. AutomationSecurity/AskToUpdateLinks assignment failures are tolerated as in the reference. Large sheets still read their full Value2 rectangle; merge restoration is skipped above 50000 cells. This mirrors the proven read flow and can accept estimated bounds.
