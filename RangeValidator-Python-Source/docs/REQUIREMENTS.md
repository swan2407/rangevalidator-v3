# Requirements — authorized GUI/EXE milestone

- Work only in C:/Dev/RangeValidator-Python. V1 source/tests are read-only;
  V2 is not accessed or ported. Retain the original 78 tests and Workbook model.
- Korean Tk/ttk four-tab GUI, worker-owned COM STA, main-thread-only Tk updates.
- Common File/Name/Type/Value declarations and section rules; multi-header catalog,
  per-type conflicts, A/B discovery, multi-sheet/file comparison and exact identity.
- Production MATCH/MISMATCH/REVIEW/SKIP follows inspected V1 rules. Unsupported
  expressions/ambiguous evidence REVIEW; ordinary integers SKIP/no range inference.
- openpyxl-first read; conditional approved COM fallback. Failure is not DRM proof.
  DispatchEx/CoInitialize, strict STA check, ReadOnly/UpdateLinks=0/Notify=False/
  AddToMru=False/IgnoreReadOnlyRecommended; disable events, links, visibility,
  screen updates, alerts and programmatic macros. Find bounds, Value2 metadata.
- Selected certified MISMATCH only, no formula/data merge/duplicate/stale edits;
  immutable session fingerprints, input sharing locks, fresh comparison, cell
  recheck, in-memory edits, new exclusive output directory and COM SaveAs.
- Verify created file and reopened values; compare original hash/mtime and
  preservation evidence. No original overwrite, decrypted temporary workbook,
  password discovery, DRM bypass or policy changes.
- Native XLSX effective Excel styles/dimensions, values/formulas/merges/sheets,
  image payload/geometry checks. Unknown objects fail conservatively. COM-only
  metadata is partial evidence; protection and full preservation remain REVIEW.
- Final folder: corrected XLSX and RangeValidator_Report.xlsx, with five sheets.
  Default diagnostics omit sensitive values/paths; optional export only.
- PyInstaller Windows x64 onedir ZIP with runtime/libraries/notices/checksums;
  extract/hash/Windows-only PATH test. Actual DRM and clean PC not claimed tested.
- Excluded: ordinary integer validation, AST, CFG, alias analysis and function I/O.
