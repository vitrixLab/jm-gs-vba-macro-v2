# Phase 3 - Findings

Findings recorded in Phase 3. Code fixes C-01 through C-15 are committed on branch `audit/vba-docs-sync`; documentation fixes D-01 through D-13 are committed. Held items are in `docs/audit/blocked.md`.

Severity counts: **1 Critical, 4 High, 13 Medium, 9 Low, 1 Info** (28 total: 15 code, 13 documentation).
Phases: **P4 = documentation fixes**, **P5 = VBA source fixes**, **P6 = checker/report**. Behaviour-changing fixes needed a decision entry before editing.

## Resolution status

Fixed in code commits:
- C-01 `da6c790`, C-14 `f75bb4f`, C-03 `97c9c16`, C-04 `ae2c755`, C-07 `49ff9dd`, C-11 `b1c4569`, C-12 `5fce068`, C-13 `5a63ce5`, C-02 `145145e`, C-15 `56bc67b`, C-08 `ab1b04d`, C-09 `26a6da3`.
- Documentation fixes D-01 through D-13 in commits `daf0e52` and `0a9d203`.

Held (do not edit; see `docs/audit/blocked.md`):
- C-05, C-06, C-10.

## Code findings

| ID | Severity | Finding | Evidence | Proposed fix |
|---|---|---|---|---|
| C-01 | Critical | `ValidateV8Structure` declares `n` twice in the same procedure; the project cannot compile as exported. | `modGLGate.bas:13` (`n As Variant`), `:29` (`n As Long`) | P5 mechanical: keep `n` as the `For Each` Variant and rename the counter (e.g. `nonblank`); add a compile check |
| C-02 | High | `InitializeEngine` wraps all loader calls in `On Error Resume Next` and sets `gInitialized = True` even when a loader fails. Callers then use partial/empty caches. | `modEngine.bas:130-142` | P5 decision: fail loudly or record load status; set the flag only after successful loads |
| C-03 | High | `ScanGL` has no top-level error handler. A runtime error in `ScanCDJ`/`ScanCRJ`/mapping can bypass `GL_AUDIT` HOLD handling. | `modGLScanEngine.bas:20-68`; no handler around `:34-35` | P5 decision: add a top-level handler that writes a HOLD audit if possible and re-raises |
| C-04 | Medium | `RefreshGL`/`RefreshAllGL`/`RefreshGLIntoSheet` call `ScanGL` directly without `ValidateV8Structure`; only `RunV8Gate` validates the workbook structure first. Missing sheets/layout therefore produce an unhandled runtime error, not a gate HOLD. | `modGLRefresh.bas:13-26`; `modGLGate.bas:39-48` | P5 decision: make structure validation part of `ScanGL` or every public wrapper |
| C-05 | Medium | `Sheet9` uses undeclared `START_ROW` and `COL_*` values, including `COL_DEBIT_MIRROR`, without `Option Explicit`; comparisons and column references do not resolve as intended. | `Sheet9.txt:1,72,78,135,163,169,187,205,212,226,249,252` | P5 decision: declare the constants from the confirmed Sheet9 layout and add `Option Explicit`; do not guess missing column values |
| C-06 | Medium | `SyncPJNonVatToCDJ` never writes `CDJ_COL_DAY` on update or insert; an existing row can retain a stale day. | `modEngine.bas:1267-1286,1299-1319`; VAT path writes it at `:884,926` | P5 mechanical: map the Sheet11 day source and write `CDJ_COL_DAY` on both paths |
| C-07 | Medium | `SyncSJToCRJ` assigns `amount` to a `Double` before testing `IsNumeric`, so a text amount raises a type-mismatch error before the guard. | `modEngine.bas:985,990` | P5 mechanical: validate in a Variant first, then convert |
| C-08 | Medium | `Sheet11`, `Sheet12` and `Sheet8` cleanup labels restore application state but swallow the original error; Sheet11/Sheet12 continue into Part 2 after a failure. | `Sheet12.txt:204,252-258`; `Sheet11.txt:207,259-265`; `Sheet8.txt:40,84-88` | P5 decision: re-raise after cleanup or exit before Part 2 |
| C-09 | Medium | `ImportSuppliersFromJSON` clears all rows before validation and checks duplicates only after all writes, with no rollback. | `ImportSuppliersFromJSON.bas:13,815-847` | P5 decision: validate before clear, or write to a temporary area and commit atomically |
| C-10 | Medium | `LookupByTIN_NonVat` writes the first match before the duplicate check; on a second match it exits, leaving a partial/ambiguous supplier row populated. | `modPJSync.bas:165-181` | P5 decision: collect matches first and write only after proving uniqueness |
| C-11 | Low | `ResetCRJ` and `ResetCDJ` disable `ScreenUpdating`/`EnableEvents` without an error handler; an error leaves Excel in the altered state. | `modEngine.bas:1163-1187,1189-1213` | P5 mechanical: restore state in a cleanup handler |
| C-12 | Low | `CDJReserveRow` unconditionally sets `EnableEvents = True` in `Done`, even if events were already disabled by the caller. | `modEngine.bas:443-463` | P5 mechanical: save/restore the prior value |
| C-13 | Low | `ResetCDJ` reports sequence/next-row cells as `Z13`/`Z14`; the constants resolve to `AA13`/`AA14`. | `modEngine.bas:64-71,1211-1212` | P5 mechanical: derive the addresses from the constants |
| C-14 | Low | `WriteScanAudit` uses `On Error Resume Next` around the sheet lookup and add; if `Worksheets.Add` fails, `ws` is still `Nothing` and the following write raises error 91. | `modGLScanEngine.bas:311-321` | P5 mechanical: handle add failure explicitly and return a status |
| C-15 | Medium | `LoadCDJReferences`/`LoadCRJReferences` use `Trim(ws.Cells(i, c).Value)`, which raises type mismatch on error cells (#N/A, #REF!) and, under the old `Resume Next`, silently truncated the cache. | `modEngine.bas:363,365,367`; `modEngine.bas:533` | P5 mechanical: replace with existing `CellText(ws, i, c)` helper |

## Documentation findings

| ID | Severity | Finding | Evidence | Proposed correction |
|---|---|---|---|---|
| D-01 | High | Both documents claim sync write errors are swallowed by empty `ErrorHandler:` labels. `On Error GoTo 0` disables the handler before the writes, so later errors propagate. | `VBA_CODEBASE.md:92`; `VBA_CODE_FLOW.md:895-899`; `modEngine.bas:851,978,1249` | P4: describe the real boundary; only pre-write/pointer-update failures are suppressed |
| D-02 | High | The documents present the recorded `compile PASS` as current evidence, but the log predates `modGLGate.bas` and the workbook and the current source has C-01. | `VBA_CODEBASE.md:99,5779`; `build/artifacts/compile.log:1` (17:42); `modGLGate.bas` mtime 18:28 | P4: label the artifact stale and state that it cannot validate the current export |
| D-03 | Medium | The entry catalogue says SJ2 uses the same sales formulas as SJ; Sheet8 writes only M/O, while Sheet3 also writes L/N mirror formulas. | `VBA_CODE_FLOW.md:35`; `Sheet8.txt:58-64`; `Sheet3.txt:72-75` | P4: replace with the actual SJ2 formula set |
| D-04 | Medium | The SJ invoice-stamp condition is documented as requiring date + amount. Sheet3 checks amount/TOTAL only; Sheet8 checks date. | `VBA_CODE_FLOW.md:248,1095`; `Sheet3.txt:62-63,78`; `Sheet8.txt:67-69` | P4: split the Sheet3 and Sheet8 conditions |
| D-05 | Medium | The SJ/SJ2 layout table conflates the two sheets: it says column I is a manual physical invoice not overwritten by Sheet3 and lists U as the auto invoice for both. Sheet8's I is auto invoice and U is reference; Sheet3's I is manual physical, U is auto invoice, V is full ID. | `VBA_CODE_FLOW.md:829-846`; `Sheet8.txt:13-18`; `Sheet3.txt:17-19` | P4: separate SJ and SJ2 tables |
| D-06 | Medium | The shared CDJ field list includes “day” for both syncs. `SyncPJNonVatToCDJ` does not write day. | `VBA_CODE_FLOW.md:208`; `modEngine.bas:1267-1286,1299-1319` | P4: mark day as VAT-path-only and record the PNV gap |
| D-07 | Medium | Fixed-map tables omit `V8_CDJMap` column 6 and `V8_CRJMap` column 8. | `VBA_CODE_FLOW.md:493-494`; `modGLWorkbookMap.bas:157,195` | P4: add the missing columns and accounts |
| D-08 | Low | The hard-coded alias summary omits `COST OF REVENUE`, `SUPPLIES`, `BANK CHARGE` and `CHARGES`. | `VBA_CODE_FLOW.md:495`; `modGLWorkbookMap.bas:140-143` | P4: complete the alias list |
| D-09 | Low | The `Sheet9` undeclared-constant list omits `START_ROW` and `COL_DEBIT_MIRROR`. | `VBA_CODEBASE.md:81`; `Sheet9.txt:72,169,226,249` | P4: complete the list |
| D-10 | Low | “Sundry T/U disagreements are audit-only” is true for numeric disagreements but not for nonnumeric T/U values, which increment `inv` and can force HOLD. | `VBA_CODEBASE.md:91`; `modGLScanEngine.bas:152-162` | P4: distinguish numeric mismatch from invalid input |
| D-11 | Low | The tracked `vba/` copies are described as byte-identical except the export header. They also have a UTF-8 BOM and trailing blank lines. | `VBA_CODEBASE.md:24,5770` | P4: say “same logic after normalising header/BOM/trailing blanks” |
| D-12 | Low | The flow says the active non-VAT path is Sheet11. `Sheet9.Worksheet_Change` is still wired and callable; it is broken, not disconnected. | `VBA_CODE_FLOW.md:235`; `Sheet9.txt:2` | P4: call Sheet11 the intended path and list Sheet9 as a reachable broken fallback |
| D-13 | Info | The empty-module sentence omits `CDJ` and `GJ` even though the table above it lists both. | `VBA_CODEBASE.md:73` | P4: add both names or remove the redundant sentence |
