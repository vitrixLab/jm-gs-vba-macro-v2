# Phase 3 - Audit report

Date: 2026-11-11
Repository: vitrixLab/jm-gs-vba-macro-v2
Branch: audit/vba-docs-sync
Report commit: pending (Step 3)

## Scope and evidence boundary

- Canonical VBA export: build/v8_5_vba/ (tracked in commit dc9dbfa; code fixes committed on top).
- Workbook provenance: build/Global-Smile_2026-v8_5.xlsm, SHA-256 D35BF20944955697008216E23E0D112B1C24EF297EB6AC991454011C7FF15234.
- No workbook was opened, compiled, or run. The embedded-VS-exported boundary remains unverified. Do not claim the workbook compiles.
- DRR_deep-research-report.md is untracked and states it had no repository/code access; treated as hypothesis only.

## Phase completion

| Phase | Status | Output |
|---|---|---|
| 0 - Files and provenance | Complete | docs/audit/00_files.md, privacy_note.md |
| 1 - Inventory | Complete | docs/audit/01_inventory.md (15 components, 146 procedures) |
| 2 - Claim verification | Complete | docs/audit/02_claims.md (33 verified, 7 partial, 8 wrong) |
| 3 - Findings | Complete | docs/audit/03_findings.md (28 findings: 1 critical, 4 high, 13 medium, 9 low, 1 info) |
| 4 - Documentation fixes | Complete | VBA_CODEBASE.md daf0e52; VBA_CODE_FLOW.md 0a9d203 |
| 5 - VBA source fixes | Complete | 12 commits, listed below |
| 6 - Checker and report | Complete | scripts/check_docs.py, scripts/check_vba_static.py, this report |

## Code fix commits (one per finding, in required order)

| Commit | Finding | Summary |
|---|---|---|
| da6c790 | C-01 Critical | Renamed duplicate `n` counter in ValidateV8Structure (modGLGate.bas:13,29). |
| f75bb4f | C-14 Low | WriteScanAudit handles Worksheets.Add failure explicitly (modGLScanEngine.bas:311-321). |
| 97c9c16 | C-03 High | ScanGL top-level error handler: restores app state, writes HOLD audit if possible, shows MsgBox, re-raises. |
| ae2c755 | C-04 Medium | ValidateV8Structure called at start of ScanGL; RunV8Gate simplified to single choke point. |
| 49ff9dd | C-07 Medium | SyncSJToCRJ validates amount with IsNumeric before converting to Double. |
| b1c4569 | C-11 Low | ResetCRJ/ResetCDJ restore ScreenUpdating/EnableEvents and re-raise in cleanup. |
| 5fce068 | C-12 Low | CDJReserveRow saves/restores EnableEvents. |
| 5a63ce5 | C-13 Low | ResetCDJ derives sequence/next-row cell addresses from constants. |
| 145145e | C-02 High | InitializeEngine fails loudly for missing target sheet/SUPPLIERS DATA; sets gInitialized only after required loaders. |
| 56bc67b | C-15 Medium | LoadCDJReferences/LoadCRJReferences use CellText helper instead of Trim(ws.Cells(...)). |
| ab1b04d | C-08 Medium | Sheet8/Sheet11/Sheet12 show operator error and exit before Part 2 on Part 1 failure. |
| 26a6da3 | C-09 Medium | ImportSuppliersFromJSON parses and validates duplicates in memory before clearing rows. |

Held findings (blocked, no logic touched): C-05, C-06, C-10. See docs/audit/blocked.md for the exact open questions.

## Documentation fixes

- VBA_CODEBASE.md: corrected v4 byte-identity claim, empty-module list (added CDJ/GJ), Sheet9 constants, Sundry T/U numeric vs nonnumeric, sync ErrorHandler boundary, stale compile PASS, and v4 same-logic wording.
- VBA_CODE_FLOW.md: corrected SJ/SJ2 layouts, SJ2 formulas, SJ stamp conditions, CDJ day field, fixed CDJ/CRJ maps, hard-coded aliases, active non-VAT path wording, and sync error path.

## Inventory summary

- 15 non-empty components: 8 standard modules, 6 worksheet modules, ThisWorkbook.
- 146 declared procedures; 7 event handlers.
- Entry points: ThisWorkbook.Workbook_Open; Sheet12/Sheet11/Sheet9/Sheet3/Sheet8 Worksheet_Change; Sheet3 Worksheet_Activate.
- No UserForms, Class modules, Auto_*, OnTime, named ranges, external file I/O, or Workbooks.Open in canonical source.
- Option Explicit missing in Sheet9.txt and Sheet10.txt (held/documented).

## Static checks

### check_vba_static.py

Run on pre-fix export (commit 0589321): 18 findings, including the C-01 duplicate `Dim n` in ValidateV8Structure, missing Option Explicit in empty/document modules, and On Error Resume Next not reset in SheetExists/Graphify/Legacy_RefreshGL.

Run on post-fix export (build/v8_5_vba/): 17 findings. C-01 duplicate is fixed. Remaining are mostly missing Option Explicit in empty/document modules and conservative line-based flags for On Error Resume Next in SheetExists, Graphify, and Legacy_RefreshGL. Legacy_RefreshGL is not on the active ScanGL path.

### check_docs.py

- Documented procedure names checked against exports: no real procedure absence after filtering common VBA/object-model identifiers.
- 56 Public procedures/event handlers are not referenced in docs; these are mostly internal helpers. Docs cover the main entry points; documenting internal helpers is a future improvement.

## Verification after each code commit

- Python byte-level check confirmed CRLF line endings, no BOM, and \r\n count == \n count for every edited file.
- git diff reviewed for each commit; only intended lines changed.
- python tests/test_vba_scope_guard.py: OK (5 tests). Note: this test targets vba/ (stale historical copies), not build/v8_5_vba/, per the existing test file.

## Manual checklist for reviewer

- [ ] Debug > Compile VBAProject in live Excel before trusting any runtime behaviour.
- [ ] Sheet*.txt are document modules: they cannot be brought in with File > Import; their code must be pasted into each sheet's module.
- [ ] Test on a COPY of the workbook; save as a new version.
- [ ] Note that vba/*.bas are stale copies and ask me whether to delete them.
- [ ] Answer held questions in docs/audit/blocked.md before enabling C-05/C-06/C-10.

## Next steps after review

1. Review and run the manual checklist above, including live Excel compile.
2. Answer blocked.md questions or keep C-05/C-06/C-10 held.
3. Optionally document internal Public helpers flagged by check_docs.py.
4. When ready, merge/push audit/vba-docs-sync as agreed.
