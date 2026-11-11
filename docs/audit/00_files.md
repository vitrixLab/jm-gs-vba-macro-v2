# Phase 0 - Files and provenance

Audit date: 2026-11-11
Branch: `audit/vba-docs-sync`, based on `2f44164` (`docs: add full VBA codebase and flow analysis`).
Scope: read-only audit of the exported VBA and the two documentation files. No workbook was opened, compiled, or run.

## Audit target

| Path | Role | Evidence |
|---|---|---|
| `VBA_CODEBASE.md` | Primary codebase/documentation claim source | 5,793 lines; commit `2f44164` |
| `VBA_CODE_FLOW.md` | Primary flow/diagram claim source | 1,117 lines; commit `2f44164` |
| `build/v8_5_vba/` | Canonical exported VBA for v8.5 | 25 files, 15 non-empty |
| `build/Global-Smile_2026-v8_5.xlsm` | Workbook provenance | SHA-256 `D35BF20944955697008216E23E0D112B1C24EF297EB6AC991454011C7FF15234` |
| `DRR_deep-research-report.md` | Outside review; untracked | 43 lines; report states it had no repository/code access |

The workbook hash matches the hash recorded in `VBA_CODEBASE.md:5` and `build/verify_summary.json`.

## Canonical VBA files

| File | Lines | Nonblank | Declared procedures |
|---|---:|---:|---:|
| `ThisWorkbook.txt` | 10 | 8 | 1 |
| `Sheet12.txt` | 290 | 255 | 9 |
| `Sheet11.txt` | 294 | 260 | 9 |
| `Sheet9.txt` | 295 | 204 | 1 |
| `Sheet3.txt` | 175 | 147 | 6 |
| `Sheet8.txt` | 88 | 74 | 2 |
| `Sheet10.txt` | 34 | 31 | 1 |
| `modEngine.bas` | 1,781 | 1,668 | 76 |
| `modPJSync.bas` | 532 | 257 | 8 |
| `modGLGate.bas` | 51 | 43 | 3 |
| `modGLRefresh.bas` | 28 | 22 | 4 |
| `modGLScanEngine.bas` | 335 | 292 | 9 |
| `modGLWorkbookMap.bas` | 211 | 189 | 12 |
| `modGraphify.bas` | 229 | 201 | 4 |
| `ImportSuppliersFromJSON.bas` | 856 | 741 | 1 |
| **Total** | | | **146** |

All 15 files use CRLF and have no BOM in `build/v8_5_vba/`. Non-ASCII bytes occur in some source comments/strings; edits must preserve the existing bytes.

## Other source copies

| Location | Contents | Audit treatment |
|---|---|---|
| `vba/` | Four GL modules: `modGLGate`, `modGLRefresh`, `modGLScanEngine`, `modGLWorkbookMap` | Older/tracked copies; not canonical |
| `build/baseline_vba/` | Baseline export used by the v8.5 build diff | Historical |
| `build/post_vba/` | Post-patch `Sheet11`/`Sheet12` exports | Historical |
| `patch/claude code/` | v7.9.5 `modEngine` and PJ patch | Historical |
| `_inspect/all_vba.txt` | Older single-file VBA dump | Superseded |

`vba/` is not byte-identical to `build/v8_5_vba/`. The differences include the exported `Attribute VB_Name` header, a UTF-8 BOM on the `vba/` copies, and trailing blank lines. The body logic is otherwise the same after those differences are removed.

## Documentation provenance

- Commit `2f44164` added only `VBA_CODEBASE.md` and `VBA_CODE_FLOW.md`. It did not add VBA source or `DRR_deep-research-report.md`.
- `DRR_deep-research-report.md` is untracked. Its own opening states that the review had no repository access, so its defect list is a hypothesis list, not evidence.
- The docs correctly state that the snapshot does not certify that the embedded project compiles or matches the exported text (`VBA_CODEBASE.md:9`).
- The recorded compile artifact is `build/artifacts/compile.log:1` (`Fri Oct 30 17:42:46 2026: compile PASS`). Its mtime is 17:42; `modGLGate.bas` has mtime 18:28 and the workbook has mtime 18:25. The artifact predates the current `modGLGate` export and cannot validate it.

## Pre-existing working-tree state

The branch was created with these unrelated changes already present; they were left untouched:

```text
 D Global-Smile_2026-v8.4.1.xlsm
?? 461-500 supplier info bug.JPG
?? DRR_deep-research-report.md
?? Global-Smile_2026-v8.4.1-Mark-Wps-5.xlsm
?? Global-Smile_2026-v8.4.1-REVIEW.xlsm
?? Global-Smile_2026-v8.4.1-with-BUG.xlsm
?? build/
?? dist/
?? patch/
```

No audit files existed when Phase 0 started; `docs/audit/` contained only the directory.
