# Initial OpenCode Audit Report

Date: 2026-11-11
Branch: `audit/vba-docs-sync`
Base commit: `2f44164` (`docs: add full VBA codebase and flow analysis`)
Status: **Phases 0-3 and Step 3 complete. Committed and pushed to origin/audit/vba-docs-sync at `1a7cde6`.**

## 1. Executive summary

The audit compared `VBA_CODEBASE.md` and `VBA_CODE_FLOW.md` against the exported v8.5 VBA in `build/v8_5_vba/`.

- The documentation is substantially accurate: all 146 procedures, all 15 full-source reproductions, module line counts, and the workbook SHA-256 were verified.
- Most material claims were confirmed: 33 of 48 checked claims are fully `VERIFIED`.
- 8 claims are `WRONG` and 7 are `PARTIAL`, mainly around the SJ/SJ2 layouts, non-VAT CDJ fields, error handling, and the compile-pass claim.
- 28 findings were recorded: **1 Critical, 4 High, 13 Medium, 9 Low, 1 Info** (15 code, 13 documentation).
- The most important code defect was a duplicate local declaration in `modGLGate.ValidateV8Structure` (`modGLGate.bas:13,29`), fixed by renaming the counter to `nonblank`.
- Exported VBA fixes were committed to `build/v8_5_vba/`; no `.xlsm`/`.xlsb`/`.xls`/`vbaProject.bin` was modified. No live Excel compile or run was performed.

## 2. Scope and sources

| Path | Role |
|---|---|
| `VBA_CODEBASE.md` | Codebase documentation under audit |
| `VBA_CODE_FLOW.md` | Flow/documentation under audit |
| `build/v8_5_vba/` | Canonical exported VBA (15 non-empty components) |
| `build/Global-Smile_2026-v8_5.xlsm` | Workbook provenance |
| `build/artifacts/compile.log` | Recorded compile artifact; found to be stale |
| `build/verify_summary.json` | Build summary and workbook hash |
| `vba/` | Older/tracked GL copies; not canonical |
| `DRR_deep-research-report.md` | Untracked outside review; hypotheses only |

Workbook SHA-256 verified as:

```text
D35BF20944955697008216E23E0D112B1C24EF297EB6AC991454011C7FF15234
```

The `DRR_deep-research-report.md` report states that it had no repository/code access. Its claims were therefore treated as hypotheses and were not promoted without source evidence.

## 3. Work completed

### Phase 0 - Files and provenance

Recorded canonical source, duplicate/older source locations, workbook hash, documentation provenance, the pre-existing dirty working tree, and the stale compile artifact.

Output: `docs/audit/00_files.md`.

### Phase 1 - Inventory

Parsed all exported components and declarations from the canonical text.

| Measure | Result |
|---|---:|
| Non-empty components | 15 |
| Standard modules | 8 |
| Worksheet modules | 6 |
| `ThisWorkbook` | 1 |
| Declared procedures | 146 |
| Event handlers | 7 |

The component totals match the documentation's Section 3 table.

Output: `docs/audit/01_inventory.md`.

### Phase 2 - Claim verification

Checked 48 material claims. The procedure index and full-source reproductions were checked as bulk claims.

| Status | Count |
|---|---:|
| VERIFIED | 33 |
| PARTIAL | 7 |
| WRONG | 8 |
| **Total** | **48** |

Verified in full:

- Workbook SHA-256 and component/procedure counts.
- Every procedure-index line and declaration.
- All 15 full-source code reproductions.
- The main active flow: PJ/PNV to CDJ, SJ to CRJ, and gated GL write-back.
- CDJ/CRJ placement, reference lookup, supplier lookup, dependency matrix, and most GL scan details.

Wrong or incomplete claims:

| Claim | Problem |
|---|---|
| `vba/` is byte-identical except the export header | Also differs by UTF-8 BOM and trailing blank lines; logic is otherwise the same |
| Sync write errors are swallowed by empty handlers | `On Error GoTo 0` disables the handlers before writes; later errors propagate |
| Both PJ syncs write the CDJ day field | `SyncPJNonVatToCDJ` never writes `CDJ_COL_DAY` |
| SJ invoice stamping requires date + amount | Sheet3 checks amount/TOTAL only; Sheet8 checks date too |
| SJ2 uses the same sales formulas as SJ | Sheet8 writes only M/O; Sheet3 also writes L/N mirrors |
| The SJ/SJ2 layout table is shared | Sheet8 column I is auto invoice and U is reference; Sheet3 I is manual physical invoice, U is auto invoice, V is full ID |
| The `compile PASS` artifact validates the build | The log predates `modGLGate.bas` and the workbook; the current export has a compile error |
| Documented error path says sync writes fail silently | Same root issue as the false ErrorHandler claim |

Output: `docs/audit/02_claims.md`.

### Phase 3 - Findings

Recorded 27 findings with severity, evidence, and proposed phase.

| Severity | Count |
|---|---:|
| Critical | 1 |
| High | 4 |
| Medium | 12 |
| Low | 9 |
| Info | 1 |
| **Total** | **27** |

By type:

| Type | Count |
|---|---:|
| Code | 14 |
| Documentation | 13 |

Output: `docs/audit/03_findings.md`.

## 4. Critical and high findings

### C-01 - Critical: duplicate local declaration

`modGLGate.ValidateV8Structure` declares `n` twice in the same procedure:

- `modGLGate.bas:13`: `Dim names As Variant, n As Variant, ws As Worksheet`
- `modGLGate.bas:29`: `Dim b As Long, r As Long, n As Long`

This is a compile-time duplicate declaration under standard VBA scoping rules. The recorded `compile.log` is not valid evidence for the current export because it predates this file.

Proposed fix: keep `n` as the `For Each` Variant and rename the second counter, for example `nonblank`.

### C-02 - High: initialization failures are hidden

`InitializeEngine` runs every loader under `On Error Resume Next` and sets `gInitialized = True` regardless of whether the loaders succeeded (`modEngine.bas:130-142`). Callers can then use partial or empty supplier/reference caches.

Proposed fix: fail loudly or record a load status; set the initialized flag only after successful loads.

### C-03 - High: `ScanGL` has no top-level error path

`ScanGL` calls `ScanCDJ` and `ScanCRJ` without a top-level error handler (`modGLScanEngine.bas:20-68`). An unexpected runtime error can bypass the `GL_AUDIT` HOLD record.

Proposed fix: add a top-level handler that attempts to write a HOLD audit and re-raises.

### D-01 - High: documentation describes the wrong sync error behavior

`VBA_CODEBASE.md:92` and `VBA_CODE_FLOW.md:895-899` say empty `ErrorHandler:` labels swallow write errors. In fact, `On Error GoTo 0` at `modEngine.bas:851`, `:978`, `:1249` disables those handlers before the write bodies. Write errors after that point propagate.

Proposed fix: correct both documents and describe the pre-write/pointer-update boundary accurately.

### D-02 - High: stale compile-pass claim

`VBA_CODEBASE.md:99` and `:5779` present a recorded compile PASS as current evidence. The log timestamp is `10/30/2026 5:42:46 PM`; `modGLGate.bas` is `6:28:12 PM` and the workbook is `6:25:58 PM`. The log predates the current source and cannot validate it.

Proposed fix: label the artifact stale and state that the current export contains C-01.

## 5. Other notable findings

- `SyncPJNonVatToCDJ` omits `CDJ_COL_DAY` on both update and insert (`modEngine.bas:1267-1286,1299-1319`).
- `SyncSJToCRJ` converts amount to `Double` before `IsNumeric`, so a text amount raises a type mismatch (`modEngine.bas:985,990`).
- `Sheet9` uses undeclared `START_ROW` and `COL_*` constants without `Option Explicit`; the legacy path is wired but unreliable.
- `Sheet11`, `Sheet12`, and `Sheet8` cleanup handlers restore application state but swallow the original error; Sheet11/Sheet12 can continue into Part 2.
- `ImportSuppliersFromJSON` clears rows before validating and only then checks duplicates, with no rollback.
- `modPJSync.LookupByTIN_NonVat` writes the first match before detecting a duplicate and exits on the second match, leaving a partial supplier row.
- `RefreshGL` wrappers do not call `ValidateV8Structure`; only `RunV8Gate` validates the workbook structure first.
- `WriteScanAudit` uses `On Error Resume Next` around sheet lookup/add and can reach an error 91 if the add fails.
- Several low-severity state-restoration and message-cell issues remain in `ResetCDJ`, `ResetCRJ`, and `CDJReserveRow`.

## 6. Evidence boundary and blockers

- The workbook was not opened, compiled, or run. The embedded VBA/source boundary remains unverified.
- `build/artifacts/compile.log` is stale; no live Excel compile result exists for the current export.
- `DRR_deep-research-report.md` is untracked and outside the reviewed code path; it remains a hypothesis source.
- No Mermaid CLI rendering was run; diagrams were checked statically.
- Pre-existing unrelated working-tree changes remain untouched, including the deleted `Global-Smile_2026-v8.4.1.xlsm` and untracked `build/`, `dist/`, `patch/`, workbook images, and the DRR.

## 7. Files created

| File | Purpose |
|---|---|
| `docs/audit/00_files.md` | Provenance and file inventory |
| `docs/audit/01_inventory.md` | Components, procedures, event handlers, patterns |
| `docs/audit/02_claims.md` | Verified/partial/wrong claim matrix |
| `docs/audit/03_findings.md` | Defect register with severity and proposed phase |
| `docs/audit/AUDIT_STATE.md` | Phase status and count summary |
| `initial_opencode_report.md` | This report |

## 8. Recommended next steps

1. **Phase 4 - Documentation fixes:** correct D-01 through D-13, including the sync-error text, stale compile claim, SJ/SJ2 layouts, CDJ day field, map tables, and `vba/` byte-identity wording.
2. **Phase 5 - VBA source fixes:** fix C-01 first, then decide and implement C-02, C-03, C-04, C-05, C-08, C-09, and C-10; apply the mechanical fixes C-06, C-07, C-11, C-12, C-13, and C-14.
3. **Phase 6 - Checker and report:** add `scripts/check_docs.py`, re-run a live Excel compile/scan where possible, and produce the final audit report.

No Phase 4-6 changes have been made. Work is stopped for approval.
