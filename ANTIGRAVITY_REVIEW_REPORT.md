# VBA Codebase Review Report — Global-Smile 2026

**Prepared by:** Antigravity  
**Date:** 2026-11-14  
**Branch:** `audit/vba-docs-sync`  
**Workbook:** `Global-Smile_2026-v8.4.1-with-BUG.xlsm` / build target `v8.5`  
**Status:** `HOLD` — fail-closed gate working as designed; three blocking issues remain before a live GL scan can PASS.

---

## 1. Scope

This report covers a full source-level read of:

- `vba/` — four GL-layer modules (`modGLGate.bas`, `modGLRefresh.bas`, `modGLScanEngine.bas`, `modGLWorkbookMap.bas`)
- `build/v8_5_vba/` — complete 15-component VBA export (8 standard modules, 6 worksheet modules, `ThisWorkbook`)
- Supporting files: `VBA_CODEBASE.md`, `VBA_CODE_FLOW.md`, `GL_REBUILD_VERIFICATION_REPORT.md`, `CLAUDE_REPORT.md`, `build/test_results.txt`, `build/tests_modtest.bas`, `tests/test_vba_scope_guard.py`

No live Excel compile or macro run was performed in this session. Runtime evidence cited is from the recorded v8.4.1 run (499 ms, `GL_AUDIT` row 9).

---

## 2. Architecture Summary

The workbook VBA is organised in three layers:

| Layer | Modules | Responsibility |
|---|---|---|
| **Event / Input** | `Sheet3`, `Sheet8`, `Sheet9`, `Sheet11`, `Sheet12`, `ThisWorkbook` | Capture user edits; call posting engine |
| **Posting Engine** | `modEngine` (76 procs), `modPJSync` (8 procs) | Supplier cache, CDJ/CRJ row placement, sequence |
| **GL Scan** | `modGLGate`, `modGLRefresh`, `modGLScanEngine`, `modGLWorkbookMap` | Read CDJ+CRJ, validate, write GL only after gate passes |

Active accounting path:

```
Workbook_Open → modEngine.InitializeEngine
PJ/SJ edits   → modEngine.SyncPJToCDJ / SyncSJToCRJ → CDJ / CRJ
RefreshGL     → modGLGate.RunV8Gate → modGLScanEngine.ScanGL → GL + GL_AUDIT
```

GJ is explicitly excluded from scope. No `GL_V8_CALC` sheet is created by the active path.

---

## 3. GL Scan Engine — Key Facts

| Property | Value |
|---|---|
| GL blocks | 46 account titles |
| Months per block | 12 (Jan–Dec) |
| Write cells per block | G (debit), H (credit), I (cumulative balance) |
| CDJ columns mapped | F–R (13 fixed accounts) + col S (Sundry, COA-resolved) |
| CRJ columns mapped | H–L (4 accounts: Cash, Excess CWT, VAT Payable, Sales) |
| Debit/credit tolerance | ₱0.01 |
| Fail-closed position | Gate fires **before** `WriteScannedGL`; audit row is always written |

### CDJ Column → GL Account Map

| Column | Account |
|---|---|
| F (6) | Cash in Bank |
| G (7) | Input VAT |
| H (8) | EWT Payable |
| I (9) | Petty Cash Fund |
| J (10) | Government Contributions (EE) |
| K (11) | Government Loans (EE) |
| L (12) | De Minimis |
| M (13) | Salaries and Wages |
| N (14) | Clinic Material and Supplies |
| O (15) | Rent |
| P (16) | Gas, Oil, Parking, Toll Fees |
| Q (17) | Light and Water Expense |
| R (18) | Professional Fees |
| S/T/U (19–21) | Sundry Account (resolved via COA_MAP / V8_ResolveAccount) |

---

## 4. Defect Register

### 4.1 Critical

| ID | Module | Line(s) | Finding |
|---|---|---|---|
| **C-01** | `modGLGate.bas` | 13, 29 | `n` declared twice — once as `Variant` (loop var), once as `Long` (nonblank counter). **Compile error.** Fixed in `build/v8_5_vba/modGLGate.bas` (counter renamed to `nonblank`). The `vba/` copy is **not fixed**. Must verify which copy is embedded in the .xlsm. |

### 4.2 High

| ID | Module | Line(s) | Finding |
|---|---|---|---|
| **H-01** | `modGLWorkbookMap.bas` | 75 | `InStr(st, "MAPPED") > 0` matches `"UNMAPPED"`. Rows explicitly marked UNMAPPED get aliased. Fix: `st = "MAPPED"` (exact equality). |
| **H-02** | `modGLWorkbookMap.bas` | 131–143 | 10 sundry labels are hard-coded as executable aliases (`MEDICAL EQUIPMENT → Dental Equipment`, `COST OF REVENUE → Clinic Material and Supplies`, etc.). README states these are not approved. Policy/implementation mismatch. |
| **H-03** | `modGLScanEngine.bas` | 82, 169 | `yr` parameter accepted but never used. Scan aggregates all matching rows by month regardless of year. Safe only for single-year workbooks with correct January-first CRJ boundaries. |
| **H-04** | `Sheet9.txt` | entire module | No `Option Explicit`. Nine column constants (`COL_TIN`, `COL_NAME`, `COL_AMOUNT`, `COL_DEBIT`, `COL_REF`, `COL_COA`, `COL_FULL_ID`, `COL_DEBIT_MIRROR`, `START_ROW`) are undeclared Variants → resolve to 0/Empty. PJ Non-Vat2 path is non-functional at runtime. |

### 4.3 Medium

| ID | Module | Finding |
|---|---|---|
| **M-01** | `modGLScanEngine.bas:293` | `bal = 0#` resets for every block — cumulative balance is within-year only; no roll-forward from prior-year closing balance. |
| **M-02** | `modEngine` sync routines | Empty `ErrorHandler:` label after `On Error GoTo 0` — CDJ/CRJ write errors propagate silently. |
| **M-03** | `Sheet11` / `Sheet12` | Duplicate implementations of 8 procedures. Bug fixes must be applied to both sheets. |
| **M-04** | `modEngine` | Legacy dead code compiled but unreachable: `GLScanGJ`, `Legacy_RefreshGL`, `BuildV8CalcSheet`, etc. |
| **M-05** | `modEngine.SyncPJNonVatToCDJ` | Does not write `CDJ_COL_DAY`. Day column is blank for Non-Vat rows. |
| **M-06** | `Sheet3` / `Sheet8` | Different invoice-key formats (`SJ-#####` vs raw Long). CRJ row identification may misfire between sheets. |

### 4.4 Low / Info

| ID | Finding |
|---|---|
| **L-01** | `ImportSuppliersFromJSON` clears SUPPLIERS DATA before inserting hard-coded records. Accidental execution destroys live data. |
| **L-02** | `WriteScanAudit` truncates `detail` to 32,000 chars. Long error runs lose diagnostic tail. |
| **L-03** | `modGLRefresh` comment references `GL_V8_CALC` which no longer exists on the active path. |
| **L-04** | `build/artifacts/compile.log` timestamp (17:42) predates `modGLGate.bas` (18:28) and the workbook (18:25). Cannot be used as proof of compile pass. |

---

## 5. Live Runtime Evidence (v8.4.1, 2026-10-30)

```
Duration : 499 ms
Status   : HOLD
Debit    : ₱8,429,543.41
Credit   : ₱9,317,172.68
Diff     : −₱887,629.26   ← HOLD trigger
Unmapped : ₱1,091,293.17  ← HOLD trigger
Invalid  : 1              ← HOLD trigger
Detail   : GL!F221 duplicate account title 'Paid-Up Capital' (first seen at GL!F208)
```

The fail-closed gate worked correctly: GL was not overwritten.

### HOLD Root Causes

1. **GL!F221 — Duplicate account title.** `Paid-Up Capital` appears in blocks at rows F208 and F221. One block must be renamed to its correct distinct account name.
2. **10 unmapped CDJ sundry labels (₱1,091,293.17).** Each label needs an explicit accounting decision: set COA_MAP column D = GL account name, column F = `MAPPED`.
3. **Debit/credit imbalance (₱887,629.26).** Downstream of #2 — unmapped amounts are excluded, creating the gap. Resolves automatically when #2 is fully approved.

---

## 6. Source-Level Test Results

### `tests/test_vba_scope_guard.py` (5 assertions, targets `vba/`)

| Test | Result |
|---|---|
| Scan engine does not read/scan GJ | ✅ PASS |
| Scan does not write `GL_V8_CALC` calc sheet | ✅ PASS |
| Account uniqueness guard runs before write-back | ✅ PASS |
| Resolver is not named for excluded journal | ✅ PASS |
| Old mapping writer (`apply_full_fix.ps1`) is disabled | ✅ PASS |

### `build/tests_modtest.bas` → `build/test_results.txt` (10 cases, run in Excel)

```
passed=10  failed=0  missing=0
```

All supplier engine tests pass against the v8.4.1 workbook.

---

## 7. Version Boundary Warning

> The `vba/` directory contains **v8.4.1** copies of the GL modules.  
> The `build/v8_5_vba/` directory contains the **v8.5** export.  
> The embedded VBA inside the .xlsm files has **not been confirmed** to match either copy.  
> Always verify by exporting from the target workbook before claiming a fix is live.

| File | `vba/` | `build/v8_5_vba/` |
|---|---|---|
| `modGLGate.bas` | Duplicate `n` — **C-01 unfixed** | `nonblank` — C-01 fixed |
| `modGLScanEngine.bas` | No top-level error handler | Has error handler (C-03 fixed) |
| `modGLWorkbookMap.bas` | Same | Same |
| `modGLRefresh.bas` | Same | Same |

---

## 8. Testing Workflow

### Tier 1 — Source-Level (No Excel Required)

```powershell
cd D:\citrixlabph\globalsmile_v2
python -m unittest discover -s tests -v
```

Expected: 5 PASS. Any FAIL → do not proceed.

Additional checks recommended:
- Verify `modGLGate.bas` has no duplicate `Dim n` (C-01).
- Verify `modGLWorkbookMap.bas:75` uses exact `st = "MAPPED"` (H-01).

### Tier 2 — Workbook-Level (Excel Desktop, Work on a Copy)

1. **Compile:** `Alt+F11` → `Debug` → `Compile VBAProject` — expect zero errors.
2. **Import modules:** Use `build/v8_5_vba/` for standard modules; paste document module code manually.
3. **Supplier engine:** `Alt+F8` → `RunAll` from `tests_modtest.bas` → `build/test_results.txt` should show 10 PASS.
4. **Fix GL title:** Rename duplicate `Paid-Up Capital` at GL!F221.
5. **Fix COA_MAP:** Get accounting sign-off on each sundry label; set status to `MAPPED` (exact).
6. **Fix H-01:** Change `InStr(st, "MAPPED")` to `st = "MAPPED"` in `modGLWorkbookMap.bas:75`; re-import.
7. **Run RefreshGL:** `Alt+F8` → `RefreshGL` → check `GL_AUDIT` for PASS row.

### Tier 3 — Acceptance Gates

| Gate | Required Evidence |
|---|---|
| Compile | Zero errors in live Excel compile |
| Scope | GJ not read — Tier 1 PASS |
| GL structure | 46 nonblank unique titles — `GL_AUDIT: Invalid=0` |
| Mapping | Every sundry label has accounting sign-off — COA_MAP reviewed |
| Reconciliation | `GL_AUDIT: Difference=0.00, Unmapped=0.00` |
| Roll-forward | Manual spot-check: 3 accounts × 12 months |
| Supplier engine | `test_results.txt: passed=10 failed=0` |

---

## 9. Fix Priority Queue

| Priority | ID | Action |
|---|---|---|
| 🔴 P1 | C-01 | Rename duplicate `Dim n` to `nonblank` in `modGLGate.bas` — required before workbook compiles |
| 🔴 P2 | — | Rename duplicate `Paid-Up Capital` title in GL!F221 |
| 🟠 P3 | H-01 | Change `InStr(st, "MAPPED")` to `st = "MAPPED"` in `modGLWorkbookMap.bas:75` |
| 🟠 P4 | H-02 | Obtain accounting sign-off on 10 COA_MAP sundry labels; update COA_MAP |
| 🟡 P5 | H-03 | Add year enforcement to `ScanCDJ` and `ScanCRJ` if multi-year scope is ever needed |
| 🟡 P6 | H-04 | Confirm Sheet9 column layout then add `Option Explicit` and constants (held: C-05) |
| 🟢 P7 | M-04 | Remove or comment-out legacy `GLScanGJ`, `BuildV8CalcSheet` etc. in `modEngine` |

---

## 10. Held Items (External Decision Required)

| ID | Blocked On |
|---|---|
| **C-05** | Confirm actual column numbers for Sheet9 `COL_TIN`, `COL_NAME` etc. before declaring constants. |
| **C-06** | Confirm which CDJ column holds the day value for Non-Vat rows before writing `CDJ_COL_DAY`. |
| **C-10** | Confirm whether `LookupByTIN_NonVat` partial-write should rollback or leave partial state. |
| **Sundry approvals** | 10 CDJ sundry labels need formal accounting authorization before `MAPPED` status can be set. |

---

*Report generated by **Antigravity** on 2026-11-14 at 15:55 PHT.*  
*Source-level analysis only. No live Excel compile or macro run was performed in this session.*
