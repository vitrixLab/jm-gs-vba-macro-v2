# Global-Smile 2026 — General Ledger Automation v2

Clean, verified repository for Global-Smile 2026 General Ledger (GL) automation rebuild.

Repository: `https://github.com/vitrixLab/jm-gs-vba-macro-v2.git`

---

## 1. Quick Start

### Running in Excel Desktop
1. Open [`Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm`](./Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm) in Microsoft Excel with macros enabled.
2. Press **Alt + F8**, select `RefreshGL`, and click **Run**.
3. The macro processes CDJ and CRJ in < 1.0 second, confirms the gate **PASS**, and writes all 46 accounts × 12 months directly into the `GL` sheet.

### Running Headless / Automation (PowerShell)
```powershell
powershell -ExecutionPolicy Bypass -File scripts/run_refresh_gl.ps1
```

---

## 2. Validation & Gate Status (`GL_AUDIT`)

The fail-closed gate in `modGLScanEngine.ScanGL` evaluates to **PASS**:

| Metric | Result | Audit Gate Criteria | Status |
|---|---|---|---|
| **Gate Status** | **PASS** | `|gd - gc| <= 0.01 AND um <= 0.01 AND inv = 0` | **PASSED** |
| **Total Debit Scanned** | **₱9,317,172.67** | CDJ Allocations + CRJ Collections | Verified |
| **Total Credit Scanned** | **₱9,317,172.68** | CDJ Cash/Withholdings + CRJ Revenue/VAT | Verified |
| **Net Difference** | **-₱0.01** | `|diff| <= 0.01` | **PASSED** |
| **Unmapped Activity** | **₱0.00** | All 10 sundry expense labels resolved | **PASSED** |
| **Invalid Records** | **0** | No missing monthly partitions or corrupt rows | **PASSED** |
| **GL Coverage** | **552 rows** | 46 accounts × 12 months | Complete |

See [`GL_REBUILD_VERIFICATION_REPORT.md`](./GL_REBUILD_VERIFICATION_REPORT.md) for full 12-month reconciliation tables and account-by-account roll-forward proofs.

---

## 3. Architecture & Strict Scope

- **Inputs:** CDJ (Cash Disbursement Journal) and CRJ (Cash Receipt Journal) only.
- **Excluded:** GJ (General Journal) is completely excluded from scan and posting.
- **Output:** The existing `GL` sheet. No secondary calculation sheets or deleted `V*GL` sheets are used.
- **Accounts:** All 46 canonical GL accounts (stride 13, rows 13 to 598).
- **Ending Balance:** Rolled forward strictly as:
  $$\text{Ending Balance}_m = \text{Ending Balance}_{m-1} + \text{Debit}_m - \text{Credit}_m$$

---

## 4. Key Fixes Applied

1. **COA_MAP Reader Fix:** Corrected loop termination in `modGLWorkbookMap.EnsureMaps` so all rows through `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` are loaded into memory.
2. **Authorized Sundry Mappings (Rows 13–22 in COA_MAP):**
   - `Representation` → `Meals, Gifts & Entertainment Expense`
   - `Visiting Clinicians Fee` → `Clinician's Fee-Visiting`
   - `Partners Clinicians Fee` → `Clinician's Fee-Corporators`
   - `Clinicians Fee- Corporatos` → `Clinician's Fee-Corporators`
   - `Clinacians Fee- Corporators` → `Clinician's Fee-Corporators`
   - `Partners Clinicians Fee- Corporators` → `Clinician's Fee-Corporators`
   - `SSS,HDMF AND PHIC ER SHARE` → `Government Contributions ER Share`
   - `Incentives` → `De Minimis`
   - `Outside Services` → `Common Use Service Area`
   - `Training and Seminar` → `Miscellaneous`
3. **ScanGJ Suppression:** Suppressed `ScanGJ` in `modGLScanEngine.ScanGL` per strict scope boundaries.
4. **No Extra Calculation Sheets:** Suppressed `WriteScannedCalc` in `modGLScanEngine.ScanGL` to prevent workbook bloat.
5. **Headless Execution Guard:** Wrapped `MsgBox` in `If Application.UserControl And Application.DisplayAlerts Then` to allow automated COM execution without modal deadlocks.
6. **String Length Safety:** Added `Left$(detail, 32000)` safeguard in `WriteScanAudit` to prevent runtime error 1004.
7. **GL Column 9 Visibility:** Unhidden with width 16 and header set to `Ending Balance`.

---

## 5. Repository Structure

- `Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm`: Certified working Excel workbook.
- `Global-Smile_2026-v8.4-NO-SHIT.xlsm`: Clean baseline workbook.
- `GL_REBUILD_VERIFICATION_REPORT.md`: Comprehensive audit and reconciliation report.
- `vba/`: Clean extracted VBA source files (`modGLScanEngine.bas`, `modGLWorkbookMap.bas`, `modGLRefresh.bas`, `modGLGate.bas`).
- `scripts/`: Verification and runner scripts (`verify_full_gl.py`, `reconcile_sources.py`, `run_refresh_gl.ps1`, `apply_full_fix.ps1`).
