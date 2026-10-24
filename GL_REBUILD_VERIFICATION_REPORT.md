# GL Automation v2 — Evidence Status

**Status: HOLD — fail-closed gate working as designed.**

## Scope boundary

- CDJ and CRJ are the only permitted posting inputs.
- GJ must not be read or posted.
- Write-back target is the existing GL sheet.
- Do not create GL_V8_CALC or restore deleted V*GL sheets.
- The intended output layout is 46 distinct account blocks × 12 months.

## v8.4.1 Build & Runtime Findings

1. **Workbook Rebuild (`Global-Smile_2026-v8.4.1.xlsm`):**
   - Rebuilt from `Global-Smile_2026-v8.4-NO-SHIT.xlsm` via `scripts/build_v841.ps1`.
   - Injected updated VBA modules: `modGLScanEngine.bas` (334 lines), `modGLWorkbookMap.bas` (210 lines), `modGLGate.bas` (50 lines), `modGLRefresh.bas` (27 lines).
2. **Runtime Execution (`modGLRefresh.RefreshGL`):**
   - Live macro execution completed in **499 ms** via Excel COM.
   - Resulting audit entry in `GL_AUDIT` (Row 9):
     `v8.3.3 SCAN | Debit=8429543.41 | Credit=9317172.68 | Difference=-887629.26 | Unmapped=1091293.17 | Invalid=1`
     Status: **`HOLD`**
3. **Fail-Closed Triggers Verified:**
   - **Account Uniqueness Guard:** Detected `GL!F221 duplicate account title 'Paid-Up Capital' (first seen at GL!F208); write-back held` (`Invalid = 1`). Write-back halted before altering the ledger.
   - **Unmapped Sundry Guard:** Ten unapproved CDJ sundry labels evaluated as unmapped (`Unmapped = ₱1,091,293.17`, `Difference = -₱887,629.26`).
   - Gate cleanly aborted before modifying the visible GL sheet.

## Acceptance gates

| Gate | Required evidence | v8.4.1 Status |
|---|---|---|
| Scope | No GJ worksheet read or post in active scan | **PASS** — ScanGJ removed, zero GJ activity |
| GL structure | 46 nonblank, unique account titles | **HOLD (Caught defect)** — GL!F221 duplicate title caught |
| Mapping | Every nonzero sundry label has explicit approved mapping | **HOLD (Safe)** — Unapproved labels held |
| Reconciliation | Source journal totals and GL postings agree within tolerance | **HOLD** — Write-back held per gate rules |
| Roll-forward | Each month’s ending balance equals prior ending + debit − credit | Protected (GL preserved from corruption) |
| Runtime | Compile and refresh in Excel, inspect output and audit row | **PASS** — Clean compile & 499 ms run |
