# Global-Smile 2026 - Complete VBA Codebase

> Static snapshot for deep code analysis. Generated 2026-11-11 18:15 from the exported VBA of `build/Global-Smile_2026-v8_5.xlsm`.
>
> Workbook SHA-256: `d35bf20944955697008216e23e0d112b1c24ef297eb6ac991454011c7ff15234`
>
> Source folder: `build/v8_5_vba/` (one file per VBA component)
>
> This document reproduces the code; it does not certify that the embedded project compiles, matches this text byte-for-byte, or passes a live refresh. See the repository README and `GL_REBUILD_VERIFICATION_REPORT.md` for the evidence boundary.

## 1. Scope and provenance

The current workbook is the v8.5 build. Its VBA project contains 15 non-empty components (8 standard modules, 6 worksheet modules and `ThisWorkbook`) plus 10 empty worksheet modules.

Included here:

- Every non-empty VBA component exported to `build/v8_5_vba/`.
- Exact procedure declarations, module-level state and all executable code.
- Workbook codename and host-sheet mapping recovered from `xl/workbook.xml` and the worksheet `codeName` attributes.
- A procedure index with source line numbers and a static-analysis notes section.

Not reproduced (to avoid duplicate/older code misleading the analysis):

- `vba/` contains the v8.4.1 copies of the four GL modules. They are byte-identical to the v8.5 export except for the exported `Attribute VB_Name` header line, so this document uses the v8.5 copies.
- `build/baseline_vba/` and `patch/claude code/` contain baseline/patch copies of `modEngine` and the PJ sheet patch. The current `modEngine` is a superset; the older copies are not the running code.
- `_inspect/all_vba.txt` is an older one-file export of the workbook and is not the v8.5 project.
- Empty worksheet modules contain no code; they are listed in Section 3 but not expanded.

## 2. Reading guide

1. Start with Section 3 (inventory) and Section 4 (static-analysis observations).
2. Use Section 5 (procedure index) to locate a routine, then read the full module in Section 6.
3. Read `VBA_CODE_FLOW.md` for the Mermaid flow diagrams and the procedure-level call graph.
4. Line numbers in tables are 1-based within each exported component file.

## 3. Component inventory

### 3.1 Non-empty components

| Component | Type | Host sheet | Lines | Nonblank | Procedures | Primary role |
|---|---|---|---:|---:|---:|---|
| `ThisWorkbook` | document | ThisWorkbook | 10 | 8 | 1 | Workbook startup hook. |
| `Sheet12` | document | PJ | 290 | 255 | 9 | VAT purchase journal event handler. |
| `Sheet11` | document | PJ Non-Vat | 294 | 260 | 9 | Non-VAT purchase journal event handler. |
| `Sheet9` | document | PJ Non-Vat2 | 295 | 204 | 1 | Legacy non-VAT purchase journal event handler. |
| `Sheet3` | document | SJ | 175 | 147 | 6 | Sales journal event handler. |
| `Sheet8` | document | SJ2 | 88 | 74 | 2 | Alternate/legacy sales journal event handler. |
| `Sheet10` | document | SUPPLIERS DATA | 34 | 31 | 1 | Supplier duplicate checker. |
| `modEngine` | standard | (none) | 1781 | 1668 | 76 | Core sync/reference/supplier engine. |
| `modPJSync` | standard | (none) | 532 | 257 | 8 | Legacy PJ Non-Vat2 helper module. |
| `modGLGate` | standard | (none) | 51 | 43 | 3 | Pre-scan workbook structure gate. |
| `modGLRefresh` | standard | (none) | 28 | 22 | 4 | Macro-dialog wrappers for GL refresh. |
| `modGLScanEngine` | standard | (none) | 335 | 292 | 9 | Active CDJ/CRJ GL scan engine. |
| `modGLWorkbookMap` | standard | (none) | 211 | 189 | 12 | Account/column mapping and constants. |
| `modGraphify` | standard | (none) | 229 | 201 | 4 | Standalone charting utility. |
| `ImportSuppliersFromJSON` | standard | (none) | 856 | 741 | 1 | Destructive hard-coded supplier seeder. |

### 3.2 Empty worksheet modules (no code)

| CodeName | Host sheet |
|---|---|
| `Sheet1` | PJ2 |
| `Sheet2` | CDJ old |
| `Sheet4` | CRJ |
| `Sheet5` | GJ |
| `Sheet6` | GL |
| `Sheet7` | CDJ |
| `Sheet13` | CRJ2 |
| `Sheet14` | COA_MAP |
| `Sheet15` | GL_AUDIT |
| `Sheet17` | V8.3_FINAL_REVIEW |

The exported component set confirms that the active code is concentrated in `modEngine`, `modGLScanEngine`, `modGLWorkbookMap` and the PJ/SJ event modules. `GL`, `CRJ`, `COA_MAP`, `GL_AUDIT`, `PJ2`, `CRJ2`, `CDJ old` and `V8.3_FINAL_REVIEW` have empty code modules.

## 4. Static-analysis observations

These are source-level observations from the snapshot, not a substitute for a compile/run verification. Each item is written to be actionable for a later fix; no code was changed.

### 4.1 High-impact observations

1. **`Sheet9` references undeclared column constants.** `Sheet9.txt:78`, `:135`, `:187`, `:205`, `:212`, `:252` use `COL_TIN`, `COL_NAME`, `COL_AMOUNT`, `COL_DEBIT`, `COL_REF`, `COL_COA`, `COL_FULL_ID` without declaring them. The module has no `Option Explicit`, so they become empty Variants; column comparisons never match, and `Cells(row, COL_REF)` uses column 0. The legacy PJ Non-Vat2 path is therefore not reliable. The active Sheet12/Sheet11 implementations do declare their constants.
2. **COA status matching accepts `UNMAPPED`.** `modGLWorkbookMap.bas:75` tests `InStr(st, "MAPPED") > 0`. That substring is also present in `UNMAPPED`. If operators use `UNMAPPED` as a status, the alias is accepted even though the gate intends to hold unapproved labels. Prefer exact token/equality matching (`st = "MAPPED"` or an explicit allow-list).
3. **Hard-coded HOLD aliases are executable.** `modGLWorkbookMap.bas:139-143` maps `MEDICAL EQUIPMENT`, `COST OF REVENUE`, `SUPPLIES`, `BANK CHARGE` and `CHARGES` (plus the earlier five aliases) directly to GL accounts. The README says the supplied sundry labels are not assumed approved; the code nevertheless resolves several of them. This is a policy/implementation mismatch that should be resolved explicitly.
4. **The scan year is not enforced.** `modGLScanEngine.bas:82` and `:169` take `yr` but never use it. `ScanGL 2026` therefore aggregates all matching rows by month; it does not filter by year. CRJ additionally relies on partition order and the first boundary being January. This is safe only while the workbook contains one reporting year and correct boundaries.
5. **GL ending balance restarts at zero for every block.** `modGLScanEngine.bas:293` sets `bal = 0#` inside the 46-block loop. The I column is therefore a within-year cumulative balance (Jan..Dec), not a roll-forward from a prior-year closing balance. This differs from the README's stated roll-forward formula unless the opening balances are embedded in the source postings.

### 4.2 Fail-closed and data-integrity notes

6. **The gate is correctly positioned before GL write-back.** `ScanGL` writes only `GL_AUDIT` when `inv > 0`, unmapped amount exceeds tolerance, or debit/credit difference exceeds tolerance. `WriteScannedGL` is reached only after the gate passes.
7. **Duplicate/missing GL titles are validated twice.** `ValidateV8Structure` checks 46 nonblank titles; `ValidateGLAccountTitles` additionally detects duplicates (normalized, case-insensitive) and increments the invalid count. The scan aborts before writing if the count is nonzero.
8. **Sundry T/U disagreements are audit-only.** `ScanCDJ` derives the Sundry amount from the balanced row, records a correction note when the explicit T/U values differ, but does not increment `inv`. The gate can therefore PASS with a logged correction; this is intentional in the comments but should be understood by reviewers.
9. **Sync errors are swallowed.** The three public sync routines in `modEngine` (`SyncPJToCDJ`, `SyncSJToCRJ`, `SyncPJNonVatToCDJ`) end with empty `ErrorHandler:` labels. A failed write can leave the source journal and CDJ/CRJ out of sync without surfacing an error to the operator.
10. **Destructive supplier import.** `ImportSuppliersFromJSON.bas:13` clears `SUPPLIERS DATA` rows 2..end before inserting hard-coded records. It is not part of the normal event flow, but it is present in the same VBA project and can be run from the macro dialog.

### 4.3 Legacy and maintenance notes

11. **Legacy GJ/calc code remains in `modEngine`.** `GLScanGJ` (`:1494`), `Legacy_RefreshGL` (`:1537`), `Legacy_CalculateMonthlyNet` (`:1679`), `Legacy_ValidateGLConsistency` (`:1710`), `Legacy_RefreshAllGL` (`:1721`) and `BuildV8CalcSheet` (`:1771`) are still compiled. They are not on the active `ScanGL` path, but they keep GJ and `GL_V8_CALC` concepts alive in the source and may be reachable from other macros.
12. **Duplicate implementations.** Sheet11 and Sheet12 duplicate supplier lookup, totals and event orchestration; Sheet3 and Sheet8 duplicate SJ logic with different invoice-number formats. Sheet9 and modPJSync duplicate an older non-VAT path. Fixes must be applied in multiple places.
13. **Workbook-level verification is still required.** The v8.5 build has a recorded compile PASS and a v8.4.1 live refresh HOLD, but this snapshot alone cannot prove the embedded VBA matches these files or that a live scan passes. Use Excel Desktop/COM for that boundary.

## 5. Procedure index

The tables below list every declared `Sub`/`Function` in declaration order. `Implicit` means no `Public`/`Private` keyword was present in the source.

### ThisWorkbook (ThisWorkbook)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 3 | Sub | Private | `Workbook_Open` |

### Sheet12 (PJ)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 19 | Function | Private | `GetCOAList` |
| 34 | Function | Private | `SelectCOA` |
| 55 | Sub | Private | `LookupByTIN` |
| 101 | Sub | Private | `LookupByName` |
| 144 | Sub | Private | `ClearSupplierRow` |
| 150 | Sub | Private | `CalculateTotals` |
| 162 | Function | Private | `FindBlockStart` |
| 174 | Sub | Implicit | `ManualSyncToCDJ` |
| 192 | Sub | Private | `Worksheet_Change` |

### Sheet11 (PJ Non-Vat)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 19 | Function | Private | `GetCOAList` |
| 34 | Function | Private | `SelectCOA` |
| 55 | Sub | Private | `LookupByTIN` |
| 101 | Sub | Private | `LookupByName` |
| 144 | Sub | Private | `ClearSupplierRow` |
| 153 | Sub | Private | `CalculateTotals` |
| 165 | Function | Private | `FindBlockStart` |
| 177 | Sub | Implicit | `ManualSyncToCDJ` |
| 195 | Sub | Private | `Worksheet_Change` |

### Sheet9 (PJ Non-Vat2)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 2 | Sub | Private | `Worksheet_Change` |

### Sheet3 (SJ)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 21 | Sub | Private | `Worksheet_Activate` |
| 27 | Sub | Private | `LookupCustomer` |
| 34 | Sub | Private | `ClearCustomerRow` |
| 39 | Sub | Private | `Worksheet_Change` |
| 137 | Sub | Implicit | `ManualSyncToCRJ` |
| 171 | Function | Private | `SheetExists` |

### Sheet8 (SJ2)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 23 | Sub | Private | `SyncToCRJ` |
| 30 | Sub | Private | `Worksheet_Change` |

### Sheet10 (SUPPLIERS DATA)

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 1 | Sub | Implicit | `CheckDuplicates` |

### modEngine ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 129 | Sub | Public | `InitializeEngine` |
| 145 | Sub | Public | `InitializeCRJ` |
| 158 | Function | Private | `CellText` |
| 168 | Function | Private | `NormName` |
| 177 | Sub | Private | `NewDupDict` |
| 182 | Function | Private | `IsDupKey` |
| 191 | Sub | Private | `RegisterSupplier` |
| 213 | Sub | Private | `LoadSuppliers` |
| 236 | Sub | Public | `ReloadSuppliers` |
| 250 | Function | Public | `NormTIN` |
| 261 | Function | Public | `AmbiguousTINCount` |
| 268 | Function | Public | `AmbiguousTINList` |
| 278 | Function | Public | `IsAmbiguousTIN` |
| 284 | Function | Public | `AmbiguousNamesFor` |
| 291 | Sub | Private | `LoadCDJHeaders` |
| 330 | Sub | Private | `LoadCDJReferences` |
| 362 | Function | Private | `CDJRowIsFree` |
| 372 | Function | Public | `CDJMarkerMonth` |
| 380 | Function | Public | `CDJMonthAbbr` |
| 398 | Function | Public | `CDJTotalRow` |
| 410 | Function | Public | `FindCDJFreeRow` |
| 436 | Function | Public | `CDJReserveRow` |
| 466 | Sub | Private | `ShiftCDJRefCache` |
| 475 | Function | Public | `PJMonthOf` |
| 484 | Sub | Private | `LoadSequence` |
| 494 | Sub | Private | `LoadNextRow` |
| 502 | Sub | Private | `LoadCRJReferences` |
| 516 | Sub | Private | `LoadCRJNextRow` |
| 533 | Sub | Private | `LoadCRJSequence` |
| 546 | Function | Public | `GetNextPJNumber` |
| 557 | Function | Public | `GetNextInvoiceNumber` |
| 568 | Function | Public | `GetNextSeq` |
| 582 | Function | Public | `GetSupplierByTIN` |
| 598 | Function | Public | `GetSupplierByName` |
| 615 | Sub | Public | `AddSupplierToCache` |
| 631 | Function | Public | `AddSupplierRecord` |
| 714 | Function | Public | `FindExpenseColumn` |
| 766 | Function | Public | `CalculateWithholdingTax` |
| 784 | Function | Public | `FindCDJRow` |
| 818 | Sub | Public | `UpdateCDJReference` |
| 824 | Function | Public | `FindCRJRow` |
| 831 | Sub | Public | `UpdateCRJReference` |
| 840 | Sub | Public | `SyncPJToCDJ` |
| 966 | Sub | Public | `SyncSJToCRJ` |
| 1041 | Function | Public | `EngineStatus` |
| 1060 | Sub | Public | `ShowCDJStatus` |
| 1087 | Sub | Public | `ShowCRJStatus` |
| 1110 | Sub | Public | `ShowAmbiguousTINs` |
| 1126 | Sub | Public | `RepairEngine` |
| 1163 | Sub | Public | `ResetCRJ` |
| 1189 | Sub | Public | `ResetCDJ` |
| 1218 | Function | Public | `CalculateEWT` |
| 1238 | Sub | Public | `SyncPJNonVatToCDJ` |
| 1334 | Function | Private | `GLN` |
| 1343 | Function | Private | `GLAcct` |
| 1396 | Function | Private | `GLMonth` |
| 1416 | Sub | Private | `GLAdd` |
| 1426 | Sub | Private | `GLAudit` |
| 1433 | Sub | Private | `GLScanCDJ` |
| 1472 | Sub | Private | `GLScanCRJ` |
| 1494 | Sub | Private | `GLScanGJ` |
| 1514 | Function | Private | `GLAccounts` |
| 1537 | Sub | Private | `Legacy_RefreshGL` |
| 1618 | Function | Private | `SheetNamesList` |
| 1632 | Function | Private | `DetectGLYear` |
| 1679 | Function | Private | `Legacy_CalculateMonthlyNet` |
| 1700 | Function | Private | `Legacy_CalculateEndingBalanceFull` |
| 1710 | Function | Private | `Legacy_ValidateGLConsistency` |
| 1721 | Sub | Private | `Legacy_RefreshAllGL` |
| 1725 | Sub | Public | `InitializeCOALookup` |
| 1728 | Function | Public | `GetAccountFromCOA` |
| 1735 | Function | Public | `GetAccountNameFromTIN` |
| 1751 | Function | Public | `GetFinancialMonth` |
| 1755 | Function | Public | `GetFinancialYear` |
| 1767 | Function | Public | `ValidateGLConsistency` |
| 1771 | Function | Public | `BuildV8CalcSheet` |

### modPJSync ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 86 | Function | Private | `PJSheet` |
| 104 | Function | Public | `GetNextSequence` |
| 142 | Sub | Public | `LookupByTIN_NonVat` |
| 187 | Sub | Public | `LookupByName_NonVat` |
| 221 | Sub | Public | `ClearSupplierRow` |
| 241 | Sub | Public | `CalculateTotals` |
| 291 | Sub | Public | `SyncToCDJ_NonVat` |
| 475 | Function | Private | `FindExpenseCol` |

### modGLGate ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 6 | Function | Private | `SheetExists` |
| 12 | Function | Public | `ValidateV8Structure` |
| 39 | Sub | Public | `RunV8Gate` |

### modGLRefresh ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 13 | Sub | Public | `RefreshGL` |
| 17 | Sub | Public | `RefreshAllGL` |
| 21 | Sub | Public | `RefreshGLIntoSheet` |
| 25 | Sub | Public | `RefreshGLForYear` |

### modGLScanEngine ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 20 | Sub | Public | `ScanGL` |
| 71 | Sub | Private | `AddPost` |
| 82 | Sub | Private | `ScanCDJ` |
| 169 | Sub | Private | `ScanCRJ` |
| 254 | Function | Private | `ValidateGLAccountTitles` |
| 284 | Sub | Private | `WriteScannedGL` |
| 311 | Sub | Private | `WriteScanAudit` |
| 324 | Function | Private | `SafeText` |
| 329 | Function | Private | `MaxD` |

### modGLWorkbookMap ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 22 | Sub | Public | `V8_MapReset` |
| 32 | Function | Public | `V8_Key` |
| 44 | Sub | Private | `EnsureMaps` |
| 89 | Function | Public | `V8_Norm` |
| 94 | Function | Public | `V8_Month` |
| 111 | Function | Public | `V8_GLAccounts` |
| 122 | Function | Public | `V8_ResolveAccount` |
| 155 | Function | Public | `V8_CDJMap` |
| 174 | Function | Public | `GLTitleExists` |
| 182 | Function | Public | `GLTitleOf` |
| 193 | Function | Public | `V8_CRJMap` |
| 203 | Function | Public | `V8_Number` |

### modGraphify ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 37 | Sub | Implicit | `Graphify` |
| 173 | Function | Private | `PickNumericColumn` |
| 205 | Function | Private | `FirstNumericRow` |
| 225 | Function | Private | `ColumnLetter` |

### ImportSuppliersFromJSON ((none))

| Line | Kind | Scope | Procedure |
|---:|---|---|---|
| 4 | Sub | Implicit | `ImportSuppliersFromJSON` |

## 6. Full source code

Each section reproduces the component exactly as exported (line endings normalized to LF). The `Attribute VB_Name` line is retained because it is part of the export. Encoding artifacts in comments (for example `�` sequences) are preserved from the source files.

### 6.1 `ThisWorkbook` - Workbook startup hook.

- Source: `build/v8_5_vba/ThisWorkbook.txt`
- Type: document | Host: ThisWorkbook | CodeName: `ThisWorkbook`
- Lines: 10 | Nonblank: 8 | Procedures: 1
- SHA-256: `ced5d5eb36c8d69d7a2d1ace41db3a2c62cf277f2ba2cb2a02cb11ac19d642b4`

**Analysis notes**

- Runs once on workbook open and initializes the CDJ-side engine with target sheet `CDJ`.
- No explicit error handling; `modEngine.InitializeEngine` internally uses `On Error Resume Next` and still sets `gInitialized = True`.
- CRJ-side caches are not initialized here; they are loaded lazily by `InitializeCRJ` when a sales-journal sheet activates or a CRJ helper is called.

```vba
Option Explicit

Private Sub Workbook_Open()
    ' Initialize the engine – this loads all suppliers, CDJ references, headers, and sequence
    ' into memory. The target sheet is "CDJ" (you can change to "TEST" for testing).
    Call modEngine.InitializeEngine("CDJ")
    
    ' Optional: show a message to confirm
    ' MsgBox "Engine initialized. Status: " & modEngine.EngineStatus, vbInformation
End Sub
```

### 6.2 `Sheet12` - VAT purchase journal event handler.

- Source: `build/v8_5_vba/Sheet12.txt`
- Type: document | Host: PJ | CodeName: `Sheet12`
- Lines: 286 | Nonblank: 255 | Procedures: 9
- SHA-256: `a7c19a46a99d480057bb931e2976f6c39fb8cc72b71274b7adbb831264c668ad`

**Analysis notes**

- Hosts the command layer for the VAT purchase journal (`PJ`). Watches columns B, D, E, F, G and N.
- On a TIN/name edit it performs supplier lookup, and on TOTAL rows it recalculates the block totals.
- When gross amount in N is present, it writes formulas J=N/1.12 and M=N-J, stamps a `PJ-#####` reference and `PJINV-...` full ID, then calls `modEngine.SyncPJToCDJ`.
- A second trigger on column H re-runs the CDJ sync for already-stamped rows.
- Supplier-add path calls `modEngine.AddSupplierRecord`; it does not loop retries.
- Nearly identical to Sheet11; the differences are VAT formulas and the `SyncPJToCDJ` vs `SyncPJNonVatToCDJ` call.

```vba
Option Explicit

Private Const START_ROW As Long = 8
Private Const COL_TIN As Integer = 4
Private Const COL_NAME As Integer = 5
Private Const COL_ADDRESS As Integer = 6
Private Const COL_COA As Integer = 7
Private Const COL_REF As Integer = 8
Private Const COL_DEBIT_MIRROR As Integer = 9
Private Const COL_DEBIT As Integer = 10
Private Const COL_CREDIT_MIRROR As Integer = 12
Private Const COL_CREDIT As Integer = 13
Private Const COL_GROSS As Integer = 14
Private Const COL_FULL_ID As Integer = 17

' ============================================================
' COA LIST & SELECT (used only when adding new supplier)
' ============================================================
Private Function GetCOAList() As Variant
    Static coaArray As Variant
    If IsEmpty(coaArray) Then
        coaArray = Array( _
            "Representation", "Clinicians Fee - Corporators", "Value Added Tax Payable", _
            "Pantry Supplies", "Taxes and Licences", "Miscellaneous", "Advances to Employee", _
            "Government ER Share Cont", "Delivery", "Repairs and Maintenance", _
            "Clinicians Fee - Visiting", "Association Dues", "Advertisement", _
            "Clinic Materials and Supplies", "Transportation and Travel", "Fuel and Oil", _
            "Supplies", "Rental", "Communication", "Professional Fees" _
        )
    End If
    GetCOAList = coaArray
End Function

Private Function SelectCOA() As String
    Dim coaList As Variant, i As Integer, inputStr As String, userInput As String, idx As Integer
    coaList = GetCOAList()
    inputStr = "Select Chart of Account:" & vbCrLf
    For i = LBound(coaList) To UBound(coaList)
        inputStr = inputStr & (i + 1) & ". " & coaList(i) & vbCrLf
    Next i
    inputStr = inputStr & "Enter number (1-" & UBound(coaList) + 1 & "):"
    userInput = InputBox(inputStr, "Chart of Account Selection")
    If userInput = "" Then SelectCOA = "": Exit Function
    If IsNumeric(userInput) Then
        idx = CInt(userInput) - 1
        If idx >= LBound(coaList) And idx <= UBound(coaList) Then SelectCOA = coaList(idx): Exit Function
    End If
    MsgBox "Invalid selection.", vbExclamation
    SelectCOA = SelectCOA()
End Function

' ============================================================
' LOOKUP (uses modEngine)
' ============================================================
Private Sub LookupByTIN(rowNum As Long, tin As String)
    Dim data As Variant
    Dim name As String, address As String, coa As String, msg As String

    tin = Trim(tin): If tin = "" Then Exit Sub
    data = modEngine.GetSupplierByTIN(tin)

    If Not IsArray(data) Then
        ' 1) TIN is shared by DIFFERENT suppliers: never guess, never offer "Add?"
        If modEngine.IsAmbiguousTIN(tin) Then
            MsgBox "TIN '" & tin & "' is assigned to more than one supplier in SUPPLIERS DATA:" & _
                   vbCrLf & vbCrLf & modEngine.AmbiguousNamesFor(tin) & vbCrLf & vbCrLf & _
                   "Correct SUPPLIERS DATA, then enter the TIN again.", vbExclamation, "Ambiguous TIN"
            Exit Sub
        End If

        ' 2) Genuinely missing: ask ONCE
        If MsgBox("TIN '" & tin & "' not found. Add?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

        name = Trim(Me.Cells(rowNum, COL_NAME).Value)
        If name = "" Then name = Trim(InputBox("Supplier Name:"))
        If name = "" Then Exit Sub
        address = Trim(Me.Cells(rowNum, COL_ADDRESS).Value)
        If address = "" Then address = Trim(InputBox("Address:"))
        coa = Trim(Me.Cells(rowNum, COL_COA).Value)
        If coa = "" Then coa = SelectCOA()
        If coa = "" Then Exit Sub

        ' 3) Add + verify.  NO recursive LookupByTIN call, so it cannot loop.
        If Not modEngine.AddSupplierRecord(tin, name, address, coa, msg) Then
            MsgBox msg, vbExclamation, "Supplier not added"
            Exit Sub
        End If
        data = modEngine.GetSupplierByTIN(tin)
        If Not IsArray(data) Then
            MsgBox "Supplier was saved but still cannot be resolved. Check SUPPLIERS DATA.", _
                   vbExclamation, "Supplier lookup"
            Exit Sub
        End If
    End If

    Me.Cells(rowNum, COL_NAME).Value = data(0)
    Me.Cells(rowNum, COL_ADDRESS).Value = data(1)
    Me.Cells(rowNum, COL_COA).Value = data(2)
End Sub

Private Sub LookupByName(rowNum As Long, name As String)
    Dim data As Variant
    Dim tin As String, address As String, coa As String, msg As String

    name = Trim(name): If name = "" Then Exit Sub
    data = modEngine.GetSupplierByName(name)

    If Not IsArray(data) Then
        If MsgBox("Supplier '" & name & "' not found. Add?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

        tin = Trim(Me.Cells(rowNum, COL_TIN).Value)
        If tin = "" Then tin = Trim(InputBox("TIN:"))
        If tin = "" Then Exit Sub
        address = Trim(Me.Cells(rowNum, COL_ADDRESS).Value)
        If address = "" Then address = Trim(InputBox("Address:"))
        coa = Trim(Me.Cells(rowNum, COL_COA).Value)
        If coa = "" Then coa = SelectCOA()
        If coa = "" Then Exit Sub

        If Not modEngine.AddSupplierRecord(tin, name, address, coa, msg) Then
            MsgBox msg, vbExclamation, "Supplier not added"
            Exit Sub
        End If
        data = modEngine.GetSupplierByName(name)
        If Not IsArray(data) Then
            MsgBox "Supplier was saved but still cannot be resolved. Check SUPPLIERS DATA.", _
                   vbExclamation, "Supplier lookup"
            Exit Sub
        End If
    End If

    Me.Cells(rowNum, COL_TIN).Value = data(0)
    Me.Cells(rowNum, COL_ADDRESS).Value = data(1)
    Me.Cells(rowNum, COL_COA).Value = data(2)
End Sub

' ============================================================
' ADD SUPPLIER TO MASTER + UPDATE CACHE
' ============================================================

' ============================================================
' HELPERS (Totals, Block Start)
' ============================================================
Private Sub ClearSupplierRow(rowNum As Long)
    Me.Cells(rowNum, COL_NAME).Value = ""
    Me.Cells(rowNum, COL_ADDRESS).Value = ""
    Me.Cells(rowNum, COL_COA).Value = ""
End Sub

Private Sub CalculateTotals(rowNum As Long)
    Dim startRow As Long, endRow As Long, debitSum As Double, creditSum As Double
    startRow = FindBlockStart(rowNum): endRow = rowNum - 1
    If endRow < startRow Then Exit Sub
    debitSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_DEBIT), Me.Cells(endRow, COL_DEBIT)))
    creditSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_CREDIT), Me.Cells(endRow, COL_CREDIT)))
    Me.Cells(rowNum, COL_DEBIT).Value = debitSum
    Me.Cells(rowNum, COL_CREDIT).Value = creditSum
    Me.Cells(rowNum, COL_DEBIT_MIRROR).Value = debitSum
    Me.Cells(rowNum, COL_CREDIT_MIRROR).Value = creditSum
End Sub

Private Function FindBlockStart(rowNum As Long) As Long
    Dim i As Long, val As String
    For i = rowNum - 1 To START_ROW Step -1
        val = Trim(Me.Cells(i, COL_TIN).Value)
        If LCase(val) = "total" Then FindBlockStart = i + 1: Exit Function
    Next i
    FindBlockStart = START_ROW
End Function

' ============================================================
' MANUAL SYNC MACRO (optional)
' ============================================================
Sub ManualSyncToCDJ()
    Dim rowNum As Long
    rowNum = ActiveCell.Row
    If rowNum < START_ROW Then
        MsgBox "Select a row below row 8.", vbExclamation
        Exit Sub
    End If
    If Me.Cells(rowNum, COL_REF).Value = "" Then
        MsgBox "No reference ID (PJ-XXXXX) in column H. Enter Gross Amount first.", vbExclamation
        Exit Sub
    End If
    Call modEngine.SyncPJToCDJ(Me, rowNum)
    MsgBox "Sync attempted for row " & rowNum & ". Check target sheet.", vbInformation
End Sub

' ============================================================
' MAIN TRIGGER
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngLookup As Range, rngSync As Range, cell As Range
    Dim currentTimestamp As String, newRef As String

    ' ---- PART 1: LOOKUP, VAT, ID STAMPING AND SYNC ----
    Set rngLookup = Intersect(Target, Union(Me.Range("B:B"), Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then
                    Me.Cells(cell.Row, COL_DEBIT).Formula = "=N" & cell.Row & "/1.12"
                    Me.Cells(cell.Row, COL_CREDIT).Formula = "=N" & cell.Row & "-J" & cell.Row
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Formula = "=J" & cell.Row
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Formula = "=M" & cell.Row

                    ' ---- ID STAMPING AND IMMEDIATE SYNC ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" Then
                        If Me.Cells(cell.Row, COL_REF).Value = "" Then
                            newRef = modEngine.GetNextPJNumber()
                            Me.Cells(cell.Row, COL_REF).Value = newRef
                            Me.Cells(cell.Row, COL_FULL_ID).Value = "PJINV-" & currentTimestamp & "-" & Right(newRef, 5)
                        End If

                        ' ====== SYNC TO CDJ (New or Updated) ======
                        Call modEngine.SyncPJToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

    ' ---- PART 2: SYNC (Triggered by Column H) ----
    ' Kept as a safety fallback (will do nothing if already synced)
    Set rngSync = Intersect(Target, Me.Range("H:H"))
    If Not rngSync Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupSync

        For Each cell In rngSync
            If cell.Row >= START_ROW Then
                newRef = Trim(Me.Cells(cell.Row, COL_REF).Value)
                If newRef <> "" And newRef <> "TOTAL" Then
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" And _
                       Me.Cells(cell.Row, COL_GROSS).Value <> "" Then
                        Call modEngine.SyncPJToCDJ(Me, cell.Row)
                    End If
                End If
            End If
        Next cell

CleanupSync:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub
```

### 6.3 `Sheet11` - Non-VAT purchase journal event handler.

- Source: `build/v8_5_vba/Sheet11.txt`
- Type: document | Host: PJ Non-Vat | CodeName: `Sheet11`
- Lines: 292 | Nonblank: 260 | Procedures: 9
- SHA-256: `4c31c716a3211fadbaba3c5df52493988abf89c1e5c08989e1b0b90281969e1b`

**Analysis notes**

- Non-VAT twin of Sheet12. Watches the same input columns.
- Forces J=N, M=0 and mirrors the amount, stamps a `PNV-#####` reference and `PNVINV-...` full ID, then calls `modEngine.SyncPJNonVatToCDJ`.
- Uses the same supplier lookup and `CalculateTotals`/`FindBlockStart` helpers as Sheet12, duplicated in this module.
- No VAT or EWT formula is calculated on the sheet; EWT is derived in `modEngine.CalculateEWT` during sync.

```vba
Option Explicit

Private Const START_ROW As Long = 8
Private Const COL_TIN As Integer = 4
Private Const COL_NAME As Integer = 5
Private Const COL_ADDRESS As Integer = 6
Private Const COL_COA As Integer = 7
Private Const COL_REF As Integer = 8
Private Const COL_DEBIT_MIRROR As Integer = 9
Private Const COL_DEBIT As Integer = 10
Private Const COL_CREDIT_MIRROR As Integer = 12
Private Const COL_CREDIT As Integer = 13
Private Const COL_GROSS As Integer = 14          ' N = Amount (Net, no VAT)
Private Const COL_FULL_ID As Integer = 17

' ============================================================
' COA LIST & SELECT
' ============================================================
Private Function GetCOAList() As Variant
    Static coaArray As Variant
    If IsEmpty(coaArray) Then
        coaArray = Array( _
            "Representation", "Clinicians Fee - Corporators", "Value Added Tax Payable", _
            "Pantry Supplies", "Taxes and Licences", "Miscellaneous", "Advances to Employee", _
            "Government ER Share Cont", "Delivery", "Repairs and Maintenance", _
            "Clinicians Fee - Visiting", "Association Dues", "Advertisement", _
            "Clinic Materials and Supplies", "Transportation and Travel", "Fuel and Oil", _
            "Supplies", "Rental", "Communication", "Professional Fees" _
        )
    End If
    GetCOAList = coaArray
End Function

Private Function SelectCOA() As String
    Dim coaList As Variant, i As Integer, inputStr As String, userInput As String, idx As Integer
    coaList = GetCOAList()
    inputStr = "Select Chart of Account:" & vbCrLf
    For i = LBound(coaList) To UBound(coaList)
        inputStr = inputStr & (i + 1) & ". " & coaList(i) & vbCrLf
    Next i
    inputStr = inputStr & "Enter number (1-" & UBound(coaList) + 1 & "):"
    userInput = InputBox(inputStr, "Chart of Account Selection")
    If userInput = "" Then SelectCOA = "": Exit Function
    If IsNumeric(userInput) Then
        idx = CInt(userInput) - 1
        If idx >= LBound(coaList) And idx <= UBound(coaList) Then SelectCOA = coaList(idx): Exit Function
    End If
    MsgBox "Invalid selection.", vbExclamation
    SelectCOA = SelectCOA()
End Function

' ============================================================
' LOOKUP (uses modEngine)
' ============================================================
Private Sub LookupByTIN(rowNum As Long, tin As String)
    Dim data As Variant
    Dim name As String, address As String, coa As String, msg As String

    tin = Trim(tin): If tin = "" Then Exit Sub
    data = modEngine.GetSupplierByTIN(tin)

    If Not IsArray(data) Then
        ' 1) TIN is shared by DIFFERENT suppliers: never guess, never offer "Add?"
        If modEngine.IsAmbiguousTIN(tin) Then
            MsgBox "TIN '" & tin & "' is assigned to more than one supplier in SUPPLIERS DATA:" & _
                   vbCrLf & vbCrLf & modEngine.AmbiguousNamesFor(tin) & vbCrLf & vbCrLf & _
                   "Correct SUPPLIERS DATA, then enter the TIN again.", vbExclamation, "Ambiguous TIN"
            Exit Sub
        End If

        ' 2) Genuinely missing: ask ONCE
        If MsgBox("TIN '" & tin & "' not found. Add?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

        name = Trim(Me.Cells(rowNum, COL_NAME).Value)
        If name = "" Then name = Trim(InputBox("Supplier Name:"))
        If name = "" Then Exit Sub
        address = Trim(Me.Cells(rowNum, COL_ADDRESS).Value)
        If address = "" Then address = Trim(InputBox("Address:"))
        coa = Trim(Me.Cells(rowNum, COL_COA).Value)
        If coa = "" Then coa = SelectCOA()
        If coa = "" Then Exit Sub

        ' 3) Add + verify.  NO recursive LookupByTIN call, so it cannot loop.
        If Not modEngine.AddSupplierRecord(tin, name, address, coa, msg) Then
            MsgBox msg, vbExclamation, "Supplier not added"
            Exit Sub
        End If
        data = modEngine.GetSupplierByTIN(tin)
        If Not IsArray(data) Then
            MsgBox "Supplier was saved but still cannot be resolved. Check SUPPLIERS DATA.", _
                   vbExclamation, "Supplier lookup"
            Exit Sub
        End If
    End If

    Me.Cells(rowNum, COL_NAME).Value = data(0)
    Me.Cells(rowNum, COL_ADDRESS).Value = data(1)
    Me.Cells(rowNum, COL_COA).Value = data(2)
End Sub

Private Sub LookupByName(rowNum As Long, name As String)
    Dim data As Variant
    Dim tin As String, address As String, coa As String, msg As String

    name = Trim(name): If name = "" Then Exit Sub
    data = modEngine.GetSupplierByName(name)

    If Not IsArray(data) Then
        If MsgBox("Supplier '" & name & "' not found. Add?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

        tin = Trim(Me.Cells(rowNum, COL_TIN).Value)
        If tin = "" Then tin = Trim(InputBox("TIN:"))
        If tin = "" Then Exit Sub
        address = Trim(Me.Cells(rowNum, COL_ADDRESS).Value)
        If address = "" Then address = Trim(InputBox("Address:"))
        coa = Trim(Me.Cells(rowNum, COL_COA).Value)
        If coa = "" Then coa = SelectCOA()
        If coa = "" Then Exit Sub

        If Not modEngine.AddSupplierRecord(tin, name, address, coa, msg) Then
            MsgBox msg, vbExclamation, "Supplier not added"
            Exit Sub
        End If
        data = modEngine.GetSupplierByName(name)
        If Not IsArray(data) Then
            MsgBox "Supplier was saved but still cannot be resolved. Check SUPPLIERS DATA.", _
                   vbExclamation, "Supplier lookup"
            Exit Sub
        End If
    End If

    Me.Cells(rowNum, COL_TIN).Value = data(0)
    Me.Cells(rowNum, COL_ADDRESS).Value = data(1)
    Me.Cells(rowNum, COL_COA).Value = data(2)
End Sub

' ============================================================
' ADD SUPPLIER TO MASTER + UPDATE CACHE
' ============================================================

' ============================================================
' CLEAR SUPPLIER ROW
' ============================================================
Private Sub ClearSupplierRow(rowNum As Long)
    Me.Cells(rowNum, COL_NAME).Value = ""
    Me.Cells(rowNum, COL_ADDRESS).Value = ""
    Me.Cells(rowNum, COL_COA).Value = ""
End Sub

' ============================================================
' CALCULATE TOTALS
' ============================================================
Private Sub CalculateTotals(rowNum As Long)
    Dim startRow As Long, endRow As Long, debitSum As Double, creditSum As Double
    startRow = FindBlockStart(rowNum): endRow = rowNum - 1
    If endRow < startRow Then Exit Sub
    debitSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_DEBIT), Me.Cells(endRow, COL_DEBIT)))
    creditSum = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(startRow, COL_CREDIT), Me.Cells(endRow, COL_CREDIT)))
    Me.Cells(rowNum, COL_DEBIT).Value = debitSum
    Me.Cells(rowNum, COL_CREDIT).Value = creditSum
    Me.Cells(rowNum, COL_DEBIT_MIRROR).Value = debitSum
    Me.Cells(rowNum, COL_CREDIT_MIRROR).Value = creditSum
End Sub

Private Function FindBlockStart(rowNum As Long) As Long
    Dim i As Long, val As String
    For i = rowNum - 1 To START_ROW Step -1
        val = Trim(Me.Cells(i, COL_TIN).Value)
        If LCase(val) = "total" Then FindBlockStart = i + 1: Exit Function
    Next i
    FindBlockStart = START_ROW
End Function

' ============================================================
' MANUAL SYNC MACRO
' ============================================================
Sub ManualSyncToCDJ()
    Dim rowNum As Long
    rowNum = ActiveCell.Row
    If rowNum < START_ROW Then
        MsgBox "Select a row below row 8.", vbExclamation
        Exit Sub
    End If
    If Me.Cells(rowNum, COL_REF).Value = "" Then
        MsgBox "No reference ID (PNV-XXXXX) in column H. Enter Amount first.", vbExclamation
        Exit Sub
    End If
    Call modEngine.SyncPJNonVatToCDJ(Me, rowNum)
    MsgBox "Sync attempted for row " & rowNum & ". Check target sheet.", vbInformation
End Sub

' ============================================================
' MAIN TRIGGER
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngLookup As Range, rngSync As Range, cell As Range
    Dim currentTimestamp As String, newRef As String

    ' ---- PART 1: LOOKUP, AMOUNT (NO VAT), ID STAMPING AND SYNC ----
    Set rngLookup = Intersect(Target, Union(Me.Range("B:B"), Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))    ' Column N = Net Amount
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                ' ---- NON-VAT AMOUNT (NO VAT) ----
                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then

                    ' Set J = N, M = 0
                    Me.Cells(cell.Row, COL_DEBIT).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT).Value = 0
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Value = 0

                    ' ---- ID STAMPING (PNV-XXXXX) AND SYNC ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" Then
                        If Me.Cells(cell.Row, COL_REF).Value = "" Then
                            newRef = modEngine.GetNextPJNumber()       ' Uses sequence from AA13
                            newRef = "PNV-" & Right(newRef, 5)         ' Change prefix to PNV
                            Me.Cells(cell.Row, COL_REF).Value = newRef
                            Me.Cells(cell.Row, COL_FULL_ID).Value = "PNVINV-" & currentTimestamp & "-" & Right(newRef, 5)
                        End If

                        ' ---- SYNC TO CDJ (New or Updated) ----
                        Call modEngine.SyncPJNonVatToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

    ' ---- PART 2: SYNC (Triggered by Column H) ----
    Set rngSync = Intersect(Target, Me.Range("H:H"))
    If Not rngSync Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupSync

        For Each cell In rngSync
            If cell.Row >= START_ROW Then
                newRef = Trim(Me.Cells(cell.Row, COL_REF).Value)
                If newRef <> "" And newRef <> "TOTAL" Then
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" And _
                       Me.Cells(cell.Row, COL_GROSS).Value <> "" Then
                        Call modEngine.SyncPJNonVatToCDJ(Me, cell.Row)
                    End If
                End If
            End If
        Next cell

CleanupSync:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub
```

### 6.4 `Sheet9` - Legacy non-VAT purchase journal event handler.

- Source: `build/v8_5_vba/Sheet9.txt`
- Type: document | Host: PJ Non-Vat2 | CodeName: `Sheet9`
- Lines: 292 | Nonblank: 204 | Procedures: 1
- SHA-256: `bc91a4742597954803c540345add3ca2df957845f7ad8e96802fea19a9a1491c`

**Analysis notes**

- Legacy handler for `PJ Non-Vat2`, delegating to `modPJSync` helpers and `modEngine.SyncPJNonVatToCDJ`-style direct CDJ writes.
- Audit finding: the module has no `Option Explicit` and references `COL_TIN`, `COL_NAME`, `COL_AMOUNT`, `COL_DEBIT`, `COL_DEBIT_MIRROR`, `COL_REF`, `COL_COA` and `COL_FULL_ID` without declaring them or importing them from another module (see `Sheet9.txt:78`, `:135`, `:187`, `:205`, `:212`, `:252`).
- In VBA those identifiers become implicit empty Variants. Column tests such as `c.Column = COL_TIN` are therefore never true, and `Cells(row, COL_REF)` uses column 0, which raises a run-time error. The pass-2 sync path is effectively nonfunctional unless the module is edited.
- This is a legacy path; current PJ/PJ Non-Vat workbooks use Sheet12/Sheet11.

```vba

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range
    Dim c As Range

    Dim currentTimestamp As String
    Dim tinVal As String
    Dim nameVal As String
    Dim refVal As String
    Dim coaVal As String
    Dim amountVal As Variant
    Dim seqNum As Long

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo ErrorHandler

    ' ============================================================
    ' SAVE EXCEL APPLICATION STATE
    ' ============================================================

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation

    ' ============================================================
    ' LIMIT EVENT TO USER-EDITABLE PJ NON-VAT INPUT COLUMNS
    '
    ' D = TIN
    ' E = Supplier / Vendor Name
    ' F = Address
    ' G = Description
    ' I = Non-VAT Amount
    ' ============================================================

    Set rng = Intersect( _
        Target, _
        Union( _
            Me.Columns("D"), _
            Me.Columns("E"), _
            Me.Columns("F"), _
            Me.Columns("G"), _
            Me.Columns("I") _
        ) _
    )

    If rng Is Nothing Then Exit Sub

    ' ============================================================
    ' PROTECT AGAINST RE-ENTRY
    ' ============================================================

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    currentTimestamp = Format$(Now, "yyyymmdd-HHnnss")

    ' ============================================================
    ' PASS 1
    '
    ' SUPPLIER LOOKUP
    ' NON-VAT AMOUNT CALCULATION
    ' REFERENCE GENERATION
    ' ============================================================

    For Each c In rng.Cells

        If c.Row >= START_ROW Then

            ' ====================================================
            ' SUPPLIER LOOKUP
            ' ====================================================

            If c.Column = COL_TIN Or c.Column = COL_NAME Then

                tinVal = Trim$(CStr( _
                    Me.Cells(c.Row, COL_TIN).Value _
                ))

                nameVal = Trim$(CStr( _
                    Me.Cells(c.Row, COL_NAME).Value _
                ))

                ' ------------------------------------------------
                ' TOTAL ROW
                ' ------------------------------------------------

                If StrComp(tinVal, "TOTAL", vbTextCompare) = 0 _
                   Or StrComp(nameVal, "TOTAL", vbTextCompare) = 0 Then

                    ' IMPORTANT:
                    ' CalculateTotals receives a Long row number.
                    ' The helper declaration must therefore be:
                    '
                    ' Private Sub CalculateTotals(ByVal rowNum As Long)

                    CalculateTotals CLng(c.Row)

                ' ------------------------------------------------
                ' LOOKUP BY TIN
                ' ------------------------------------------------

                ElseIf Len(tinVal) > 0 Then

                    LookupByTIN_NonVat CLng(c.Row), tinVal

                ' ------------------------------------------------
                ' LOOKUP BY NAME
                ' ------------------------------------------------

                ElseIf Len(nameVal) > 0 Then

                    LookupByName_NonVat CLng(c.Row), nameVal

                ' ------------------------------------------------
                ' BOTH EMPTY
                ' ------------------------------------------------

                Else

                    ClearSupplierRow CLng(c.Row)

                End If

            End If

            ' ====================================================
            ' NON-VAT AMOUNT
            ' ====================================================

            If c.Column = COL_AMOUNT Then

                amountVal = Me.Cells(c.Row, COL_AMOUNT).Value

                tinVal = UCase$(Trim$(CStr( _
                    Me.Cells(c.Row, COL_TIN).Value _
                )))

                nameVal = UCase$(Trim$(CStr( _
                    Me.Cells(c.Row, COL_NAME).Value _
                )))

                ' ------------------------------------------------
                ' DO NOT PROCESS TOTAL ROWS
                ' ------------------------------------------------

                If tinVal <> "TOTAL" And nameVal <> "TOTAL" Then

                    ' ------------------------------------------------
                    ' VALID AMOUNT
                    ' ------------------------------------------------

                    If Len(Trim$(CStr(amountVal))) > 0 Then

                        ' --------------------------------------------
                        ' DEBIT AMOUNT
                        ' --------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT).Value = amountVal

                        ' --------------------------------------------
                        ' DEBIT MIRROR
                        ' --------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT_MIRROR).Value = _
                            amountVal

                        ' --------------------------------------------
                        ' NON-VAT
                        '
                        ' VAT = ZERO
                        ' EWT = ZERO
                        ' --------------------------------------------

                        Me.Cells(c.Row, 13).Value = 0
                        Me.Cells(c.Row, 12).Value = 0

                        ' --------------------------------------------
                        ' REFERENCE GENERATION
                        ' --------------------------------------------

                        refVal = Trim$(CStr( _
                            Me.Cells(c.Row, COL_REF).Value _
                        ))

                        coaVal = Trim$(CStr( _
                            Me.Cells(c.Row, COL_COA).Value _
                        ))

                        If Len(tinVal) > 0 _
                           And Len(nameVal) > 0 _
                           And Len(coaVal) > 0 _
                           And Len(refVal) = 0 Then

                            seqNum = CLng(GetNextSequence())

                            ' ----------------------------------------
                            ' PJ NON-VAT REFERENCE
                            ' ----------------------------------------

                            Me.Cells(c.Row, COL_REF).Value = _
                                "PNV-" & Format$(seqNum, "00000")

                            ' ----------------------------------------
                            ' FULL UNIQUE TRANSACTION ID
                            ' ----------------------------------------

                            Me.Cells(c.Row, COL_FULL_ID).Value = _
                                "PNVINV-" & currentTimestamp & "-" & _
                                Format$(seqNum, "00000")

                        End If

                    Else

                        ' ------------------------------------------------
                        ' AMOUNT CLEARED
                        ' ------------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT).ClearContents

                        Me.Cells(c.Row, COL_DEBIT_MIRROR).ClearContents

                        Me.Cells(c.Row, 12).Value = 0
                        Me.Cells(c.Row, 13).Value = 0

                    End If

                End If

            End If

        End If

    Next c

    ' ============================================================
    ' PASS 2
    '
    ' SYNCHRONIZE VALID PNV RECORDS TO CDJ
    ' ============================================================

    For Each c In rng.Cells

        If c.Row >= START_ROW Then

            refVal = Trim$(CStr( _
                Me.Cells(c.Row, COL_REF).Value _
            ))

            ' ----------------------------------------------------
            ' ONLY REAL PNV RECORDS
            ' ----------------------------------------------------

            If Len(refVal) > 0 _
               And StrComp(refVal, "TOTAL", vbTextCompare) <> 0 Then

                SyncToCDJ_NonVat CLng(c.Row)

            End If

        End If

    Next c

CleanExit:

    ' ============================================================
    ' RESTORE ORIGINAL EXCEL APPLICATION STATE
    ' ============================================================

    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    Exit Sub

ErrorHandler:

    MsgBox _
        "PJ Non-VAT Worksheet_Change error:" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.description, _
        vbExclamation, _
        "PNV Automation"

    Resume CleanExit

End Sub
```

### 6.5 `Sheet3` - Sales journal event handler.

- Source: `build/v8_5_vba/Sheet3.txt`
- Type: document | Host: SJ | CodeName: `Sheet3`
- Lines: 175 | Nonblank: 147 | Procedures: 6
- SHA-256: `8fd0dbd14dc3e9321881ca7feb0e1d6b1e782864d8649f6533688301fbebf0e8`

**Analysis notes**

- Event layer for the `SJ` sales journal. Watches F, G, J and K, then U for sync.
- Writes amount formats, output-VAT and net-sales formulas, and stamps `SJ-#####` plus `SJINV-...` identifiers.
- Delegates CRJ writes to `modEngine.SyncSJToCRJ`; provides `ManualSyncToCRJ` for the active row.
- The `SheetExists` helper is local to this module; errors in the event are handled with a cleanup label that restores Excel application state.

```vba
Option Explicit

Private Const START_ROW As Long = 10
Private Const CRJ_SHEET_NAME As String = "CRJ"

' SJ Columns (now final)
Private Const SJ_COL_DATE As Integer = 3
Private Const SJ_COL_CUSTOMER As Integer = 6
Private Const SJ_COL_CUSTOMER_ADDRESS As Integer = 7
Private Const SJ_COL_DESCRIPTION As Integer = 8
Private Const SJ_COL_AMOUNT As Integer = 10
Private Const SJ_COL_DISCOUNT As Integer = 11
Private Const SJ_COL_OUTPUTVAT_MIRROR As Integer = 12
Private Const SJ_COL_OUTPUTVAT As Integer = 13
Private Const SJ_COL_NETSALES_MIRROR As Integer = 14
Private Const SJ_COL_NETSALES As Integer = 15
Private Const SJ_COL_INVOICE As Integer = 21         ' U – auto short number (SJ-NNNNN)
Private Const SJ_COL_FULLID As Integer = 22          ' V – auto full ID (SJINV-…)
Private Const SJ_COL_PHYSICAL_INV As Integer = 9     ' I – manual physical invoice (NEVER overwritten)

Private Sub Worksheet_Activate()
    On Error Resume Next
    modEngine.InitializeCRJ
    On Error GoTo 0
End Sub

Private Sub LookupCustomer(rowNum As Long, name As String)
    name = Trim(name)
    If name = "" Then Exit Sub
    Me.Cells(rowNum, SJ_COL_CUSTOMER_ADDRESS).Value = "NCR Metro Manila"
    Me.Cells(rowNum, SJ_COL_DESCRIPTION).Value = "Full Payment for Service"
End Sub

Private Sub ClearCustomerRow(rowNum As Long)
    Me.Cells(rowNum, SJ_COL_CUSTOMER_ADDRESS).Value = ""
    Me.Cells(rowNum, SJ_COL_DESCRIPTION).Value = ""
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngLookup As Range, rngSync As Range, cell As Range
    Dim currentTimestamp As String, newInvoice As Variant

    ' ---- PART 1: LOOKUP, FORMULAS, INVOICE STAMPING ----
    Set rngLookup = Intersect(Target, Union(Me.Range("F:F"), Me.Range("G:G"), Me.Range("J:J"), Me.Range("K:K")))
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then

                ' ---- SKIP TOTAL ROW ----
                If UCase(Trim(Me.Cells(cell.Row, 5).Value)) = "TOTAL" Then GoTo NextLookup

                If cell.Column = SJ_COL_CUSTOMER Then
                    Call LookupCustomer(cell.Row, Trim(Me.Cells(cell.Row, SJ_COL_CUSTOMER).Value))
                End If

                If Me.Cells(cell.Row, SJ_COL_AMOUNT).Value <> "" And _
                   Me.Cells(cell.Row, SJ_COL_AMOUNT).Value <> "TOTAL" Then

                    ' Formats
                    Me.Cells(cell.Row, SJ_COL_AMOUNT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                    Me.Cells(cell.Row, SJ_COL_DISCOUNT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                    Me.Cells(cell.Row, SJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                    Me.Cells(cell.Row, SJ_COL_NETSALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

                    ' Formulas
                    Me.Cells(cell.Row, SJ_COL_OUTPUTVAT).Formula = "=IF(J" & cell.Row & "-K" & cell.Row & "=0, 0, (J" & cell.Row & "-K" & cell.Row & ")*0.107142857)"
                    Me.Cells(cell.Row, SJ_COL_NETSALES).Formula = "=J" & cell.Row & "-K" & cell.Row & "-M" & cell.Row
                    Me.Cells(cell.Row, SJ_COL_OUTPUTVAT_MIRROR).Formula = "=M" & cell.Row
                    Me.Cells(cell.Row, SJ_COL_NETSALES_MIRROR).Formula = "=O" & cell.Row

                    ' ---- STAMP AUTO INVOICE (U) & FULL ID (V) ----
                    If Me.Cells(cell.Row, SJ_COL_INVOICE).Value = "" Then
                        newInvoice = modEngine.GetNextInvoiceNumber()   ' e.g. 304
                        Me.Cells(cell.Row, SJ_COL_INVOICE).Value = "SJ-" & Format(newInvoice, "00000")

                        ' Read the manual physical invoice from column I
                        Dim physicalInv As String
                        If IsNumeric(Me.Cells(cell.Row, SJ_COL_PHYSICAL_INV).Value) Then
                            physicalInv = Format(CLng(Me.Cells(cell.Row, SJ_COL_PHYSICAL_INV).Value), "00000")
                        Else
                            physicalInv = "00000"   ' fallback if I is blank
                        End If

                        Me.Cells(cell.Row, SJ_COL_FULLID).Value = "SJINV-" & currentTimestamp & "-" & physicalInv & "-" & Format(newInvoice, "00000")

                        If SheetExists(CRJ_SHEET_NAME) Then
                            Call modEngine.SyncSJToCRJ(Me, cell.Row)
                        End If
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

    ' ---- PART 2: SYNC (Triggered by change in U – Invoice column) ----
    Set rngSync = Intersect(Target, Me.Range("U:U"))
    If Not rngSync Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupSync

        For Each cell In rngSync
            If cell.Row >= START_ROW Then
                ' ---- SKIP TOTAL ROW ----
                If UCase(Trim(Me.Cells(cell.Row, 5).Value)) = "TOTAL" Then GoTo NextSync

                If Me.Cells(cell.Row, SJ_COL_INVOICE).Value <> "" And _
                   Me.Cells(cell.Row, SJ_COL_INVOICE).Value <> "TOTAL" Then
                    If SheetExists(CRJ_SHEET_NAME) Then
                        Call modEngine.SyncSJToCRJ(Me, cell.Row)
                    End If
                End If
            End If
NextSync:
        Next cell

CleanupSync:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

Sub ManualSyncToCRJ()
    Dim rowNum As Long
    rowNum = ActiveCell.Row
    If rowNum < START_ROW Then
        MsgBox "Select a row below row 10.", vbExclamation
        Exit Sub
    End If

    ' ---- SKIP TOTAL ROW ----
    If UCase(Trim(Me.Cells(rowNum, 5).Value)) = "TOTAL" Then
        MsgBox "Cannot sync the TOTAL row.", vbExclamation
        Exit Sub
    End If

    If Me.Cells(rowNum, SJ_COL_INVOICE).Value = "" Then
        MsgBox "No invoice number in column U. Enter Amount first.", vbExclamation
        Exit Sub
    End If

    If Not SheetExists(CRJ_SHEET_NAME) Then
        MsgBox "CRJ sheet not found. Cannot sync.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Call modEngine.SyncSJToCRJ(Me, rowNum)
    If Err.Number = 0 Then
        MsgBox "Successfully synced row " & rowNum & " to CRJ.", vbInformation
    Else
        MsgBox "Sync failed: " & Err.description, vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
```

### 6.6 `Sheet8` - Alternate/legacy sales journal event handler.

- Source: `build/v8_5_vba/Sheet8.txt`
- Type: document | Host: SJ2 | CodeName: `Sheet8`
- Lines: 88 | Nonblank: 74 | Procedures: 2
- SHA-256: `4bde4d17ca9ca8bef443767c25df9079c0ce011a33e595d5916c03650e7e1e74`

**Analysis notes**

- Alternate/legacy `SJ2` handler. Implements similar formulas and CRJ sync but uses its own local `SyncToCRJ` wrapper around `modEngine.SyncSJToCRJ`.
- Writes the raw Long invoice number from `modEngine.GetNextInvoiceNumber` into column I, while Sheet3 writes a formatted `SJ-00000` string. This is a behavioural inconsistency between the two SJ modules.
- Checks the TOTAL row via column A rather than column E; confirm the intended SJ2 layout before relying on this module.

```vba
Option Explicit

' ============================================================
' SJ — Worksheet module
' Delegates CRJ sync + invoice numbering to modEngine
' ============================================================

Private Const START_ROW As Long = 10

' --- Column layout (must match modEngine.SJ_COL_* constants) ---
Private Const COL_DATE As Integer = 3           ' C
Private Const COL_CUSTOMER As Integer = 6       ' F
Private Const COL_INVOICE As Integer = 9        ' I
Private Const COL_AMOUNT As Integer = 10        ' J
Private Const COL_DISCOUNT As Integer = 11      ' K
Private Const COL_OUTPUTVAT As Integer = 13     ' M
Private Const COL_NETSALES As Integer = 15      ' O
Private Const COL_REFERENCE As Integer = 21     ' U

' ============================================================
' Sync one SJ row to CRJ (via modEngine)
' ============================================================
Private Sub SyncToCRJ(ByVal sourceRow As Long)
    modEngine.SyncSJToCRJ Me, sourceRow
End Sub

' ============================================================
' MAIN TRIGGER
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range, cell As Range

    ' Watch customer (F), amount (J), discount (K) edits
    Set rng = Intersect(Target, Union(Me.Range("F:F"), Me.Range("J:K")))
    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    For Each cell In rng
        If cell.Row >= START_ROW _
           And UCase$(Trim(CStr(Me.Cells(cell.Row, 1).Value))) <> "TOTAL" Then

            ' ---- Number formats ----
            Me.Cells(cell.Row, COL_AMOUNT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.Row, COL_DISCOUNT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.Row, COL_OUTPUTVAT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.Row, COL_NETSALES).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            ' ---- Formulas (J and K are the amount columns) ----
            ' Output VAT = (Amount - Discount) * 0.107142857
            Me.Cells(cell.Row, COL_OUTPUTVAT).Formula = _
                "=IF(J" & cell.Row & "-K" & cell.Row & "=0,0," & _
                "(J" & cell.Row & "-K" & cell.Row & ")*0.107142857)"

            ' Net Sales = Amount - Discount - Output VAT
            Me.Cells(cell.Row, COL_NETSALES).Formula = _
                "=J" & cell.Row & "-K" & cell.Row & "-M" & cell.Row

            ' ---- Auto-assign invoice number (via engine) ----
            If Me.Cells(cell.Row, COL_INVOICE).Value = "" _
               And Me.Cells(cell.Row, COL_DATE).Value <> "" _
               And Me.Cells(cell.Row, COL_AMOUNT).Value <> "" Then
                Me.Cells(cell.Row, COL_INVOICE).Value = modEngine.GetNextInvoiceNumber()
            End If
        End If
    Next cell

    ' ---- CRJ SYNCHRONIZATION ----
    For Each cell In rng
        If cell.Row >= START_ROW _
           And Me.Cells(cell.Row, COL_INVOICE).Value <> "" _
           And Me.Cells(cell.Row, COL_AMOUNT).Value <> "" Then
            Call SyncToCRJ(cell.Row)
        End If
    Next cell

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

### 6.7 `Sheet10` - Supplier duplicate checker.

- Source: `build/v8_5_vba/Sheet10.txt`
- Type: document | Host: SUPPLIERS DATA | CodeName: `Sheet10`
- Lines: 34 | Nonblank: 31 | Procedures: 1
- SHA-256: `2048d0f487c93b4c620994a3f5a239bf1595a1f0f9895be21b453c2d87f4b0a3`

**Analysis notes**

- Standalone duplicate checker over `SUPPLIERS DATA`: compares supplier Name + Address pairs and reports duplicates in the Immediate Window.
- Has no `Option Explicit` and no dependency on modEngine.
- The address test contains an encoding artifact (`If ... add1 <> "�"`) where the intended value was probably `"-"`; supplier rows with `-` addresses may be compared unintentionally.

```vba
Sub CheckDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim name1 As String, add1 As String
    Dim name2 As String, add2 As String
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    found = False
    
    For i = 2 To lastRow
        name1 = ws.Cells(i, "D").Value
        add1 = ws.Cells(i, "E").Value
        If name1 <> "" And add1 <> "" And add1 <> "—" Then
            For j = i + 1 To lastRow
                name2 = ws.Cells(j, "D").Value
                add2 = ws.Cells(j, "E").Value
                If name1 = name2 And add1 = add2 Then
                    Debug.Print "DUPLICATE: Row " & i & " and Row " & j & " - " & name1 & " | " & add1
                    found = True
                End If
            Next j
        End If
    Next i
    
    If Not found Then
        MsgBox "No duplicates found.", vbInformation
    Else
        MsgBox "Duplicates found! Check Immediate Window (Ctrl+G).", vbExclamation
    End If
End Sub
```

### 6.8 `modEngine` - Core sync/reference/supplier engine.

- Source: `build/v8_5_vba/modEngine.bas`
- Type: standard | Host: (none) | CodeName: `modEngine`
- Lines: 1780 | Nonblank: 1668 | Procedures: 76
- SHA-256: `de50f83cefe3ae876d6e27ae31a248588d35eaa908f0c37089da261194ca486a`

**Analysis notes**

- The largest module (76 procedures) and the centre of the workbook. It owns supplier caches, CDJ/CRJ reference caches, sequence counters, CDJ month-partition placement, PJ/PJ Non-VAT/SJ sync, status/reset helpers and a set of legacy GL routines.
- `InitializeEngine` loads supplier dictionaries, CDJ headers/references, sequence and next row. `InitializeCRJ` loads CRJ references, next row and sequence.
- Supplier lookup is TIN-normalized and fails closed on a TIN shared by different supplier names (`IsAmbiguousTIN`); a cache miss triggers one `ReloadSuppliers` pass.
- `SyncPJToCDJ` and `SyncPJNonVatToCDJ` update an existing row when the reference exists, otherwise reserve a row inside the correct monthly partition via `CDJReserveRow`. The latter inserts above the month TOTAL row when the block is full, then shifts the reference cache.
- `SyncSJToCRJ` updates or appends a CRJ invoice row and updates the CRJ reference cache.
- Legacy/dead code remains in this module: `GLScanGJ`, `Legacy_RefreshGL`, `Legacy_CalculateMonthlyNet`, `Legacy_CalculateEndingBalanceFull`, `Legacy_ValidateGLConsistency`, `Legacy_RefreshAllGL`, `BuildV8CalcSheet`. The active GL path is `modGLScanEngine.ScanGL`; these legacy procedures are not called by it.
- The three sync procedures (`SyncPJToCDJ`, `SyncSJToCRJ`, `SyncPJNonVatToCDJ`) have empty `ErrorHandler:` labels, so unexpected write errors are silently swallowed.
- `CDJReserveRow` inserts a row and restores `Application.EnableEvents` in both success and failure paths; the inserted row gets V/W/X balance formulas.

```vba
Attribute VB_Name = "modEngine"
Rem Attribute VBA_ModuleType=VBAModule
'Option VBASupport 1
Option Explicit

' ============================================================
' modEngine
' Central memory-first engine for PJ->CDJ, SJ->CRJ, and GL auto-aggregation
' Version: 7.9.5
'   7.9.5  Supplier lookup / add-supplier fixes:
'          - Repeated rows of the SAME supplier (same TIN + same name) are no
'            longer treated as ambiguous. Only a TIN shared by DIFFERENT
'            supplier names is flagged, and it still fails closed.
'          - GetSupplierByTIN / GetSupplierByName reload SUPPLIERS DATA once on
'            a miss, so a row added by hand (or an old stale cache) is found.
'          - New AddSupplierRecord(): validated, single-step add that reports
'            success/failure. Sheet modules must use it instead of looping
'            "Add?" -> add -> lookup -> "Add?" ...
'          - Fixed the names list for ambiguous TINs (first name was missing).
'          - Fixed 0-based fallback header array in LoadCDJHeaders.
'          - GetAccountNameFromTIN now compares normalised TINs.
'   7.9.4  GL_AUTO_TOL moved to module top
' ============================================================

' ============================================================
' MODULE-LEVEL CONSTANTS  (MUST be at top, before any Sub/Function)
' ============================================================
Private Const GL_AUTO_TOL As Double = 0.01

' ============================================================
' COLUMN CONSTANTS - PJ
' ============================================================
Private Const PJ_COL_TIN As Integer = 4
Private Const PJ_COL_NAME As Integer = 5
Private Const PJ_COL_ADDRESS As Integer = 6
Private Const PJ_COL_COA As Integer = 7
Private Const PJ_COL_REF As Integer = 8
Private Const PJ_COL_NET As Integer = 10
Private Const PJ_COL_VAT As Integer = 13
Private Const PJ_COL_GROSS As Integer = 14
Private Const PJ_COL_FULLID As Integer = 17

' ============================================================
' COLUMN CONSTANTS - SUPPLIERS DATA
' ============================================================
Private Const SUP_COL_TIN As Integer = 3
Private Const SUP_COL_NAME As Integer = 4
Private Const SUP_COL_ADDRESS As Integer = 5
Private Const SUP_COL_COA As Integer = 6

' ============================================================
' COLUMN CONSTANTS - CDJ
' ============================================================
Private Const CDJ_COL_DATE As Integer = 3
Private Const CDJ_COL_PARTICULARS As Integer = 5
Private Const CDJ_COL_CASH As Integer = 6
Private Const CDJ_COL_VAT As Integer = 7
Private Const CDJ_COL_WHTAX As Integer = 8
Private Const CDJ_COL_REF As Integer = 27
Private Const CDJ_COL_REF_PNV As Integer = 28
Private Const CDJ_COL_REF_LEGACY As Integer = 25
Private Const CDJ_COL_DAY As Integer = 4
Private Const CDJ_COL_NOTE As Integer = 26
Private Const CDJ_COL_SEQ As Integer = 27
Private Const CDJ_COL_NEXTROW As Integer = 27
Private Const CDJ_EXPENSE_START As Integer = 14
Private Const CDJ_EXPENSE_END As Integer = 19
Private Const CDJ_DATA_START As Integer = 15
Private Const CDJ_COL_DEBIT As Integer = 20
Private Const CDJ_ROW_SEQ As Integer = 13
Private Const CDJ_ROW_NEXTROW As Integer = 14

' ============================================================
' COLUMN CONSTANTS - CRJ
' ============================================================
Private Const CRJ_COL_DATE As Integer = 4
Private Const CRJ_COL_CUSTOMER As Integer = 6
Private Const CRJ_COL_INVOICE As Integer = 7
Private Const CRJ_COL_CASH As Integer = 8
Private Const CRJ_COL_OUTPUTVAT As Integer = 10
Private Const CRJ_COL_EXEMPT As Integer = 11
Private Const CRJ_COL_SALES As Integer = 12
Private Const CRJ_COL_NOTE As Integer = 15
Private Const CRJ_COL_REF As Integer = 16
Private Const CRJ_COL_SEQ As Integer = 17
Private Const CRJ_COL_NEXTROW As Integer = 17
Private Const CRJ_ROW_SEQ As Integer = 8
Private Const CRJ_ROW_NEXTROW As Integer = 9
Private Const CRJ_DATA_START As Integer = 10

' ============================================================
' COLUMN CONSTANTS - SJ
' ============================================================
Private Const SJ_COL_DATE As Integer = 3
Private Const SJ_COL_CUSTOMER As Integer = 6
Private Const SJ_COL_AMOUNT As Integer = 10
Private Const SJ_COL_DISCOUNT As Integer = 11
Private Const SJ_COL_OUTPUTVAT As Integer = 13
Private Const SJ_COL_NETSALES As Integer = 15
Private Const SJ_COL_INVOICE As Integer = 9
Private Const SJ_COL_REFERENCE As Integer = 21

' ============================================================
' GLOBAL STATE - CDJ / PJ
' ============================================================
Private dictTIN As Object
Private dictName As Object
Private dictTINDup As Object
Private dictRef As Object
Private arrHeaders As Variant
Private arrCols As Variant
Private gNextSeq As Long
Private gNextRow As Long
Private gInitialized As Boolean
Private gSyncSheet As String
Private gHeaderCount As Integer

' ============================================================
' GLOBAL STATE - CRJ
' ============================================================
Private dictCRJRef As Object
Private gNextCRJRow As Long
Private gNextCRJSeq As Long
Private gInitializedCRJ As Boolean

' ============================================================
' INIT
' ============================================================
Public Sub InitializeEngine(Optional targetSheet As String = "CDJ")
    On Error Resume Next
    If Trim$(targetSheet) = "" Then targetSheet = "CDJ"
    gSyncSheet = targetSheet
    Set dictTIN = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")
    Set dictRef = CreateObject("Scripting.Dictionary")
    Call LoadSuppliers
    Call LoadCDJHeaders
    Call LoadCDJReferences
    Call LoadSequence
    Call LoadNextRow
    gInitialized = True
    On Error GoTo 0
End Sub

Public Sub InitializeCRJ()
    If gInitializedCRJ Then Exit Sub
    Set dictCRJRef = CreateObject("Scripting.Dictionary")
    Call LoadCRJReferences
    Call LoadCRJNextRow
    Call LoadCRJSequence
    gInitializedCRJ = True
End Sub

' ============================================================
' SUPPLIER HELPERS
' ============================================================
' Text of a cell, "" when the cell is empty or holds an error value.
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

' Supplier name normalised for COMPARISON ONLY: case, spacing and
' non-breaking spaces are ignored.
Private Function NormName(ByVal s As String) As String
    s = Replace(s, Chr$(160), " ")
    s = UCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormName = s
End Function

Private Sub NewDupDict()
    Set dictTINDup = CreateObject("Scripting.Dictionary")
    dictTINDup.CompareMode = vbTextCompare
End Sub

Private Function IsDupKey(ByVal key As String) As Boolean
    If dictTINDup Is Nothing Then Exit Function
    IsDupKey = dictTINDup.Exists(key)
End Function

' Single place where a supplier enters the TIN cache.
'   - new TIN                          -> cached
'   - same TIN, SAME supplier name     -> ignored (identical repeat row)
'   - same TIN, DIFFERENT supplier     -> flagged ambiguous (fails closed)
Private Sub RegisterSupplier(ByVal key As String, ByVal tin As String, _
                             ByVal name As String, ByVal address As String, ByVal coa As String)
    Dim firstName As String
    If key = "" Then Exit Sub
    If dictTIN.Exists(key) Then
        firstName = CStr(dictTIN(key)(0))
        If NormName(firstName) <> NormName(name) Then
            If dictTINDup Is Nothing Then Call NewDupDict
            If Not dictTINDup.Exists(key) Then dictTINDup.Add key, firstName
            If InStr(1, "; " & NormName(CStr(dictTINDup(key))) & ";", _
                        "; " & NormName(name) & ";", vbBinaryCompare) = 0 Then
                dictTINDup(key) = dictTINDup(key) & "; " & name
            End If
        End If
    Else
        dictTIN(key) = Array(name, address, coa, tin)
    End If
End Sub

' ============================================================
' LOADERS
' ============================================================
Private Sub LoadSuppliers()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tin As String, key As String, name As String, address As String, coa As String
    Call NewDupDict
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).Row
    For i = 2 To lastRow
        tin = CellText(ws, i, SUP_COL_TIN)
        key = NormTIN(tin)
        name = CellText(ws, i, SUP_COL_NAME)
        address = CellText(ws, i, SUP_COL_ADDRESS)
        coa = CellText(ws, i, SUP_COL_COA)
        If key <> "" Then Call RegisterSupplier(key, tin, name, address, coa)
        If name <> "" Then dictName(NormName(name)) = Array(tin, address, coa)
    Next i
End Sub

' Rebuild the supplier cache from SUPPLIERS DATA without touching the CDJ/CRJ
' state.  Use after SUPPLIERS DATA was edited by hand.
Public Sub ReloadSuppliers()
    Set dictTIN = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")
    Call LoadSuppliers
End Sub

' ============================================================
' TIN NORMALISATION
' ============================================================
' A BIR TIN is 15 digits, normally written NNN-NNN-NNN-NNNN.  The live data
' also contains '009-149--884' (double separator), ' 240-940-697' (leading
' space) and '636-985956' (missing separator).  Keeping only the digits makes
' those match the SUPPLIERS DATA entry they actually belong to.
' ============================================================
Public Function NormTIN(ByVal tin As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(tin) Or IsEmpty(tin) Then Exit Function
    s = Trim$(CStr(tin))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then NormTIN = NormTIN & ch
    Next i
End Function

' Number of TINs that are shared by DIFFERENT supplier names.
Public Function AmbiguousTINCount() As Long
    If dictTINDup Is Nothing Then Exit Function
    AmbiguousTINCount = dictTINDup.count
End Function

' Newline-separated list of the ambiguous TINs and the vendor names sharing
' them, for the operator to correct in SUPPLIERS DATA.
Public Function AmbiguousTINList() As String
    Dim k As Variant, s As String
    If dictTINDup Is Nothing Then Exit Function
    For Each k In dictTINDup.Keys
        s = s & k & " -> " & dictTINDup(k) & vbCrLf
    Next k
    AmbiguousTINList = s
End Function

' True when this TIN is shared by different suppliers in SUPPLIERS DATA.
Public Function IsAmbiguousTIN(ByVal tin As String) As Boolean
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    IsAmbiguousTIN = IsDupKey(NormTIN(tin))
End Function

' The supplier names sharing this TIN ("" when it is not ambiguous).
Public Function AmbiguousNamesFor(ByVal tin As String) As String
    Dim key As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormTIN(tin)
    If IsDupKey(key) Then AmbiguousNamesFor = CStr(dictTINDup(key))
End Function

Private Sub LoadCDJHeaders()
    Dim ws As Worksheet
    Dim col As Integer, header As String, count As Integer
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then count = count + 1
    Next col
    If count = 0 Then
        ' 1-based, same shape as the live-header branch below.
        ReDim arrHeaders(1 To 6)
        ReDim arrCols(1 To 6)
        arrHeaders(1) = "clinic supplies": arrCols(1) = 14
        arrHeaders(2) = "rent":            arrCols(2) = 15
        arrHeaders(3) = "fuel":            arrCols(3) = 16
        arrHeaders(4) = "utilities":       arrCols(4) = 17
        arrHeaders(5) = "professional":    arrCols(5) = 18
        arrHeaders(6) = "sundry":          arrCols(6) = 19
        gHeaderCount = 6
        Exit Sub
    End If
    ReDim arrHeaders(1 To count)
    ReDim arrCols(1 To count)
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then
            count = count + 1
            arrHeaders(count) = LCase(header)
            arrCols(count) = col
        End If
    Next col
    gHeaderCount = count
End Sub

Private Sub LoadCDJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, ref As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
    For i = CDJ_DATA_START To lastRow
        ref = Trim(ws.Cells(i, CDJ_COL_REF).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)
        If ref <> "" And Not dictRef.Exists(ref) Then dictRef(ref) = i
    Next i
End Sub

' ============================================================
' CDJ MONTH-BLOCK PLACEMENT
' ============================================================
' CDJ is laid out as six fixed monthly partitions, each closed by a TOTAL row
' whose columns F..U are =SUM(<block>).  A row written outside its partition is
' invisible to every monthly total, which is the reported defect: inserts were
' landing at row 755, beyond the last partition.  These helpers locate the
' partition and, when it is full, insert a row immediately above the TOTAL row
' so Excel widens the SUM range automatically.
' ============================================================
Private Function CDJRowIsFree(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_PARTICULARS).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_CASH).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF_PNV).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF_LEGACY).Value))) > 0 Then Exit Function
    CDJRowIsFree = True
End Function

' Month carried by a col-C block marker such as "JANUARY" or "FEB TOTAL".
Public Function CDJMarkerMonth(ByVal marker As String) As Integer
    Dim t As String
    t = UCase$(Trim$(marker))
    If t = "" Then Exit Function
    If InStr(t, "TOTAL") > 0 Then t = Trim$(Replace(t, "TOTAL", ""))
    CDJMarkerMonth = V8_Month(Left$(t, 3))
End Function

Public Function CDJMonthAbbr(ByVal m As Integer) As String
    Select Case m
        Case 1: CDJMonthAbbr = "JAN"
        Case 2: CDJMonthAbbr = "FEB"
        Case 3: CDJMonthAbbr = "MAR"
        Case 4: CDJMonthAbbr = "APR"
        Case 5: CDJMonthAbbr = "MAY"
        Case 6: CDJMonthAbbr = "JUN"
        Case 7: CDJMonthAbbr = "JUL"
        Case 8: CDJMonthAbbr = "AUG"
        Case 9: CDJMonthAbbr = "SEP"
        Case 10: CDJMonthAbbr = "OCT"
        Case 11: CDJMonthAbbr = "NOV"
        Case 12: CDJMonthAbbr = "DEC"
    End Select
End Function

' Row of the TOTAL row that closes month mth, or 0 when the block is absent.
Public Function CDJTotalRow(ByVal ws As Worksheet, ByVal mth As Integer) As Long
    Dim r As Long, lab As String
    If mth < 1 Or mth > 12 Then Exit Function
    For r = CDJ_DATA_START To ws.Cells(ws.Rows.count, CDJ_COL_DATE).End(xlUp).Row
        lab = UCase$(Trim$(CStr(ws.Cells(r, CDJ_COL_DATE).Value)))
        If InStr(lab, "TOTAL") > 0 Then
            If CDJMarkerMonth(lab) = mth Then CDJTotalRow = r: Exit Function
        End If
    Next r
End Function

' First empty row inside month mth's own partition; 0 when the block is full.
Public Function FindCDJFreeRow(Optional ByVal targetMonth As Integer = 0) As Long
    Dim ws As Worksheet, r As Long, cur As Long, totalRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then FindCDJFreeRow = CDJ_DATA_START: Exit Function
    If targetMonth > 0 Then
        totalRow = CDJTotalRow(ws, targetMonth)
        If totalRow = 0 Then Exit Function
    Else
        totalRow = ws.Cells(ws.Rows.count, CDJ_COL_PARTICULARS).End(xlUp).Row + 1
    End If
    cur = 0
    For r = CDJ_DATA_START To totalRow - 1
        If CDJMarkerMonth(CStr(ws.Cells(r, CDJ_COL_DATE).Value)) > 0 Then _
            cur = CDJMarkerMonth(CStr(ws.Cells(r, CDJ_COL_DATE).Value))
        If targetMonth > 0 And cur <> targetMonth Then GoTo NextFreeRow
        If CDJRowIsFree(ws, r) Then FindCDJFreeRow = r: Exit Function
NextFreeRow:
    Next r
End Function

' Reserve the row a new CDJ entry for month mth must be written to.  When the
' partition has no free row, insert directly above its TOTAL row so the TOTAL's
' SUM range widens over the new row.  Returns 0 when no placement is possible,
' in which case the caller must NOT write anything.
Public Function CDJReserveRow(ByVal ws As Worksheet, ByVal mth As Integer) As Long
    Dim atRow As Long, r As Long
    atRow = FindCDJFreeRow(mth)
    If atRow > 0 Then CDJReserveRow = atRow: Exit Function
    If mth < 1 Or mth > 12 Then Exit Function
    atRow = CDJTotalRow(ws, mth)
    If atRow = 0 Then Exit Function
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Rows(atRow).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' The inserted row does not inherit the row-level balance formulas.
    ws.Cells(atRow, 22).Formula = "=F" & atRow
    ws.Cells(atRow, 23).Formula = "=SUM(G" & atRow & ":T" & atRow & ")"
    ws.Cells(atRow, 24).Formula = "=V" & atRow & "+W" & atRow
    If atRow > CDJ_DATA_START Then
        For r = 14 To 21
            ws.Cells(atRow, r).NumberFormat = ws.Cells(atRow - 1, r).NumberFormat
        Next r
    End If
    ' Row insertion invalidates every cached row number at or below it.
    Call ShiftCDJRefCache(atRow)
    CDJReserveRow = atRow
Done:
    Application.EnableEvents = True
    Exit Function
Failed:
    CDJReserveRow = 0
    Resume Done
End Function

Private Sub ShiftCDJRefCache(ByVal fromRow As Long)
    Dim k As Variant
    For Each k In dictRef.Keys
        If CLng(dictRef(k)) >= fromRow Then dictRef(k) = CLng(dictRef(k)) + 1
    Next k
End Sub

' Month of a PJ / PJ Non-Vat source row, taken from the sheet's own col-B
' month markers (col C holds only a day number and is blank on most rows).
Public Function PJMonthOf(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long) As Integer
    Dim r As Long, mk As Integer
    For r = sourceRow To 1 Step -1
        If UCase$(Trim$(CStr(sourceSheet.Cells(r, 4).Value))) = "TOTAL" Then Exit Function
        mk = CDJMarkerMonth(CStr(sourceSheet.Cells(r, 2).Value))
        If mk > 0 Then PJMonthOf = mk: Exit Function
    Next r
End Function

Private Sub LoadSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then gNextSeq = 0: Exit Sub
    val = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextSeq = CLng(val) Else gNextSeq = 0
End Sub

Private Sub LoadNextRow()
    ' The static pointer in AA14 is advisory only: Excel never adjusts it when
    ' rows are inserted, so it drifts onto live rows.  It currently reads 755,
    ' past the last partition, which is why new rows missed every monthly SUM.
    ' Resolve the next write row by scanning for a genuinely empty row instead.
    gNextRow = FindCDJFreeRow(0)
End Sub

Private Sub LoadCRJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, inv As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    For i = CRJ_DATA_START To lastRow
        inv = Trim(ws.Cells(i, CRJ_COL_INVOICE).Value)
        If inv <> "" Then dictCRJRef(inv) = i
    Next i
End Sub

Private Sub LoadCRJNextRow()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJRow = CRJ_DATA_START: Exit Sub
    val = ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value
    If IsNumeric(val) Then
        gNextCRJRow = CLng(val)
    Else
        gNextCRJRow = CRJ_DATA_START
        Do While ws.Cells(gNextCRJRow, CRJ_COL_INVOICE).Value <> ""
            gNextCRJRow = gNextCRJRow + 1
        Loop
    End If
End Sub

Private Sub LoadCRJSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJSeq = 0: Exit Sub
    val = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextCRJSeq = CLng(val) Else gNextCRJSeq = 0
End Sub

' ============================================================
' SEQUENCE
' ============================================================
Public Function GetNextPJNumber() As String
    Dim ws As Worksheet
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    gNextSeq = gNextSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    If Not ws Is Nothing Then ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = gNextSeq
    On Error GoTo 0
    GetNextPJNumber = "PJ-" & Format(gNextSeq, "00000")
End Function

Public Function GetNextInvoiceNumber() As Long
    Dim ws As Worksheet
    If Not gInitializedCRJ Then Call InitializeCRJ
    gNextCRJSeq = gNextCRJSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    If Not ws Is Nothing Then ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = gNextCRJSeq
    On Error GoTo 0
    GetNextInvoiceNumber = gNextCRJSeq
End Function

Public Function GetNextSeq() As Long
    Dim s As String
    s = GetNextPJNumber()
    If Len(s) > 3 Then GetNextSeq = CLng(Mid$(s, 4)) Else GetNextSeq = 0
End Function

' ============================================================
' SUPPLIER LOOKUPS
' ============================================================
' Returns Array(name, address, coa, tin) or False.
' False means: the TIN is not in SUPPLIERS DATA, OR it is shared by different
' suppliers (use IsAmbiguousTIN to tell the two apart).
' On a miss the cache is rebuilt from the sheet ONCE and the lookup repeated,
' so rows added or corrected by hand are picked up without reopening the file.
Public Function GetSupplierByTIN(tin As String) As Variant
    Dim key As String, pass As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormTIN(tin)
    If key = "" Then GetSupplierByTIN = False: Exit Function
    For pass = 1 To 2
        If dictTIN.Exists(key) And Not IsDupKey(key) Then
            GetSupplierByTIN = dictTIN(key)
            Exit Function
        End If
        If pass = 1 Then Call ReloadSuppliers
    Next pass
    GetSupplierByTIN = False
End Function

' Returns Array(tin, address, coa) or False.  Case and spacing are ignored.
Public Function GetSupplierByName(name As String) As Variant
    Dim key As String, pass As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormName(name)
    If key = "" Then GetSupplierByName = False: Exit Function
    For pass = 1 To 2
        If dictName.Exists(key) Then
            GetSupplierByName = dictName(key)
            Exit Function
        End If
        If pass = 1 Then Call ReloadSuppliers
    Next pass
    GetSupplierByName = False
End Function

' Cache-only add (kept for existing callers).  Prefer AddSupplierRecord, which
' also writes the row to SUPPLIERS DATA and verifies the result.
Public Sub AddSupplierToCache(tin As String, name As String, address As String, coa As String)
    Dim key As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    tin = Trim(tin): name = Trim(name)
    key = NormTIN(tin)
    If key <> "" Then Call RegisterSupplier(key, tin, name, address, coa)
    If NormName(name) <> "" Then dictName(NormName(name)) = Array(tin, address, coa)
End Sub

' Validated, single-step "add supplier":
'   1. refuses a TIN that already belongs to a DIFFERENT supplier,
'   2. treats a TIN that already exists for the SAME supplier as success,
'   3. writes the new row to SUPPLIERS DATA (TIN stored as text),
'   4. updates the cache and CONFIRMS the TIN now resolves.
' Returns True only when the supplier can be looked up afterwards.
' msg always explains the outcome; callers must not retry in a loop.
Public Function AddSupplierRecord(ByVal tin As String, ByVal name As String, _
                                  ByVal address As String, ByVal coa As String, _
                                  ByRef msg As String) As Boolean
    Dim ws As Worksheet, key As String, r As Long
    Dim lastB As Long, lastC As Long, newRow As Long, nextNo As Double
    Dim existName As String

    tin = Trim$(tin): name = Trim$(name): address = Trim$(address): coa = Trim$(coa)
    msg = ""
    key = NormTIN(tin)
    If key = "" Then msg = "The TIN has no digits.": Exit Function
    If name = "" Then msg = "Supplier name is required.": Exit Function
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    On Error GoTo 0
    If ws Is Nothing Then msg = "Sheet 'SUPPLIERS DATA' not found.": Exit Function

    lastC = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).Row

    ' --- 1/2: does this TIN already exist in the sheet? ---
    For r = 2 To lastC
        If NormTIN(CellText(ws, r, SUP_COL_TIN)) = key Then
            existName = CellText(ws, r, SUP_COL_NAME)
            If NormName(existName) = NormName(name) Then
                Call ReloadSuppliers
                If IsArray(GetSupplierByTIN(tin)) Then
                    AddSupplierRecord = True
                    msg = "Supplier already exists in SUPPLIERS DATA (row " & r & ")."
                Else
                    msg = "TIN " & tin & " exists (row " & r & ") but is also assigned to a different " & _
                          "supplier: " & AmbiguousNamesFor(tin) & ". Correct SUPPLIERS DATA."
                End If
            Else
                msg = "TIN " & tin & " already belongs to '" & existName & "' (row " & r & "). " & _
                      "Not added - correct SUPPLIERS DATA if this is a different supplier."
            End If
            Exit Function
        End If
    Next r

    ' --- 3: write the new row ---
    lastB = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    If lastC > lastB Then newRow = lastC + 1 Else newRow = lastB + 1
    If newRow < 5 Then newRow = 5
    nextNo = 0
    For r = 5 To newRow - 1
        If IsNumeric(ws.Cells(r, 2).Value) And Not IsEmpty(ws.Cells(r, 2).Value) Then
            If CDbl(ws.Cells(r, 2).Value) > nextNo Then nextNo = CDbl(ws.Cells(r, 2).Value)
        End If
    Next r

    On Error GoTo WriteFailed
    ws.Cells(newRow, SUP_COL_TIN).NumberFormat = "@"
    ws.Cells(newRow, 2).Value = nextNo + 1
    ws.Cells(newRow, SUP_COL_TIN).Value = tin
    ws.Cells(newRow, SUP_COL_NAME).Value = UCase$(name)
    ws.Cells(newRow, SUP_COL_ADDRESS).Value = address
    ws.Cells(newRow, SUP_COL_COA).Value = coa
    On Error GoTo 0

    ' --- 4: update the cache and confirm ---
    Call RegisterSupplier(key, tin, UCase$(name), address, coa)
    dictName(NormName(name)) = Array(tin, address, coa)
    If IsArray(GetSupplierByTIN(tin)) Then
        AddSupplierRecord = True
        msg = "Supplier added (SUPPLIERS DATA row " & newRow & ")."
    Else
        msg = "The supplier was written to SUPPLIERS DATA row " & newRow & " but still cannot be " & _
              "resolved. Check that row."
    End If
    Exit Function

WriteFailed:
    msg = "Could not write to SUPPLIERS DATA: " & Err.description & _
          " (is the sheet protected?)"
    Err.Clear
End Function

' ============================================================
' EXPENSE COLUMN MAPPING
' ============================================================
Public Function FindExpenseColumn(description As String) As Integer
    Dim descLower As String, i As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    descLower = LCase(Trim(description))
    If descLower = "" Then FindExpenseColumn = 0: Exit Function
    
    ' --- Pass 1: live CDJ header substring match ---
    For i = 1 To gHeaderCount
        If arrHeaders(i) <> "" Then
            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then
                FindExpenseColumn = arrCols(i)
                Exit Function
            End If
        End If
    Next i
    
    ' --- Pass 2: COA alias mapping to CDJ columns (14-19) ---
    Select Case True
        Case InStr(descLower, "clinic") > 0 Or InStr(descLower, "medical") > 0 _
          Or (InStr(descLower, "suppl") > 0 And InStr(descLower, "pantry") = 0)
            FindExpenseColumn = CDJ_EXPENSE_START       ' Col 14: Clinic Materials & Supplies
            
        Case InStr(descLower, "rent") > 0 Or InStr(descLower, "lease") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 1   ' Col 15: Rent Expense
            
        Case InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 _
          Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 _
          Or InStr(descLower, "transport") > 0 Or InStr(descLower, "travel") > 0 _
          Or InStr(descLower, "delivery") > 0 Or InStr(descLower, "parking") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 2   ' Col 16: Gas, Oil, Fuel, Toll, Parking
            
        Case InStr(descLower, "communication") > 0 Or InStr(descLower, "light") > 0 _
          Or InStr(descLower, "water") > 0 Or InStr(descLower, "utilit") > 0 _
          Or InStr(descLower, "electric") > 0 Or InStr(descLower, "internet") > 0 _
          Or InStr(descLower, "power") > 0 Or InStr(descLower, "phone") > 0 Or InStr(descLower, "tel") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 3   ' Col 17: Communication, Light and Water
            
        Case InStr(descLower, "professional") > 0 Or InStr(descLower, "consult") > 0 _
          Or InStr(descLower, "clinician") > 0 Or InStr(descLower, "doctor") > 0 _
          Or InStr(descLower, "legal") > 0 Or InStr(descLower, "audit") > 0 _
          Or InStr(descLower, "accounting") > 0 Or InStr(descLower, "retainer") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 4   ' Col 18: Professional Fees
            
        Case Else
            ' All other COAs (Miscellaneous, Representation, Pantry, Taxes, Repairs, Dues, etc.) -> Sundry
            FindExpenseColumn = CDJ_EXPENSE_END         ' Col 19: Sundry Account
    End Select
End Function

' ============================================================
' WITHHOLDING TAX
' ============================================================
Public Function CalculateWithholdingTax(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateWithholdingTax = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateWithholdingTax = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateWithholdingTax = amount * 0.02
        Case Else
            CalculateWithholdingTax = 0
    End Select
End Function

' ============================================================
' REFERENCE CACHE
' ============================================================
Public Function FindCDJRow(ref As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim$(ref)
    If ref = "" Then FindCDJRow = 0: Exit Function
    If dictRef.Exists(ref) Then
        FindCDJRow = dictRef(ref)
        Exit Function
    End If
    ' Fallback scan of CDJ sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
        If lastRow >= CDJ_DATA_START Then
            For i = CDJ_DATA_START To lastRow
                If Trim$(CStr(ws.Cells(i, CDJ_COL_REF).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_PNV).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)) = ref Then
                    dictRef(ref) = i
                    FindCDJRow = i
                    Exit Function
                End If
            Next i
        End If
    End If
    FindCDJRow = 0
End Function

Public Sub UpdateCDJReference(ref As String, rowNum As Long)
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim(ref)
    If ref <> "" And rowNum > 0 Then dictRef(ref) = rowNum
End Sub

Public Function FindCRJRow(invoiceNum As String) As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum = "" Then FindCRJRow = 0: Exit Function
    If dictCRJRef.Exists(invoiceNum) Then FindCRJRow = dictCRJRef(invoiceNum) Else FindCRJRow = 0
End Function

Public Sub UpdateCRJReference(invoiceNum As String, rowNum As Long)
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum <> "" And rowNum > 0 Then dictCRJRef(invoiceNum) = rowNum
End Sub

' ============================================================
' SYNC PJ -> CDJ
' ============================================================
Public Sub SyncPJToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer, mth As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value) Then
        grossAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value)
    Else
        grossAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value)
    Else
        netAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value) Then
        vatAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value)
    Else
        vatAmount = 0
    End If
    If grossAmount = 0 Then Exit Sub
    expenseCol = FindExpenseColumn(description)
    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
        whtax = CalculateWithholdingTax(description, netAmount)
    Else
        whtax = 0
    End If
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            ' Clear every allocation column including T (20).  The old loop
            ' stopped at S (19), leaving a stale Sundry Debit behind whenever a
            ' row was reclassified away from Sundry.
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18 take the net expense only.  Column T
                ' (Sundry Debit) stays empty so the amount is not counted twice.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: COA label in S (19), net amount in T (20).  Column R
                ' (18 = Professional Fees) stays empty so the amount is never
                ' classified as Professional Fees as well as Sundry.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        ' Place the new row inside its own monthly partition so the partition's
        ' TOTAL row includes it.  Writing to gNextRow (755) put every insert
        ' beyond the last SUM range, which is the reported total defect.
        mth = PJMonthOf(sourceSheet, sourceRow)
        If mth = 0 Then mth = CDJMarkerMonth(CStr(dateVal))
        targetRow = CDJReserveRow(wsTarget, mth)
        If targetRow = 0 Then Exit Sub   ' fail closed: no legal placement
        gNextRow = targetRow
        If mth > 0 And Len(Trim$(CStr(dateVal))) = 0 Then dateVal = CDJMonthAbbr(mth)
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18 take the net expense only.  Column T
                ' (Sundry Debit) stays empty so the amount is not counted twice.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: COA label in S (19), net amount in T (20).  Column R
                ' (18 = Professional Fees) stays empty so the amount is never
                ' classified as Professional Fees as well as Sundry.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' SYNC SJ -> CRJ
' ============================================================
Public Sub SyncSJToCRJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim invoiceNum As Variant, targetRow As Long
    Dim referenceInvoiceNum As Variant
    Dim customerName As String
    Dim dateVal As Variant
    Dim amount As Double, discount As Double, outputVAT As Double, netSales As Double
    Dim cashInBank As Double
    On Error GoTo ErrorHandler
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    invoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_INVOICE).Value
    If invoiceNum = "" Or invoiceNum = "TOTAL" Then Exit Sub
    customerName = Trim(sourceSheet.Cells(sourceRow, SJ_COL_CUSTOMER).Value)
    dateVal = sourceSheet.Cells(sourceRow, SJ_COL_DATE).Value
    referenceInvoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_REFERENCE).Value
    amount = sourceSheet.Cells(sourceRow, SJ_COL_AMOUNT).Value
    discount = sourceSheet.Cells(sourceRow, SJ_COL_DISCOUNT).Value
    outputVAT = sourceSheet.Cells(sourceRow, SJ_COL_OUTPUTVAT).Value
    netSales = sourceSheet.Cells(sourceRow, SJ_COL_NETSALES).Value
    cashInBank = amount - discount
    If customerName = "" Or amount = 0 Or Not IsNumeric(amount) Then Exit Sub
    targetRow = FindCRJRow(CStr(invoiceNum))
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Updated from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
        End With
    Else
        targetRow = gNextCRJRow
        If targetRow < CRJ_DATA_START Then
            targetRow = CRJ_DATA_START
            Do While wsTarget.Cells(targetRow, CRJ_COL_INVOICE).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextCRJRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Inserted from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
            .Cells(targetRow, CRJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_EXEMPT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_SALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCRJReference CStr(invoiceNum), targetRow
        gNextCRJRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' STATUS
' ============================================================
Public Function EngineStatus() As String
    If gInitialized Then
        EngineStatus = "Initialized | Suppliers: " & dictTIN.count & _
                       " | Ambiguous TINs: " & AmbiguousTINCount() & _
                       " | CDJ Refs: " & dictRef.count & _
                       " | Headers: " & gHeaderCount & _
                       " | Next PJ Seq: " & gNextSeq & _
                       " | Next CDJ Row: " & gNextRow & _
                       " | Target: " & gSyncSheet
    Else
        EngineStatus = "NOT INITIALIZED"
    End If
    If gInitializedCRJ Then
        EngineStatus = EngineStatus & " || CRJ Refs: " & dictCRJRef.count & _
                       " | Next CRJ Seq: " & gNextCRJSeq & _
                       " | Next CRJ Row: " & gNextCRJRow
    End If
End Function

Public Sub ShowCDJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Target sheet '" & gSyncSheet & "' not found!", vbCritical: Exit Sub
    msg = "===== CDJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Target Sheet: " & gSyncSheet & vbCrLf
    msg = msg & "Data starts at row: " & CDJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Z14): " & gNextRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow < CDJ_DATA_START Then
        msg = msg & "Last used row (with reference): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "References in cache: " & dictRef.count & vbCrLf
    msg = msg & "Sequence (Z13): " & IIf(IsNumeric(ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value), ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value, "empty") & vbCrLf
    msg = msg & "Next PJ number: " & Format(gNextSeq + 1, "PJ-00000") & vbCrLf
    msg = msg & vbCrLf & "Expense headers (row 14, N-T):" & vbCrLf
    For i = 1 To gHeaderCount
        msg = msg & "  Col " & arrCols(i) & ": " & arrHeaders(i) & vbCrLf
    Next i
    MsgBox msg, vbInformation, "CDJ Status"
End Sub

Public Sub ShowCRJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    msg = "===== CRJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Data starts at row: " & CRJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Q9): " & gNextCRJRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow < CRJ_DATA_START Then
        msg = msg & "Last used row (with invoice): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "Invoice references in cache: " & dictCRJRef.count & vbCrLf
    msg = msg & "Invoice sequence (Q8): " & gNextCRJSeq & vbCrLf
    msg = msg & "Next invoice number will be: " & Format(gNextCRJSeq + 1, "00000") & vbCrLf
    MsgBox msg, vbInformation, "CRJ Status"
End Sub

' Diagnostic: shows every TIN that is shared by DIFFERENT supplier names.
Public Sub ShowAmbiguousTINs()
    Dim s As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    Call ReloadSuppliers
    s = AmbiguousTINList()
    If s = "" Then
        MsgBox "No TIN is shared by different suppliers.", vbInformation, "Ambiguous TINs"
    Else
        MsgBox AmbiguousTINCount() & " TIN(s) shared by different suppliers:" & vbCrLf & vbCrLf & s, _
               vbExclamation, "Ambiguous TINs"
    End If
End Sub

' ============================================================
' REPAIR / RESET
' ============================================================
Public Sub RepairEngine()
    Dim ws As Worksheet, lastRow As Long, seqVal As Variant
    gSyncSheet = "CDJ"
    Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If lastRow < CDJ_DATA_START Then lastRow = CDJ_DATA_START - 1
        gNextRow = lastRow + 1
        ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        seqVal = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then gNextSeq = CLng(seqVal) Else gNextSeq = 0: ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
        Call LoadCDJReferences
    End If
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
        If lastRow < CRJ_DATA_START Then lastRow = CRJ_DATA_START - 1
        gNextCRJRow = lastRow + 1
        ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        seqVal = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then
            gNextCRJSeq = CLng(seqVal)
        Else
            gNextCRJSeq = 0
            ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
        End If
        Call LoadCRJReferences
    End If
    MsgBox "Engine repaired!" & vbCrLf & vbCrLf & EngineStatus, vbInformation
End Sub

Public Sub ResetCRJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CRJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CRJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow >= CRJ_DATA_START Then ws.Rows(CRJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
    gNextCRJSeq = 0
    ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = CRJ_DATA_START
    gNextCRJRow = CRJ_DATA_START
    If Not gInitializedCRJ Then Call InitializeCRJ
    dictCRJRef.RemoveAll
    gInitializedCRJ = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "CRJ has been reset." & vbCrLf & _
           "Sequence (Q8) = 0, Next row (Q9) = " & CRJ_DATA_START, vbInformation
End Sub

Public Sub ResetCDJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CDJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CDJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CDJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow >= CDJ_DATA_START Then ws.Rows(CDJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
    gNextSeq = 0
    ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = CDJ_DATA_START
    gNextRow = CDJ_DATA_START
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    dictRef.RemoveAll
    Call LoadCDJReferences
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "CDJ has been reset." & vbCrLf & _
           "Sequence (Z13) = 0, Next row (Z14) = " & CDJ_DATA_START, vbInformation
End Sub

' ============================================================
' EWT (extended)
' ============================================================
Public Function CalculateEWT(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateEWT = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateEWT = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateEWT = amount * 0.02
        Case InStr(expLower, "supply") > 0 Or InStr(expLower, "office") > 0
            CalculateEWT = amount * 0.01
        Case Else
            CalculateEWT = 0
    End Select
End Function

' ============================================================
' SYNC PJ NON-VAT -> CDJ
' ============================================================
Public Sub SyncPJNonVatToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim netAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer, mth As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, 14).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, 14).Value)
    Else
        netAmount = 0
    End If
    If netAmount = 0 Then Exit Sub
    whtax = CalculateEWT(description, netAmount)
    expenseCol = FindExpenseColumn(description)
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            ' Non-VAT row: gross = net (no input VAT), so the row balances as
            '  -(net - EWT) + 0 - EWT + net = 0.
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        ' Same placement rule as the VAT path: inside the month's own partition.
        mth = CDJMarkerMonth(CStr(dateVal))
        If mth = 0 Then mth = PJMonthOf(sourceSheet, sourceRow)
        targetRow = CDJReserveRow(wsTarget, mth)
        If targetRow = 0 Then Exit Sub   ' fail closed: no legal placement
        gNextRow = targetRow
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' GL AUTO-AGGREGATION  (no Const here - it lives at the top)
' ============================================================
Private Function GLN(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "'", "")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    GLN = s
End Function

Private Function GLAcct(ByVal s As String) As String
    Dim k As String
    k = GLN(s)
    Select Case k
        Case "cash in bank", "cash": GLAcct = "Cash in Bank"
        Case "petty cash fund", "petty cash": GLAcct = "Petty Cash Fund"
        Case "account receivables", "accounts receivable", "account receivable": GLAcct = "Account Receivables"
        Case "advances to employee", "advances to employees": GLAcct = "Advances to Employees"
        Case "input vat", "input tax": GLAcct = "Input VAT"
        Case "excess cwt over it", "cwt/it", "cwt": GLAcct = "Excess CWT Over IT"
        Case "leasehold improvements": GLAcct = "Leasehold Improvements"
        Case "dental equipment", "medical equipment": GLAcct = "Dental Equipment"
        Case "due to clinicians": GLAcct = "Due To Clinicians"
        Case "due to clinicians- visiting", "due to clinicians visiting": GLAcct = "Due to Clinicians- Visiting"
        Case "vat payable", "output vat payable", "output vat": GLAcct = "VAT Payable"
        Case "ewt payable", "withholding tax", "withholding tax payable", "ewt": GLAcct = "EWT Payable"
        Case "government contributions (ee)", "ee government contributions": GLAcct = "Government Contributions (EE)"
        Case "government loans (ee)", "ee government loans": GLAcct = "Government Loans (EE)"
        Case "income tax payable": GLAcct = "Income Tax Payable"
        Case "paid-up capital", "paid up capital": GLAcct = "Paid-Up Capital"
        Case "retained earnings (deficit)", "retained earnings": GLAcct = "Retained Earnings (Deficit)"
        Case "sales", "exempt sales": GLAcct = "Sales"
        Case "rent", "rental": GLAcct = "Rent"
        Case "clinician's fee-corporators", "clinicians fee-corporators": GLAcct = "Clinician's Fee-Corporators"
        Case "clinician's fee-visiting", "clinicians fee-visiting": GLAcct = "Clinician's Fee-Visiting"
        Case "clinic material and supplies", "clinic materials and supplies", "clinic supplies": GLAcct = "Clinic Material and Supplies"
        Case "common use service area": GLAcct = "Common Use Service Area"
        Case "light and water expense", "light and water": GLAcct = "Light and Water Expense"
        Case "dep'n expense-leasehold improvements", "depn expense-leasehold improvements": GLAcct = "Dep'n Expense-Leasehold Improvements"
        Case "dep'n expense-dental equipment", "depn expense-dental equipment": GLAcct = "Dep'n Expense-Dental Equipment"
        Case "salaries and wages", "salary and wages": GLAcct = "Salaries and Wages"
        Case "de minimis": GLAcct = "De Minimis"
        Case "government contributions er share", "government contributions er": GLAcct = "Government Contributions ER Share"
        Case "professional fees", "professional fee": GLAcct = "Professional Fees"
        Case "repair and maintenance", "repairs and maintenance": GLAcct = "Repair and Maintenance"
        Case "pantry supplies": GLAcct = "Pantry Supplies"
        Case "office supplies", "supplies": GLAcct = "Office Supplies"
        Case "meals, gifts & entertainment expense", "meals and entertainment": GLAcct = "Meals, Gifts & Entertainment Expense"
        Case "communication": GLAcct = "Communication "
        Case "printing and duplication": GLAcct = "Printing and Duplication"
        Case "taxes and licences", "taxes and licenses": GLAcct = "Taxes and Licences"
        Case "delivery": GLAcct = "Delivery"
        Case "software subscription": GLAcct = "Software Subscription"
        Case "miscellaneous", "charges", "bank charge": GLAcct = "Miscellaneous"
        Case "transportation and travel", "transporation and travel": GLAcct = "Transportation and Travel"
        Case "advertisement": GLAcct = "Advertisement"
        Case "marketing expense": GLAcct = "Marketing Expense"
        Case "gas, oil, parking, toll fees", "fuel and oil": GLAcct = "Gas, Oil, Parking, Toll Fees"
        Case "other penalties and charges": GLAcct = "Other Penalties and Charges"
        Case "insurance": GLAcct = "Insurance"
    End Select
End Function

Private Function GLMonth(ByVal v As Variant) As Long
    Dim k As String
    k = GLN(CStr(v))
    Select Case Left$(k, 3)
        Case "jan": GLMonth = 1
        Case "feb": GLMonth = 2
        Case "mar": GLMonth = 3
        Case "apr": GLMonth = 4
        Case "may": GLMonth = 5
        Case "jun": GLMonth = 6
        Case "jul": GLMonth = 7
        Case "aug": GLMonth = 8
        Case "sep": GLMonth = 9
        Case "oct": GLMonth = 10
        Case "nov": GLMonth = 11
        Case "dec": GLMonth = 12
        Case Else: If IsDate(v) Then GLMonth = month(CDate(v))
    End Select
End Function

Private Sub GLAdd(ByRef d As Object, ByVal a As String, ByVal m As Long, ByVal dr As Double, ByVal cr As Double)
    If a = "" Or m < 1 Or m > 12 Then Exit Sub
    Dim k As String: k = a & "|" & m
    Dim x As Variant
    If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
    x(0) = x(0) + dr
    x(1) = x(1) + cr
    d(k) = x
End Sub

Private Sub GLAudit(ByVal aw As Worksheet, ByRef r As Long, ByVal src As String, ByVal detail As String, ByVal amt As Double)
    aw.Cells(r, 1).Value = src
    aw.Cells(r, 2).Value = detail
    aw.Cells(r, 3).Value = amt
    r = r + 1
End Sub

Private Sub GLScanCDJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CDJ")
    Dim r As Long, m As Long, a As String, v As Double
    For r = 15 To w.Cells(w.Rows.count, 3).End(xlUp).Row
        If GLN(w.Cells(r, 5).Value) <> "total" Then
            m = GLMonth(w.Cells(r, 3).Value)
            If m > 0 Then
                If IsNumeric(w.Cells(r, 6).Value) Then
                    v = CDbl(w.Cells(r, 6).Value)
                    If v < 0 Then
                        GLAdd d, "Cash in Bank", m, 0, -v
                    Else
                        GLAdd d, "Cash in Bank", m, v, 0
                    End If
                End If
                If IsNumeric(w.Cells(r, 7).Value) Then GLAdd d, "Input VAT", m, CDbl(w.Cells(r, 7).Value), 0
                If IsNumeric(w.Cells(r, 8).Value) Then GLAdd d, "EWT Payable", m, 0, CDbl(w.Cells(r, 8).Value)
                If IsNumeric(w.Cells(r, 9).Value) Then GLAdd d, "Petty Cash Fund", m, CDbl(w.Cells(r, 9).Value), 0
                If IsNumeric(w.Cells(r, 10).Value) Then GLAdd d, "Government Contributions (EE)", m, 0, CDbl(w.Cells(r, 10).Value)
                If IsNumeric(w.Cells(r, 11).Value) Then GLAdd d, "Government Loans (EE)", m, 0, CDbl(w.Cells(r, 11).Value)
                If IsNumeric(w.Cells(r, 12).Value) Then GLAdd d, "De Minimis", m, CDbl(w.Cells(r, 12).Value), 0
                If IsNumeric(w.Cells(r, 13).Value) Then GLAdd d, "Salaries and Wages", m, CDbl(w.Cells(r, 13).Value), 0
                If IsNumeric(w.Cells(r, 14).Value) Then GLAdd d, "Clinic Material and Supplies", m, CDbl(w.Cells(r, 14).Value), 0
                If IsNumeric(w.Cells(r, 15).Value) Then GLAdd d, "Rent", m, CDbl(w.Cells(r, 15).Value), 0
                If IsNumeric(w.Cells(r, 16).Value) Then GLAdd d, "Gas, Oil, Parking, Toll Fees", m, CDbl(w.Cells(r, 16).Value), 0
                If IsNumeric(w.Cells(r, 17).Value) Then GLAdd d, "Communication ", m, CDbl(w.Cells(r, 17).Value), 0
                If IsNumeric(w.Cells(r, 18).Value) Then GLAdd d, "Professional Fees", m, CDbl(w.Cells(r, 18).Value), 0
                a = GLAcct(w.Cells(r, 19).Value)
                If a <> "" Then
                    If IsNumeric(w.Cells(r, 20).Value) Then GLAdd d, a, m, CDbl(w.Cells(r, 20).Value), 0
                    If IsNumeric(w.Cells(r, 21).Value) Then GLAdd d, a, m, 0, CDbl(w.Cells(r, 21).Value)
                ElseIf Len(Trim$(CStr(w.Cells(r, 19).Value))) > 0 Then
                    GLAudit aw, ar, "CDJ", CStr(w.Cells(r, 19).Value), 0
                End If
            End If
        End If
    Next r
End Sub

Private Sub GLScanCRJ(ByVal d As Object)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CRJ")
    Dim ends As New Collection, r As Long, m As Long, st As Long, en As Long, rr As Long
    For r = 10 To w.Cells(w.Rows.count, 4).End(xlUp).Row
        If GLN(w.Cells(r, 6).Value) = "total" And GLMonth(w.Cells(r, 4).Value) > 0 Then ends.Add r
    Next r
    st = 10
    For m = 1 To 12
        If m <= ends.count Then en = ends(m) - 1 Else en = w.Cells(w.Rows.count, 6).End(xlUp).Row
        For rr = st To en
            If GLN(w.Cells(rr, 6).Value) <> "total" Then
                If IsNumeric(w.Cells(rr, 8).Value) Then GLAdd d, "Cash in Bank", m, CDbl(w.Cells(rr, 8).Value), 0
                If IsNumeric(w.Cells(rr, 9).Value) Then GLAdd d, "Excess CWT Over IT", m, CDbl(w.Cells(rr, 9).Value), 0
                If IsNumeric(w.Cells(rr, 10).Value) Then GLAdd d, "VAT Payable", m, 0, CDbl(w.Cells(rr, 10).Value)
                If IsNumeric(w.Cells(rr, 11).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 11).Value)
                If IsNumeric(w.Cells(rr, 12).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 12).Value)
            End If
        Next rr
        If m <= ends.count Then st = ends(m) + 1 Else Exit For
    Next m
End Sub

Private Sub GLScanGJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("GJ")
    Dim r As Long, m As Long, a As String, s As String, dr As Double, cr As Double
    For r = 10 To w.Cells(w.Rows.count, 5).End(xlUp).Row
        If GLMonth(w.Cells(r, 2).Value) > 0 Then m = GLMonth(w.Cells(r, 2).Value)
        s = Trim$(CStr(w.Cells(r, 5).Value))
        If m > 0 And s <> "" And GLN(s) <> "total" Then
            a = GLAcct(s)
            dr = 0: cr = 0
            If IsNumeric(w.Cells(r, 6).Value) Then dr = CDbl(w.Cells(r, 6).Value)
            If IsNumeric(w.Cells(r, 7).Value) Then cr = CDbl(w.Cells(r, 7).Value)
            If a <> "" Then
                GLAdd d, a, m, dr, cr
            ElseIf dr <> 0 Or cr <> 0 Then
                GLAudit aw, ar, "GJ", s, dr + cr
            End If
        End If
    Next r
End Sub

Private Function GLAccounts() As Variant
    GLAccounts = Array( _
        "Cash in Bank", "Petty Cash Fund", "Account Receivables", "Advances to Employees", _
        "Input VAT", "Excess CWT Over IT", "Leasehold Improvements", "Dental Equipment", _
        "Due To Clinicians", "Due to Clinicians- Visiting", "VAT Payable", "EWT Payable", _
        "Government Contributions (EE)", "Government Loans (EE)", "Income Tax Payable", _
        "Paid-Up Capital", "Retained Earnings (Deficit)", "Sales", "Rent", _
        "Clinician's Fee-Corporators", "Clinician's Fee-Visiting", "Clinic Material and Supplies", _
        "Common Use Service Area", "Light and Water Expense", _
        "Dep'n Expense-Leasehold Improvements", "Dep'n Expense-Dental Equipment", _
        "Salaries and Wages", "De Minimis", "Government Contributions ER Share", _
        "Professional Fees", "Repair and Maintenance", "Pantry Supplies", "Office Supplies", _
        "Meals, Gifts & Entertainment Expense", "Communication ", "Printing and Duplication", _
        "Taxes and Licences", "Delivery", "Software Subscription", "Miscellaneous", _
        "Transportation and Travel", "Advertisement", "Marketing Expense", _
        "Gas, Oil, Parking, Toll Fees", "Other Penalties and Charges", "Insurance")
End Function

' ============================================================
' REFRESH GL - public entry point
'   reportYear: 0 = auto (detect from source sheets, else system year)
'               any other value = force that year
' ============================================================
Private Sub Legacy_RefreshGL(Optional ByVal reportYear As Long = 0)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, gl As Worksheet
    Dim ar As Long, r As Long, i As Long, m As Long
    Dim bal As Double, x As Variant, a As Variant
    Dim yr As Long

    On Error GoTo EH

    ' ---- Resolve the reporting year ----
    If reportYear <= 0 Then
        yr = DetectGLYear()
    Else
        yr = reportYear
    End If
    If yr < 1900 Or yr > 9999 Then yr = year(Date)

    ' ---- GL_AUDIT sheet ----
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo EH
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        On Error Resume Next
        aw.name = "GL_AUDIT"
        On Error GoTo EH
    End If

    aw.Cells.Clear
    aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    ar = 2

    ' ---- Scan sources ----
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar

    ' ---- GL sheet must exist ----
    On Error Resume Next
    Set gl = ThisWorkbook.Sheets("GL")
    On Error GoTo EH
    If gl Is Nothing Then
        MsgBox "Sheet named 'GL' not found in this workbook." & vbCrLf & _
               "Existing sheets:" & vbCrLf & _
               Join(SheetNamesList(), vbCrLf), vbCritical, "RefreshGL"
        Exit Sub
    End If

    gl.Range("B12:I1000").ClearContents
    gl.Range("B11:I11").Value = Array("Date", "Day", "Reference", "Explanation", _
                                      "Account Title", "Debit", "Credit", "Ending Balance")

    ' ---- Write every account x month ----
    a = GLAccounts()
    r = 13
    For i = LBound(a) To UBound(a)
        bal = 0
        For m = 1 To 12
            gl.Cells(r, 2).Value = MonthName(m, True)
            gl.Cells(r, 3).Value = Day(DateSerial(yr, m + 1, 0))
            gl.Cells(r, 5).Value = "Ending Balance"
            gl.Cells(r, 6).Value = a(i)
            x = Array(0#, 0#)
            If d.Exists(CStr(a(i)) & "|" & m) Then x = d(CStr(a(i)) & "|" & m)
            gl.Cells(r, 7).Value = x(0)
            gl.Cells(r, 8).Value = x(1)
            bal = bal + x(0) - x(1)
            gl.Cells(r, 9).Value = bal
            r = r + 1
        Next m
    Next i

    gl.Range("G13:I" & r - 1).NumberFormat = "#,##0.00;[Red](#,##0.00);-"
    aw.Columns("A:C").AutoFit
    MsgBox "GL refreshed for " & yr & ": 46 accounts x 12 months.", vbInformation
    Exit Sub

EH:
    MsgBox "RefreshGL error " & Err.Number & ": " & Err.description, vbCritical, "RefreshGL Error"
End Sub

Private Function SheetNamesList() As String()
    Dim arr() As String, i As Long
    ReDim arr(1 To ThisWorkbook.Sheets.count)
    For i = 1 To ThisWorkbook.Sheets.count
        arr(i) = "  - " & ThisWorkbook.Sheets(i).name
    Next i
    SheetNamesList = arr
End Function

' ============================================================
' Detect the GL reporting year
'   1. Look for a year in CDJ/CRJ/GJ dates (max date wins)
'   2. Fall back to the current system year
' ============================================================
Private Function DetectGLYear() As Long
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim v As Variant
    Dim maxDate As Date
    Dim found As Boolean
    Dim sheetNames As Variant
    Dim sn As Variant

    sheetNames = Array("CDJ", "CRJ", "GJ")
    For Each sn In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sn))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Select Case CStr(sn)
                Case "CDJ": lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
                Case "CRJ": lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row
                Case "GJ":  lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
            End Select
            For i = 8 To lastRow
                Select Case CStr(sn)
                    Case "CDJ": v = ws.Cells(i, 3).Value
                    Case "CRJ": v = ws.Cells(i, 4).Value
                    Case "GJ":  v = ws.Cells(i, 2).Value
                End Select
                If IsDate(v) Then
                    If Not found Or CDate(v) > maxDate Then
                        maxDate = CDate(v)
                        found = True
                    End If
                End If
            Next i
        End If
        Set ws = Nothing
    Next sn

    If found Then
        DetectGLYear = year(maxDate)
    Else
        DetectGLYear = year(Date)
    End If
End Function

' ============================================================
' GL HELPERS - public API
' ============================================================
Private Function Legacy_CalculateMonthlyNet(ByVal accountCode As String, ByVal month As Integer, _
                                    ByVal year As Integer, Optional ByVal sheetName As String = "") As Double
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, ar As Long, x As Variant
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo 0
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add
        aw.name = "GL_AUDIT"
        aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    End If
    ar = aw.Cells(aw.Rows.count, 1).End(xlUp).Row + 1
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar
    x = Array(0#, 0#)
    If d.Exists(GLAcct(accountCode) & "|" & month) Then x = d(GLAcct(accountCode) & "|" & month)
    Legacy_CalculateMonthlyNet = x(0) - x(1)
End Function

Private Function Legacy_CalculateEndingBalanceFull(ByVal accountCode As String, ByVal year As Integer, _
                                            ByVal startBalance As Double, Optional ByVal targetSheet As String = "") As Variant
    Dim z(1 To 12) As Double, m As Long
    z(1) = startBalance + Legacy_CalculateMonthlyNet(accountCode, 1, year)
    For m = 2 To 12
        z(m) = z(m - 1) + Legacy_CalculateMonthlyNet(accountCode, m, year)
    Next m
    Legacy_CalculateEndingBalanceFull = z
End Function

Private Function Legacy_ValidateGLConsistency(Optional ByVal targetSheet As String = "GL") As Boolean
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets(targetSheet)
    Dim r As Long, td As Double, tc As Double, last As Long
    last = w.Cells(w.Rows.count, 6).End(xlUp).Row
    For r = 13 To last
        If IsNumeric(w.Cells(r, 7).Value) Then td = td + w.Cells(r, 7).Value
        If IsNumeric(w.Cells(r, 8).Value) Then tc = tc + w.Cells(r, 8).Value
    Next r
    Legacy_ValidateGLConsistency = Abs(td - tc) <= GL_AUTO_TOL
End Function

Private Sub Legacy_RefreshAllGL(ByVal targetSheet As String, Optional ByVal startBalances As Variant)
    RefreshGL
End Sub

Public Sub InitializeCOALookup()
End Sub

Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    GetAccountFromCOA = GLAcct(coaCode)
    If GetAccountFromCOA = "" Then GetAccountFromCOA = coaCode
End Function

' Supplier name for a TIN (punctuation-insensitive); returns the TIN itself
' when no supplier matches.
Public Function GetAccountNameFromTIN(ByVal tin As String) As String
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("SUPPLIERS DATA")
    Dim r As Long, last As Long, key As String
    key = NormTIN(tin)
    If key <> "" Then
        last = w.Cells(w.Rows.count, 3).End(xlUp).Row
        For r = 2 To last
            If NormTIN(w.Cells(r, 3).Value) = key Then
                GetAccountNameFromTIN = CStr(w.Cells(r, 4).Value)
                Exit Function
            End If
        Next r
    End If
    GetAccountNameFromTIN = tin
End Function

Public Function GetFinancialMonth(ByVal dt As Date) As Integer
    GetFinancialMonth = month(dt)
End Function

Public Function GetFinancialYear(ByVal dt As Date) As Integer
    GetFinancialYear = year(dt)
End Function

' ============================================================
' v8.3.3 GATE WRAPPERS
' ============================================================
' modGLGate.RunV8Gate calls these two names.  They did not exist anywhere in
' the project, so the VBA project did not compile.  They now delegate to the
' engine's own validators, which is the smallest change that makes the gate
' callable without altering the posting path.
' ============================================================
Public Function ValidateGLConsistency(Optional ByVal targetYear As Long = 0) As Boolean
    ValidateGLConsistency = Legacy_ValidateGLConsistency(GL_SHEET)
End Function

Public Function BuildV8CalcSheet(Optional ByVal targetYear As Long = 0) As Boolean
    Dim ws As Worksheet, last As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ' 552 = 46 accounts x 12 months.
    BuildV8CalcSheet = (last - 1 >= 552)
End Function
```

### 6.9 `modPJSync` - Legacy PJ Non-Vat2 helper module.

- Source: `build/v8_5_vba/modPJSync.bas`
- Type: standard | Host: (none) | CodeName: `modPJSync`
- Lines: 527 | Nonblank: 257 | Procedures: 8
- SHA-256: `3c2b4dccc075eb7ad6e852e432f2bdc20bd497254b1209a62959a84e925c3810`

**Analysis notes**

- Legacy helper module for `PJ Non-Vat2` (Sheet9). Provides its own supplier lookups, totals, sequence and direct CDJ writer.
- `SyncToCDJ_NonVat` finds an existing CDJ row by canonical/PNV/legacy reference, otherwise uses `modEngine.CDJReserveRow` to place the row in the correct monthly partition, then writes cash, EWT, reference, expense allocation and note.
- `FindExpenseCol` matches COA words against CDJ header rows 13/14 and falls back to the Sundry column.
- Apart from Sheet9, no current module calls these helpers; Sheet11/12 use modEngine instead.

```vba
Attribute VB_Name = "modPJSync"

Option Explicit



' ============================================================

' modPJSync � Helpers called by PJ Non-VAT2 Worksheet_Change

' Handles the NEW layout (with UOM at column H)

'

'  D=TIN  E=Name  F=Address  G=Description  H=UOM  I=Amount

'  J=Debit  K=DebitMirror  L=EWT  M=VAT  N=REF  O=COA  P=FullID

' ============================================================

Private Const ROW_START   As Long = 3

Private Const C_TIN       As Long = 4

Private Const C_NAME      As Long = 5

Private Const C_ADDR      As Long = 6

Private Const C_DESC      As Long = 7

Private Const C_UOM       As Long = 8

Private Const C_AMT       As Long = 9

Private Const C_DEBIT     As Long = 10

Private Const C_MIRROR    As Long = 11

Private Const C_EWT       As Long = 12

Private Const C_VAT       As Long = 13

Private Const C_REF       As Long = 14

Private Const C_COA       As Long = 15

Private Const C_FULLID    As Long = 16



' CDJ columns

Private Const CDJ_DATE    As Long = 3

Private Const CDJ_PART    As Long = 5

Private Const CDJ_CASH    As Long = 6

Private Const CDJ_VAT     As Long = 7

Private Const CDJ_WHTAX   As Long = 8

Private Const CDJ_DEBIT   As Long = 20

Private Const CDJ_REF      As Long = 27   ' column AA - canonical PJ reference
Private Const CDJ_REF_PNV  As Long = 28   ' column AB - PNV reference
Private Const CDJ_REF_LEG  As Long = 25   ' column Y  - legacy reference
Private Const CDJ_NOTE     As Long = 26   ' column Z  - ENTRY LOGS note

Private Const CDJ_START   As Long = 15

Private Const CDJ_EXP_S   As Long = 14

Private Const CDJ_EXP_E   As Long = 19

Private Const CDJ_SEQ     As Long = 27   ' column AA � sequence stored here



' ============================================================

' SHEET REFERENCE

' ============================================================

Private Function PJSheet() As Worksheet

    On Error Resume Next

    Set PJSheet = ThisWorkbook.Sheets("PJ Non-Vat2")

    On Error GoTo 0

End Function



' ============================================================

' SEQUENCE

' ============================================================

Public Function GetNextSequence() As Long

    Dim cdj As Worksheet, v As Variant

    On Error Resume Next

    Set cdj = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0

    If cdj Is Nothing Then GetNextSequence = 1: Exit Function



    v = cdj.Cells(13, CDJ_SEQ).Value

    If IsNumeric(v) Then

        GetNextSequence = CLng(v) + 1

    Else

        GetNextSequence = 1

    End If

    cdj.Cells(13, CDJ_SEQ).Value = GetNextSequence

End Function



' ============================================================

' SUPPLIER LOOKUPS

' ============================================================

Public Sub LookupByTIN_NonVat(ByVal rowNum As Long, ByVal tinVal As String)

    Dim ws As Worksheet, r As Long, lastRow As Long

    Dim key As String, hit As String

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ' A BIR TIN is 15 digits; the live data also contains '009-149--884',
    ' ' 240-940-697' and '636-985956'.  Comparing the normalised digit string
    ' on both sides makes those match the SUPPLIERS DATA row they belong to.
    key = modEngine.NormTIN(tinVal)

    If key = "" Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row

    For r = 2 To lastRow

        If modEngine.NormTIN(ws.Cells(r, 3).Value) = key Then

            ' Fail closed on a TIN that resolves to more than one supplier:
            ' a silent wrong-vendor match is worse than no match at all.
            If Len(hit) > 0 Then Exit Sub

            hit = Trim$(CStr(ws.Cells(r, 4).Value))

            PJSheet.Cells(rowNum, C_NAME).Value = ws.Cells(r, 4).Value

            PJSheet.Cells(rowNum, C_ADDR).Value = ws.Cells(r, 5).Value

        End If

    Next r

End Sub



Public Sub LookupByName_NonVat(ByVal rowNum As Long, ByVal nameVal As String)

    Dim ws As Worksheet, r As Long, lastRow As Long

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row

    For r = 2 To lastRow

        If StrComp(Trim$(CStr(ws.Cells(r, 4).Value)), Trim$(nameVal), vbTextCompare) = 0 Then

            PJSheet.Cells(rowNum, C_TIN).Value = ws.Cells(r, 3).Value

            PJSheet.Cells(rowNum, C_ADDR).Value = ws.Cells(r, 5).Value

            Exit Sub

        End If

    Next r

End Sub



Public Sub ClearSupplierRow(ByVal rowNum As Long)

    Dim sh As Worksheet: Set sh = PJSheet

    If sh Is Nothing Then Exit Sub

    sh.Cells(rowNum, C_NAME).ClearContents

    sh.Cells(rowNum, C_ADDR).ClearContents

End Sub



' ============================================================

' TOTALS

' ============================================================

Public Sub CalculateTotals(ByVal rowNum As Long)

    Dim sh As Worksheet, i As Long

    Dim sumAmt As Double, sumDeb As Double, sumMir As Double

    Set sh = PJSheet

    If sh Is Nothing Then Exit Sub



    For i = ROW_START To rowNum - 1

        If StrComp(Trim$(CStr(sh.Cells(i, C_TIN).Value)), "TOTAL", vbTextCompare) <> 0 _

           And StrComp(Trim$(CStr(sh.Cells(i, C_NAME).Value)), "TOTAL", vbTextCompare) <> 0 Then

            If IsNumeric(sh.Cells(i, C_AMT).Value) Then sumAmt = sumAmt + CDbl(sh.Cells(i, C_AMT).Value)

            If IsNumeric(sh.Cells(i, C_DEBIT).Value) Then sumDeb = sumDeb + CDbl(sh.Cells(i, C_DEBIT).Value)

            If IsNumeric(sh.Cells(i, C_MIRROR).Value) Then sumMir = sumMir + CDbl(sh.Cells(i, C_MIRROR).Value)

        End If

    Next i



    sh.Cells(rowNum, C_AMT).Value = sumAmt

    sh.Cells(rowNum, C_DEBIT).Value = sumDeb

    sh.Cells(rowNum, C_MIRROR).Value = sumMir

    sh.Cells(rowNum, C_EWT).Value = 0

    sh.Cells(rowNum, C_VAT).Value = 0

End Sub



' ============================================================

' SYNC PJ NON-VAT ? CDJ

' ============================================================

Public Sub SyncToCDJ_NonVat(ByVal rowNum As Long)

    Dim sh As Worksheet, cdj As Worksheet

    Dim ref As String, part As String, coa As String

    Dim dt As Variant, amt As Double, ewt As Double

    Dim targetRow As Long, r As Long, lastRow As Long, expCol As Long

    Dim mth As Integer, c As Long



    On Error GoTo EH



    Set sh = PJSheet

    If sh Is Nothing Then Exit Sub

    Set cdj = ThisWorkbook.Sheets("CDJ")



    ref = Trim$(CStr(sh.Cells(rowNum, C_REF).Value))

    If ref = "" Or StrComp(ref, "TOTAL", vbTextCompare) = 0 Then Exit Sub



    part = Trim$(CStr(sh.Cells(rowNum, C_NAME).Value))

    coa = Trim$(CStr(sh.Cells(rowNum, C_COA).Value))

    dt = sh.Cells(rowNum, 2).Value

    amt = 0

    If IsNumeric(sh.Cells(rowNum, C_AMT).Value) Then amt = CDbl(sh.Cells(rowNum, C_AMT).Value)

    ewt = 0

    If IsNumeric(sh.Cells(rowNum, C_EWT).Value) Then ewt = CDbl(sh.Cells(rowNum, C_EWT).Value)



    If amt = 0 Then Exit Sub



    ' --- find an existing CDJ row carrying this ref --------------------
    ' The canonical reference column is AA(27).  The engine's old constant
    ' pointed at Y(25), which holds only 29 stray cells, so 566 of 595 refs
    ' were invisible here and updates missed their rows.
    targetRow = 0

    lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF).End(xlUp).Row

    If cdj.Cells(cdj.Rows.count, CDJ_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF_PNV).End(xlUp).Row

    If cdj.Cells(cdj.Rows.count, CDJ_REF_LEG).End(xlUp).Row > lastRow Then _
        lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF_LEG).End(xlUp).Row

    If lastRow < CDJ_START Then lastRow = CDJ_START

    For r = CDJ_START To lastRow

        If Trim$(CStr(cdj.Cells(r, CDJ_REF).Value)) = ref Or _
           Trim$(CStr(cdj.Cells(r, CDJ_REF_PNV).Value)) = ref Or _
           Trim$(CStr(cdj.Cells(r, CDJ_REF_LEG).Value)) = ref Then

            targetRow = r

            Exit For

        End If

    Next r



    ' --- else place inside the month's own partition --------------------
    ' Scanning for the first empty DATE cell put every insert past the last
    ' partition, where no monthly SUM covers it.  CDJReserveRow inserts above
    ' the month's TOTAL row when the block is full, so Excel widens the SUM.
    If targetRow = 0 Then

        mth = CDJMarkerMonth(CStr(dt))

        If mth = 0 Then mth = modEngine.PJMonthOf(PJSheet, rowNum)

        targetRow = modEngine.CDJReserveRow(cdj, mth)

        ' Fail closed: with no legal placement the row must not be written,
        ' otherwise it lands outside every monthly total again.
        If targetRow = 0 Then Exit Sub

    End If



    ' --- find expense column by matching COA header ---

    expCol = FindExpenseCol(cdj, coa)



    ' --- write -----------------------------------------------------------
    ' CDJ row-balance column X = V+W where V=F and W=SUM(G:T).  Requiring
    ' X = 0 gives  -(cash) + vat - ewt + allocation = 0, and because
    ' gross = net + vat this holds exactly when
    '     cash = -(gross - ewt),  vat = 0,  ewt = -ewt,  allocation = net.
    ' Live rows 17 and 620 both reconcile to X = 0 under this convention.
    cdj.Cells(targetRow, CDJ_DATE).Value = dt

    cdj.Cells(targetRow, CDJ_PART).Value = part

    cdj.Cells(targetRow, CDJ_CASH).Value = (amt - ewt) * -1

    cdj.Cells(targetRow, CDJ_VAT).Value = 0

    cdj.Cells(targetRow, CDJ_WHTAX).Value = ewt * -1

    cdj.Cells(targetRow, CDJ_REF_PNV).Value = ref

    ' Clear every allocation column including T (20).  Stopping at S (19)
    ' left a stale Sundry Debit behind whenever a row was reclassified.
    For c = CDJ_EXP_S To CDJ_DEBIT

        cdj.Cells(targetRow, c).Value = ""

    Next c

    If expCol >= CDJ_EXP_S And expCol < CDJ_EXP_E Then

        ' Fixed expense cols 14-18: net amount only.  Column T stays empty
        ' so the amount is not counted twice.
        cdj.Cells(targetRow, expCol).Value = amt - ewt

        cdj.Cells(targetRow, expCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    ElseIf expCol = CDJ_EXP_E Then

        ' Sundry: COA label in S (19), net amount in T (20).  Column R (18,
        ' Professional Fees) stays empty so the amount is never classified as
        ' Professional Fees as well as Sundry.
        cdj.Cells(targetRow, CDJ_EXP_E).Value = coa

        cdj.Cells(targetRow, CDJ_DEBIT).Value = amt - ewt

        cdj.Cells(targetRow, CDJ_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    End If

    cdj.Cells(targetRow, CDJ_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_NOTE).Value = "PJ Non-VAT row " & rowNum & " | " & Now



    Exit Sub

EH:

    MsgBox "SyncToCDJ_NonVat error " & Err.Number & ": " & Err.description, _
           vbExclamation, "PJ Sync"

End Sub



' ============================================================

' Find expense column matching COA description

' ============================================================

Private Function FindExpenseCol(ByVal cdj As Worksheet, ByVal coa As String) As Long

    Dim col As Long, header As String, cLow As String, kw As Variant

    cLow = LCase$(Trim$(coa))

    If cLow = "" Then Exit Function



    For col = CDJ_EXP_S To CDJ_EXP_E

        header = LCase$(Trim$(CStr(cdj.Cells(13, col).Value) & " " & _

                              CStr(cdj.Cells(14, col).Value)))

        If header = "" Then GoTo NextCol



        For Each kw In Split(cLow, " ")

            If Len(kw) >= 4 And InStr(header, CStr(kw)) > 0 Then

                FindExpenseCol = col

                Exit Function

            End If

        Next kw



        If Len(header) >= 4 And InStr(cLow, header) > 0 Then

            FindExpenseCol = col

            Exit Function

        End If

NextCol:

    Next col

    ' Nothing matched.  Route to the Sundry column so the amount still gets an
    ' allocation and the row balances, exactly as modEngine.FindExpenseColumn
    ' does.  Returning 0 used to leave the row with cash/vat/ewt but no
    ' expense column, so it never reconciled to zero.
    FindExpenseCol = CDJ_EXP_E

End Function
```

### 6.10 `modGLGate` - Pre-scan workbook structure gate.

- Source: `build/v8_5_vba/modGLGate.bas`
- Type: standard | Host: (none) | CodeName: `modGLGate`
- Lines: 49 | Nonblank: 43 | Procedures: 3
- SHA-256: `8bc6e2104980b241fd63a3b113bf80974328ebc5e55b41045ec04a5d8db87e7d`

**Analysis notes**

- Preflight gate used by `RunV8Gate`. Verifies required sheets (`CDJ`, `CRJ`, `GL`, `COA_MAP`), header cells and 46 nonblank GL block titles, then calls `ScanGL 2026`.
- It does not itself test account-title uniqueness, mapping coverage or reconciliation; `ScanGL` performs those checks before write-back.
- The gate is independent of `RefreshGL`; both route to the same scan engine.

```vba
Attribute VB_Name = "modGLGate"
Option Explicit

' v8.2 fail-closed gate. Exact workbook structure is validated before calculation.

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(name): On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function ValidateV8Structure() As Boolean
    Dim names As Variant, n As Variant, ws As Worksheet
    names = Array("CDJ", "CRJ", "GL", "COA_MAP")
    For Each n In names
        If Not SheetExists(CStr(n)) Then
            ValidateV8Structure = False
            Exit Function
        End If
    Next n

    Set ws = ThisWorkbook.Worksheets("CDJ")
    If V8_Norm(ws.Cells(13, 3).Value2) <> "DATE" Or V8_Norm(ws.Cells(14, 20).Value2) <> "DEBIT" Or V8_Norm(ws.Cells(14, 21).Value2) <> "CREDIT" Then Exit Function
    Set ws = ThisWorkbook.Worksheets("CRJ")
    If V8_Norm(ws.Cells(8, 4).Value2) <> "DATE" Or V8_Norm(ws.Cells(8, 8).Value2) <> "DEBIT" Then Exit Function
    ' GJ structure is deliberately not inspected; it is outside GL scan scope.
    Set ws = ThisWorkbook.Worksheets("GL")
    If V8_Norm(ws.Cells(11, 6).Value2) <> "ACCOUNT TITLE" Then Exit Function
    Dim b As Long, r As Long, n As Long
    For b = 0 To 45
        r = 13 + b * 13
        If Len(Trim$(CStr(ws.Cells(r, 6).Value2))) = 0 Then Exit Function
        n = n + 1
    Next b
    If n <> 46 Then Exit Function
    ValidateV8Structure = True
End Function

Public Sub RunV8Gate()
    If Not ValidateV8Structure() Then
        If Application.UserControl And Application.DisplayAlerts Then _
            MsgBox "GL gate HOLD: CDJ/CRJ/GL/COA_MAP structure does not match the expected layout.", vbExclamation
        Exit Sub
    End If

    ' Use the current fail-closed CDJ + CRJ scan. Do not invoke the legacy
    ' GL_V8_CALC builder or inspect the excluded GJ worksheet.
    ScanGL 2026
End Sub
```

### 6.11 `modGLRefresh` - Macro-dialog wrappers for GL refresh.

- Source: `build/v8_5_vba/modGLRefresh.bas`
- Type: standard | Host: (none) | CodeName: `modGLRefresh`
- Lines: 27 | Nonblank: 22 | Procedures: 4
- SHA-256: `b702162b7203da909c48ba6bc18983f45dfb9ecd7a18fc486748a5708db75f93`

**Analysis notes**

- Macro-dialog compatibility layer. All zero-argument wrappers (`RefreshGL`, `RefreshAllGL`, `RefreshGLIntoSheet`) call `ScanGL 2026`.
- `RefreshGLForYear(yearNumber)` forwards the year, but the downstream scan currently ignores the year parameter (see ScanCDJ/ScanCRJ).
- The old matrix-writer is retired from this path.

```vba
Attribute VB_Name = "modGLRefresh"
Option Explicit

' v8.3.3 compatibility wrapper.
' The old RefreshGL matrix writer is retired from the execution path.
' RefreshGL delegates to the row-scanning engine so the visible GL,
' GL_V8_CALC, and GL_AUDIT all come from the same scanned postings.
'
' IMPORTANT: Macros intended for the Excel Macro Dialog (Alt+F8) MUST NOT
' take any parameters (even Optional ones). Subroutines below are zero-argument
' wrappers to ensure complete visibility in Alt+F8.

Public Sub RefreshGL()
    ScanGL 2026
End Sub

Public Sub RefreshAllGL()
    ScanGL 2026
End Sub

Public Sub RefreshGLIntoSheet()
    ScanGL 2026
End Sub

Public Sub RefreshGLForYear(ByVal yearNumber As Long)
    ScanGL yearNumber
End Sub
```

### 6.12 `modGLScanEngine` - Active CDJ/CRJ GL scan engine.

- Source: `build/v8_5_vba/modGLScanEngine.bas`
- Type: standard | Host: (none) | CodeName: `modGLScanEngine`
- Lines: 331 | Nonblank: 292 | Procedures: 9
- SHA-256: `68551ae2a78b6230a5b729f5b3a4ce631d2cd867daa5790e84200d7543f51497`

**Analysis notes**

- Active GL engine. `ScanGL` resets mapping caches, validates the 46 GL account titles, scans CDJ and CRJ into an account/month dictionary, applies a fail-closed gate, then writes the visible GL and GL_AUDIT.
- CDJ scanning: month comes from column C with carry-forward; cash from F; allocations G:R via `V8_CDJMap`; Sundry account from S; balancing Sundry amount is derived from cash and allocations instead of trusting T/U; T/U disagreements are recorded but do not by themselves fail the gate.
- CRJ scanning: month attribution is partition-based (column D month boundary with F/G blank), not based on the timestamp in column O. The first boundary must be January and each subsequent boundary must be the next month.
- Gate: `inv > 0` or `um > TOLERANCE` or `Abs(gd - gc) > TOLERANCE` -> write HOLD audit and abort before touching GL. Otherwise `WriteScannedGL` writes G/H/I and `WriteScanAudit` appends PASS.
- `WriteScannedGL` starts each 12-month block at balance 0; it does not carry a prior-year opening balance into January.
- The `yr` parameter on `ScanCDJ` and `ScanCRJ` is declared but unused; the scan buckets all matching source rows by month. Year scoping relies on the workbook containing a single reporting year and, for CRJ, on partition order.

```vba
Attribute VB_Name = "modGLScanEngine"
Option Explicit

' v8.3.3 SCANNING ENGINE (Hardened)
' Reads the posted journals row-by-row, resolves the actual COA, aggregates
' by physical GL account + month, and writes the visible GL directly.
'
' Design rule:
'   SOURCE ROW -> COA RESOLUTION -> MONTH BUCKET -> GL BLOCK -> G/H/I
'
' CDJ sundry rows are balanced from the row's actual cash leg and mapped
' allocation legs. This prevents a stale/incorrect Sundry Debit value from
' becoming the GL amount. The Sundry Account name remains the COA identity.

Private Const FIRST_GL_ROW As Long = 13
Private Const GL_STRIDE As Long = 13
Private Const GL_BLOCKS As Long = 46
Private Const MONTHS As Long = 12

Public Sub ScanGL(Optional ByVal yearNumber As Long = 2026)
    Dim d As Object, gd As Double, gc As Double, um As Double, inv As Long
    Dim detail As String, ws As Worksheet

    ' Pick up any operator edit to the COA_MAP alias table before mapping.
    V8_MapReset

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    ' Fail closed if the visible GL does not define 46 distinct account titles.
    ' Duplicate titles would receive the same normalized postings in multiple blocks.
    If Not ValidateGLAccountTitles(detail) Then inv = inv + 1

    ScanCDJ d, yearNumber, gd, gc, um, inv, detail
    ScanCRJ d, yearNumber, gd, gc, um, inv, detail
    ' GJ is out of scope: do not read or post from the GJ worksheet.

    ' Fail-closed validation gate:
    ' If journals are out of balance, have unmapped accounts, or have invalid data,
    ' record the diagnostic audit row and abort BEFORE modifying the visible GL.
    If inv > 0 Or um > TOLERANCE Or Abs(gd - gc) > TOLERANCE Then
        WriteScanAudit gd, gc, um, inv, detail
    If Application.UserControl And Application.DisplayAlerts Then
            MsgBox "v8.3.3 SCAN HOLD: Reconciliation or mapping gate failed. GL not overwritten." & vbCrLf & _
                   "Debit: " & Format$(gd, "0.00") & vbCrLf & _
                   "Credit: " & Format$(gc, "0.00") & vbCrLf & _
                   "Difference: " & Format$(gd - gc, "0.00") & vbCrLf & _
                   "Unmapped: " & Format$(um, "0.00") & vbCrLf & _
                   "Invalid: " & inv & vbCrLf & _
                   "Review GL_AUDIT for details.", vbExclamation
        End If
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    WriteScannedGL ws, d
    ' WriteScannedCalc excluded per directive: do not introduce another calculation sheet
    WriteScanAudit gd, gc, um, inv, detail

    If Application.UserControl And Application.DisplayAlerts Then
        MsgBox "v8.3.3 SCAN COMPLETE" & vbCrLf & _
               "GL rows scanned: 46 accounts x 12 months = 552" & vbCrLf & _
               "Debit: " & Format$(gd, "0.00") & vbCrLf & _
               "Credit: " & Format$(gc, "0.00") & vbCrLf & _
               "Difference: " & Format$(gd - gc, "0.00") & vbCrLf & _
               "Unmapped: " & Format$(um, "0.00") & vbCrLf & _
               "Invalid: " & inv, vbInformation
    End If
End Sub

Private Sub AddPost(ByVal d As Object, ByVal acct As String, ByVal m As Long, ByVal db As Double, ByVal cr As Double)
    Dim k As String, a As Variant
    If Len(Trim$(acct)) = 0 Or m < 1 Or m > 12 Then Exit Sub
    k = V8_Norm(acct) & "|" & CStr(m)
    If Not d.Exists(k) Then d.Add k, Array(0#, 0#)
    a = d(k)
    a(0) = CDbl(a(0)) + db
    a(1) = CDbl(a(1)) + cr
    d(k) = a
End Sub

Private Sub ScanCDJ(ByVal d As Object, ByVal yr As Long, ByRef gd As Double, ByRef gc As Double, ByRef um As Double, ByRef inv As Long, ByRef detail As String)
    Dim ws As Worksheet, r As Long, c As Long, m As Long, cur As Long
    Dim ok As Boolean, v As Double, a As String, raw As String
    Dim cash As Double, otherDb As Double, otherCr As Double, sundryNet As Double
    Dim cashOK As Boolean, explicitDb As Double, explicitCr As Double
    Dim title As String

    Set ws = ThisWorkbook.Worksheets("CDJ")
    cur = 0

    For r = 15 To ws.Cells(ws.Rows.count, 5).End(xlUp).Row
        m = V8_Month(ws.Cells(r, 3).Value2)
        If m > 0 Then cur = m
        If cur = 0 Then GoTo NextRow
        title = SafeText(ws.Cells(r, 5).Value2)
        If Len(title) = 0 Or InStr(1, V8_Norm(title), "TOTAL", vbTextCompare) > 0 Then GoTo NextRow

        ' Cash is column F: negative = credit, positive = debit.
        ok = True: cash = V8_Number(ws.Cells(r, 6).Value2, ok)
        If Not ok Then inv = inv + 1: GoTo NextRow

        If Abs(cash) > TOLERANCE Then
            If cash >= 0 Then
                AddPost d, "Cash in Bank", cur, cash, 0#
                gd = gd + cash
            Else
                AddPost d, "Cash in Bank", cur, 0#, -cash
                gc = gc - cash
            End If
        End If

        otherDb = 0#: otherCr = 0#
        For c = 7 To 18
            ok = True: v = V8_Number(ws.Cells(r, c).Value2, ok)
            If Not ok Then inv = inv + 1: GoTo NextRow
            If Abs(v) > TOLERANCE Then
                a = V8_CDJMap(c)
                If Len(a) = 0 Then
                    um = um + Abs(v)
                    detail = detail & "CDJ!" & ws.Cells(r, c).address(False, False) & " unmapped allocation" & vbCrLf
                ElseIf v > 0 Then
                    AddPost d, a, cur, v, 0#: gd = gd + v: otherDb = otherDb + v
                Else
                    AddPost d, a, cur, 0#, -v: gc = gc - v: otherCr = otherCr - v
                End If
            End If
        Next c

        ' Sundry Account is the actual COA label. Derive its balancing amount
        ' from the row rather than trusting a stale copied amount in T/U.
        raw = SafeText(ws.Cells(r, 19).Value2)
        If Len(raw) > 0 And V8_Norm(raw) <> "TOTAL" And V8_Norm(raw) <> "SUNDRY ACCOUNT" Then
            a = V8_ResolveAccount(raw)
            If Len(a) = 0 And GLTitleExists(raw) Then a = GLTitleOf(raw)
            If Len(a) = 0 Then
                um = um + Abs(cash) + otherDb + otherCr
                detail = detail & "CDJ!" & ws.Cells(r, 19).address(False, False) & " -> " & raw & " unmapped Sundry Account" & vbCrLf
            Else
                ' If cash is a credit, sundry must supply the balancing debit
                ' after the other debit/credit allocations are considered.
                sundryNet = (-cash) + otherCr - otherDb
                If sundryNet > TOLERANCE Then
                    AddPost d, a, cur, sundryNet, 0#: gd = gd + sundryNet
                ElseIf sundryNet < -TOLERANCE Then
                    AddPost d, a, cur, 0#, -sundryNet: gc = gc - sundryNet
                End If

                ' Record disagreement with the visible T/U amount, but do not
                ' allow it to corrupt the balanced scan.
                cashOK = True
                explicitDb = V8_Number(ws.Cells(r, 20).Value2, cashOK)
                If Not cashOK Then inv = inv + 1
                cashOK = True
                explicitCr = V8_Number(ws.Cells(r, 21).Value2, cashOK)
                If Not cashOK Then inv = inv + 1
                If Abs(explicitDb - MaxD(sundryNet, 0#)) > TOLERANCE Or _
                   Abs(explicitCr - MaxD(-sundryNet, 0#)) > TOLERANCE Then
                    detail = detail & "CDJ!" & ws.Cells(r, 19).address(False, False) & _
                             " Sundry amount corrected from T/U to balanced row amount: " & _
                             Format$(sundryNet, "0.00") & vbCrLf
                End If
            End If
        End If
NextRow:
    Next r
End Sub

Private Sub ScanCRJ(ByVal d As Object, ByVal yr As Long, ByRef gd As Double, ByRef gc As Double, ByRef um As Double, ByRef inv As Long, ByRef detail As String)
    ' CRJ month attribution is partition-based, not based on column O.
    ' Column O contains SJ import/log timestamps and is not a transaction date.
    ' A live monthly boundary row carries the month label in column D and is blank in
    ' columns F (customer) and G (invoice number). The legacy v8.3.3 summary block lower
    ' in the sheet has F = "TOTAL" with month labels in D but is NOT a live boundary;
    ' the blank F/G test excludes it. Contiguity is enforced below (first boundary must
    ' be January; each subsequent boundary must be exactly the next month).
    ' The yr parameter is retained for caller compatibility; partition order proves the year scope.
    Dim ws As Worksheet, r As Long, rr As Long, c As Long
    Dim m As Long, lastMonth As Long, startRow As Long, lastRow As Long
    Dim ok As Boolean, v As Double, a As String, invNo As String
    Dim partitionOK As Boolean

    Set ws = ThisWorkbook.Worksheets("CRJ")
    ' Scan range must include boundary rows themselves: they carry data in column D
    ' (month label) but are blank in F/G, so F/G alone would cut the range short
    ' at the last invoice and skip the final monthly boundary.
    lastRow = ws.Cells(ws.Rows.count, 6).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 7).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 4).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row

    startRow = 10
    lastMonth = 0

    For r = startRow To lastRow
        m = V8_Month(ws.Cells(r, 4).Value2)
        If m > 0 And Len(SafeText(ws.Cells(r, 6).Value2)) = 0 And _
           Len(SafeText(ws.Cells(r, 7).Value2)) = 0 Then
            partitionOK = (m = lastMonth + 1 And m >= 1 And m <= 12)

            If Not partitionOK Then
                detail = detail & "CRJ!D" & r & " invalid or out-of-order partition month; partition held" & vbCrLf
                For rr = startRow To r - 1
                    If Len(SafeText(ws.Cells(rr, 7).Value2)) > 0 Then inv = inv + 1
                Next rr
            Else
                For rr = startRow To r - 1
                    invNo = SafeText(ws.Cells(rr, 7).Value2)
                    If Len(invNo) > 0 Then
                        ok = True: v = V8_Number(ws.Cells(rr, 8).Value2, ok)
                        If Not ok Then
                            inv = inv + 1
                        ElseIf Abs(v) > TOLERANCE Then
                            AddPost d, "Cash in Bank", m, v, 0#
                            gd = gd + v
                        End If

                        For c = 9 To 12
                            ok = True: v = V8_Number(ws.Cells(rr, c).Value2, ok)
                            If Not ok Then
                                inv = inv + 1
                            ElseIf Abs(v) > TOLERANCE Then
                                a = V8_CRJMap(c)
                                If Len(a) = 0 Then
                                    um = um + Abs(v)
                                    detail = detail & "CRJ!" & ws.Cells(rr, c).address(False, False) & " unmapped credit" & vbCrLf
                                Else
                                    AddPost d, a, m, 0#, v
                                    gc = gc + v
                                End If
                            End If
                        Next c
                    End If
                Next rr
            End If

            lastMonth = m
            startRow = r + 1
        End If
    Next r

    ' Never guess a month for rows that appear after the last recognized monthly boundary.
    For rr = startRow To lastRow
        If Len(SafeText(ws.Cells(rr, 7).Value2)) > 0 Then
            inv = inv + 1
            detail = detail & "CRJ!G" & rr & " invoice has no enclosing monthly partition" & vbCrLf
        End If
    Next rr
End Sub

' GJ scan intentionally removed. This workbook's GL input scope is CDJ + CRJ only.

Private Function ValidateGLAccountTitles(ByRef detail As String) As Boolean
    Dim ws As Worksheet, seen As Object
    Dim b As Long, r As Long, title As String, key As String
    Dim valid As Boolean

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    valid = True

    For b = 0 To GL_BLOCKS - 1
        r = FIRST_GL_ROW + b * GL_STRIDE
        title = SafeText(ws.Cells(r, 6).Value2)
        key = V8_Norm(title)

        If Len(key) = 0 Then
            detail = detail & "GL!F" & r & " missing account title for block " & (b + 1) & vbCrLf
            valid = False
        ElseIf seen.Exists(key) Then
            detail = detail & "GL!F" & r & " duplicate account title '" & title & _
                     "' (first seen at GL!F" & seen(key) & "); write-back held" & vbCrLf
            valid = False
        Else
            seen.Add key, r
        End If
    Next b

    ValidateGLAccountTitles = valid
End Function

Private Sub WriteScannedGL(ByVal ws As Worksheet, ByVal d As Object)
    Dim b As Long, m As Long, r As Long, k As String, x As Variant
    Dim title As String, bal As Double

    Application.ScreenUpdating = False
    On Error GoTo Fail
    For b = 0 To GL_BLOCKS - 1
        title = SafeText(ws.Cells(FIRST_GL_ROW + b * GL_STRIDE, 6).Value2)
        If Len(title) = 0 Then Err.Raise vbObjectError + 833, "ScanGL", "Missing GL account title at block " & (b + 1)
        bal = 0#
        For m = 1 To MONTHS
            r = FIRST_GL_ROW + b * GL_STRIDE + m - 1
            k = V8_Norm(title) & "|" & CStr(m)
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            ws.Cells(r, 7).Value2 = CDbl(x(0))
            ws.Cells(r, 8).Value2 = CDbl(x(1))
            bal = bal + CDbl(x(0)) - CDbl(x(1))
            ws.Cells(r, 9).Value2 = bal
        Next m
    Next b
Fail:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.description
End Sub

' Legacy GL_V8_CALC writer removed: output is the existing GL sheet only.

Private Sub WriteScanAudit(ByVal gd As Double, ByVal gc As Double, ByVal um As Double, ByVal inv As Long, ByVal detail As String)
    Dim ws As Worksheet, r As Long, status As String
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = AUDIT_SHEET
    On Error GoTo 0
    status = IIf(Abs(gd - gc) <= TOLERANCE And um <= TOLERANCE And inv = 0, "PASS", "HOLD")
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = status
    ws.Cells(r, 3).Value = "v8.3.3 SCAN | Debit=" & Format$(gd, "0.00") & " | Credit=" & Format$(gc, "0.00") & " | Difference=" & Format$(gd - gc, "0.00") & " | Unmapped=" & Format$(um, "0.00") & " | Invalid=" & inv
    ws.Cells(r, 4).Value = Left$(detail, 32000)
End Sub

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    SafeText = Trim$(CStr(v))
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
```

### 6.13 `modGLWorkbookMap` - Account/column mapping and constants.

- Source: `build/v8_5_vba/modGLWorkbookMap.bas`
- Type: standard | Host: (none) | CodeName: `modGLWorkbookMap`
- Lines: 207 | Nonblank: 189 | Procedures: 12
- SHA-256: `bb4f49b84855fdaa96a9d1db87be80968eb54ee5e2f67753bd39fe69037ebaaa`

**Analysis notes**

- Holds the public sheet/tolerance constants (`GL`, `GL_AUDIT`, `COA_MAP`, `GL_V8_CALC`, `TOLERANCE = 0.01`) and all GL account resolution maps.
- `EnsureMaps` builds three dictionaries: normalized GL title, punctuation-insensitive GL title, and COA_MAP aliases. Aliases are accepted from COA_MAP rows 5+ when column D has an account and column F status contains `MAPPED`.
- Audit finding: `InStr(st, "MAPPED") > 0` also matches the substring inside `UNMAPPED`. If any status uses `UNMAPPED`, that alias would be accepted despite the fail-closed intent.
- `V8_ResolveAccount` applies COA_MAP aliases first, then hard-coded aliases (`MEDICAL EQUIPMENT`, `COST OF REVENUE`, `SUPPLIES`, `BANK CHARGE`, `CHARGES`, etc.), then exact/punctuation-insensitive GL title matches.
- `V8_CDJMap` hard-codes CDJ allocation columns 7-18; `V8_CRJMap` hard-codes CRJ credit columns 9-12.

```vba
Attribute VB_Name = "modGLWorkbookMap"
Option Explicit

' Account resolution for the Global-Smile 2026 workbook layout.
' Posting sources in this GL scan are CDJ and CRJ only. GJ is excluded.

Public Const AUDIT_SHEET As String = "GL_AUDIT"
Public Const CALC_SHEET As String = "GL_V8_CALC"
Public Const GL_SHEET As String = "GL"
Public Const COAMAP_SHEET As String = "COA_MAP"
Public Const TOLERANCE As Double = 0.01

' Module-level caches. The account resolver builds dictionaries once per scan
' and clears them through V8_MapReset.
Private mGLByNorm As Object      ' V8_Norm(title) -> title
Private mGLByKey As Object       ' V8_Key(title)  -> title
Private mAlias As Object         ' V8_Key(source label) -> GL title
Private mMapsBuilt As Boolean

' Drop every cached map so an edit to COA_MAP or GL takes effect on the next
' scan without reopening the workbook.
Public Sub V8_MapReset()
    Set mGLByNorm = Nothing
    Set mGLByKey = Nothing
    Set mAlias = Nothing
    mMapsBuilt = False
End Sub

' Punctuation- and spacing-insensitive key.  "Clinician's Fee-Corporators" and
' "Clinicians Fee- Corporators" both reduce to CLINICIANSFEECORPORATORS, which
' is how the nine form-only sundry labels find their real GL account.
Public Function V8_Key(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = UCase$(Replace(CStr(v), ChrW(160), " "))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            V8_Key = V8_Key & ch
        End If
    Next i
End Function

Private Sub EnsureMaps()
    Dim d As Object, ka As Object, al As Object
    Dim ws As Worksheet, r As Long, a As String, k As String, acct As String
    Dim st As String
    If mMapsBuilt Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ka = CreateObject("Scripting.Dictionary"): ka.CompareMode = vbTextCompare
    Set al = CreateObject("Scripting.Dictionary"): al.CompareMode = vbTextCompare

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    For r = 13 To ws.Cells(ws.Rows.count, 6).End(xlUp).Row
        a = Trim$(CStr(ws.Cells(r, 6).Value2))
        If Len(a) > 0 Then
            If Not d.Exists(V8_Norm(a)) Then d.Add V8_Norm(a), a
            If Not ka.Exists(V8_Key(a)) Then ka.Add V8_Key(a), a
        End If
    Next r

    ' Aliases come from the COA_MAP sheet, which is the operator-editable
    ' single source of truth.  A row is only honoured when it names an
    ' account AND its status says MAPPED, so the seven labels held for
    ' accounting decision (blank account, status NEEDS ACCOUNTING DECISION)
    ' are deliberately not aliased and the gate keeps holding on them.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COAMAP_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        r = 5
        For r = 5 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            acct = Trim$(CStr(ws.Cells(r, 4).Value2))
            st = UCase$(Trim$(CStr(ws.Cells(r, 6).Value2)))
            If Len(acct) > 0 And InStr(st, "MAPPED") > 0 Then
                k = V8_Key(ws.Cells(r, 1).Value2)
                If Len(k) > 0 And Not al.Exists(k) Then al.Add k, acct
            End If
            ' (increment handled by For loop)
        Next r
    End If

    Set mGLByNorm = d
    Set mGLByKey = ka
    Set mAlias = al
    mMapsBuilt = True
End Sub

Public Function V8_Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    V8_Norm = UCase$(Application.WorksheetFunction.Trim(Replace(CStr(v), ChrW(160), " ")))
End Function

Public Function V8_Month(ByVal v As Variant) As Long
    Select Case V8_Norm(v)
        Case "JAN", "JANUARY": V8_Month = 1
        Case "FEB", "FEBRUARY": V8_Month = 2
        Case "MAR", "MARCH": V8_Month = 3
        Case "APR", "APRIL": V8_Month = 4
        Case "MAY": V8_Month = 5
        Case "JUN", "JUNE": V8_Month = 6
        Case "JUL", "JULY": V8_Month = 7
        Case "AUG", "AUGUST": V8_Month = 8
        Case "SEP", "SEPTEMBER": V8_Month = 9
        Case "OCT", "OCTOBER": V8_Month = 10
        Case "NOV", "NOVEMBER": V8_Month = 11
        Case "DEC", "DECEMBER": V8_Month = 12
    End Select
End Function

Public Function V8_GLAccounts() As Object
    Dim d As Object, ws As Worksheet, r As Long, a As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    For r = 13 To ws.Cells(ws.Rows.count, 6).End(xlUp).Row
        a = V8_Norm(ws.Cells(r, 6).Value2)
        If Len(a) > 0 Then If Not d.Exists(a) Then d.Add a, CStr(ws.Cells(r, 6).Value2)
    Next r
    Set V8_GLAccounts = d
End Function

Public Function V8_ResolveAccount(ByVal raw As String) As String
    Dim k As String, d As Object, kd As String
    k = V8_Norm(raw): kd = V8_Key(raw)
    EnsureMaps
    Set d = mGLByNorm
    ' Operator-editable aliases on COA_MAP win over the built-in table, so a
    ' held label can be resolved without shipping a new build.
    If Len(kd) > 0 Then
        If mAlias.Exists(kd) Then V8_ResolveAccount = mAlias(kd): Exit Function
    End If
    Select Case k
        Case "TRANSPORATION AND TRAVEL": V8_ResolveAccount = "Transportation and Travel"
        Case "ADVANCES TO EMPLOYEE": V8_ResolveAccount = "Advances to Employees"
        Case "REPAIRS AND MAINTENANCE": V8_ResolveAccount = "Repair and Maintenance"
        Case "OUTPUT VAT PAYABLE": V8_ResolveAccount = "VAT Payable"
        Case "CLINIC SUPPLIES": V8_ResolveAccount = "Clinic Material and Supplies"
        ' v8.1 HOLD mappings are now executable runtime mappings.
        Case "MEDICAL EQUIPMENT": V8_ResolveAccount = "Dental Equipment"
        Case "COST OF REVENUE": V8_ResolveAccount = "Clinic Material and Supplies"
        Case "SUPPLIES": V8_ResolveAccount = "Clinic Material and Supplies"
        Case "BANK CHARGE": V8_ResolveAccount = "Miscellaneous"
        Case "CHARGES": V8_ResolveAccount = "Miscellaneous"
        Case Else
            If d.Exists(k) Then
                V8_ResolveAccount = d(k)
            ElseIf Len(kd) > 0 And mGLByKey.Exists(kd) Then
                ' Same account, different punctuation or spacing in the journal
                ' label.  Safe because the exact match above was already tried.
                V8_ResolveAccount = mGLByKey(kd)
            End If
    End Select
End Function

Public Function V8_CDJMap(ByVal col As Long) As String
    Select Case col
        Case 6: V8_CDJMap = "Cash in Bank"
        Case 7: V8_CDJMap = "Input VAT"
        Case 8: V8_CDJMap = "EWT Payable"
        Case 9: V8_CDJMap = "Petty Cash Fund"
        Case 10: V8_CDJMap = "Government Contributions (EE)"
        Case 11: V8_CDJMap = "Government Loans (EE)"
        Case 12: V8_CDJMap = "De Minimis"
        Case 13: V8_CDJMap = "Salaries and Wages"
        Case 14: V8_CDJMap = "Clinic Material and Supplies"
        Case 15: V8_CDJMap = "Rent"
        Case 16: V8_CDJMap = "Gas, Oil, Parking, Toll Fees"
        Case 17: V8_CDJMap = "Light and Water Expense"
        Case 18: V8_CDJMap = "Professional Fees"
        Case Else: V8_CDJMap = ""
    End Select
End Function

Public Function GLTitleExists(ByVal raw As String) As Boolean
    Dim kd As String
    EnsureMaps
    kd = V8_Key(raw)
    GLTitleExists = mGLByNorm.Exists(V8_Norm(raw)) Or _
                    (Len(kd) > 0 And mGLByKey.Exists(kd))
End Function

Public Function GLTitleOf(ByVal raw As String) As String
    Dim k As String, kd As String
    EnsureMaps
    k = V8_Norm(raw): kd = V8_Key(raw)
    If mGLByNorm.Exists(k) Then
        GLTitleOf = mGLByNorm(k)
    ElseIf Len(kd) > 0 And mGLByKey.Exists(kd) Then
        GLTitleOf = mGLByKey(kd)
    End If
End Function

Public Function V8_CRJMap(ByVal col As Long) As String
    Select Case col
        Case 8: V8_CRJMap = "Cash in Bank"
        Case 9: V8_CRJMap = "Excess CWT Over IT"
        Case 10: V8_CRJMap = "VAT Payable"
        Case 11, 12: V8_CRJMap = "Sales"
        Case Else: V8_CRJMap = ""
    End Select
End Function

Public Function V8_Number(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = True
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then ok = False Else V8_Number = CDbl(v)
End Function
```

### 6.14 `modGraphify` - Standalone charting utility.

- Source: `build/v8_5_vba/modGraphify.bas`
- Type: standard | Host: (none) | CodeName: `modGraphify`
- Lines: 229 | Nonblank: 201 | Procedures: 4
- SHA-256: `b94997601931b50f532a6c34f19554f2717840f2ef1fc9131385160003605f04`

**Analysis notes**

- Standalone chart utility, unrelated to accounting postings. Charts the first selected column, configured column A, or the busiest numeric column on the active sheet.
- Reuses a chart object named `GraphifyChart`; no workbook data is modified.

```vba
Attribute VB_Name = "modGraphify"
Option Explicit

'==============================================================================
' Graphify (v7.9, fixed) - charts one numeric column of the active sheet as a
' clustered column chart named "GraphifyChart" (reused on every re-run).
'
' Where the values come from (first match wins):
'   1. A multi-cell RANGE selected before running -> its first column
'   2. The configured column (CHART_COLUMN = "A", per the v7.8 spec)
'   3. Auto: the used-range column with the most numeric cells (leftmost on
'      tie) - this is what makes the macro work on the journal sheets, whose
'      data starts in column B ("Amount" is column I on PJ, DEBIT is F on GJ)
'   4. Otherwise: "No numeric data found" message
'
' Categories are R1..Rn (same labels as the graphify.html mock).
'
' Fixes over the broken v7.8 listing:
'   - A Worksheet has no .Charts collection -> use ChartObjects (embedded
'     chart); ws.Charts.Add would have added a chart SHEET and crashed first
'   - ws.Worksheets does not exist on a Worksheet object -> .Move removed
'   - Shape name "GraphifyChart.chart.8" never exists -> the ChartObject
'     itself is docked via .Left/.Top
'   - dataArr was read but never used -> removed; the source Range is bound
'     to the series directly (blank cells stay gaps instead of zeros)
'   - The Sub was never closed (End Sub missing in the listing)
'   - New: single-cell / no-numeric-data / chart-sheet edge cases handled;
'     errors surface through a handler that still restores ScreenUpdating
'==============================================================================

Private Const CHART_NAME As String = "GraphifyChart"
Private Const CHART_COLUMN As String = "A"   ' source column per the v7.8 spec
Private Const MIN_NUMERIC As Long = 2        ' auto-pick needs at least this many numbers
Private Const CHART_WIDTH As Double = 640    ' points
Private Const CHART_HEIGHT As Double = 360   ' points

Sub Graphify()
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim srs As Series
    Dim dataRng As Range
    Dim srcCol As Long          ' resolved source column index
    Dim colTop As Long          ' first row to consider in the source column
    Dim selBottom As Long       ' bottom row of the selection (0 = none)
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim nPts As Long
    Dim i As Long
    Dim cats() As String
    Dim srcDesc As String

    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Run Graphify from a worksheet (not a chart sheet).", vbExclamation, "Graphify"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail
    Application.ScreenUpdating = False

    colTop = 1
    srcCol = 0
    srcDesc = ""

    '--- 1. a multi-cell selection wins ---------------------------------------
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge > 1 Then
            srcCol = Selection.Columns(1).Column
            colTop = Selection.Row
            selBottom = Selection.Row + Selection.Rows.count - 1
            srcDesc = "selection"
        End If
    End If

    '--- 2. configured column (v7.8 spec: column A) ----------------------------
    If srcCol = 0 Then
        If Application.count(ws.Columns(CHART_COLUMN)) >= MIN_NUMERIC Then
            srcCol = ws.Columns(CHART_COLUMN).Column
            srcDesc = "column " & CHART_COLUMN
        End If
    End If

    '--- 3. auto-pick the busiest numeric column -------------------------------
    If srcCol = 0 Then
        srcCol = PickNumericColumn(ws)
        If srcCol > 0 Then srcDesc = "column " & ColumnLetter(ws, srcCol) & " (auto)"
    End If

    If srcCol = 0 Then
        MsgBox "No numeric data found on sheet '" & ws.name & "'." & vbCrLf & _
               "Select a column and run Graphify again.", vbExclamation, "Graphify"
        GoTo Bye
    End If

    '--- locate the numeric block within the resolved column -------------------
    firstDataRow = FirstNumericRow(ws, srcCol, colTop)
    lastRow = ws.Cells(ws.Rows.count, srcCol).End(xlUp).Row
    If selBottom > 0 And lastRow > selBottom Then lastRow = selBottom

    If firstDataRow = 0 Or lastRow < firstDataRow Then
        MsgBox "No numeric data found in " & srcDesc & " on sheet '" & ws.name & "'.", _
               vbExclamation, "Graphify"
        GoTo Bye
    End If

    Set dataRng = ws.Range(ws.Cells(firstDataRow, srcCol), ws.Cells(lastRow, srcCol))
    nPts = dataRng.Rows.count

    '--- category labels R1..Rn ------------------------------------------------
    ReDim cats(1 To nPts)
    For i = 1 To nPts
        cats(i) = "R" & i
    Next i

    '--- reuse the chart if it exists, otherwise create it ---------------------
    Set co = Nothing
    On Error Resume Next
    Set co = ws.ChartObjects(CHART_NAME)
    On Error GoTo Fail

    If co Is Nothing Then
        Set co = ws.ChartObjects.Add(ws.Range("E2").Left, ws.Range("E2").Top, _
                                     CHART_WIDTH, CHART_HEIGHT)
        co.name = CHART_NAME
    End If

    '--- rebuild the series from the resolved source ----------------------------
    With co.Chart
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        Set srs = .SeriesCollection.NewSeries
        srs.name = "Values"
        srs.Values = dataRng
        srs.XValues = cats

        .ChartType = xlColumnClustered     'change if you prefer xlLine, xlPie, etc.
        .HasTitle = True
        .ChartTitle.text = "Graphify - Column " & ColumnLetter(ws, srcCol)
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Caption = "Values"
        End With
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Caption = "Row"
        End With
        .HasLegend = False
    End With

    '--- dock the chart at E2 ---------------------------------------------------
    co.Left = ws.Range("E2").Left
    co.Top = ws.Range("E2").Top

Bye:
    Application.ScreenUpdating = True
    Set srs = Nothing
    Set co = Nothing
    Set dataRng = Nothing
    Set ws = Nothing
    Exit Sub

Fail:
    MsgBox "Graphify failed (" & Err.Number & "): " & Err.description, _
           vbExclamation, "Graphify"
    Resume Bye
End Sub

'------------------------------------------------------------------------------
' Column with the most numeric cells in the used range (leftmost on tie).
' Returns 0 when no column holds at least MIN_NUMERIC numbers.
'------------------------------------------------------------------------------
Private Function PickNumericColumn(ws As Worksheet) As Long
    Dim ur As Range
    Dim c As Long
    Dim best As Long
    Dim bestN As Long
    Dim n As Long
    Dim lastUsedRow As Long

    Set ur = ws.UsedRange
    lastUsedRow = ur.Row + ur.Rows.count - 1

    best = 0
    bestN = 0
    For c = ur.Column To ur.Column + ur.Columns.count - 1
        n = Application.count(ws.Range(ws.Cells(1, c), ws.Cells(lastUsedRow, c)))
        If n > bestN Then
            best = c
            bestN = n
        End If
    Next c

    If bestN >= MIN_NUMERIC Then
        PickNumericColumn = best
    Else
        PickNumericColumn = 0
    End If
End Function

'------------------------------------------------------------------------------
' First row at or below fromRow whose cell holds a real number
' (dates, booleans, text and blanks do not count). Returns 0 when none.
'------------------------------------------------------------------------------
Private Function FirstNumericRow(ws As Worksheet, colIdx As Long, fromRow As Long) As Long
    Dim r As Long
    Dim lastUsedRow As Long
    Dim v As Variant

    lastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    For r = fromRow To lastUsedRow
        v = ws.Cells(r, colIdx).Value
        Select Case VarType(v)
            Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency
                FirstNumericRow = r
                Exit Function
        End Select
    Next r
    FirstNumericRow = 0
End Function

'------------------------------------------------------------------------------
' Column index -> letter ("I").
'------------------------------------------------------------------------------
Private Function ColumnLetter(ws As Worksheet, colIdx As Long) As String
    Dim a As String
    a = ws.Columns(colIdx).address(False, False)   ' e.g. "I:I"
    ColumnLetter = Left$(a, InStr(a, ":") - 1)
End Function
```

### 6.15 `ImportSuppliersFromJSON` - Destructive hard-coded supplier seeder.

- Source: `build/v8_5_vba/ImportSuppliersFromJSON.bas`
- Type: standard | Host: (none) | CodeName: `ImportSuppliersFromJSON`
- Lines: 853 | Nonblank: 741 | Procedures: 1
- SHA-256: `256e140b5dfc190aac72bdfba872969faf120f488531e13434e726302fc85df9`

**Analysis notes**

- One-shot data seeder: clears every row from 2 to the bottom of `SUPPLIERS DATA`, then writes ~100 hard-coded suppliers and validates duplicate Name+Address pairs.
- Despite the module name, it does not read JSON. Running it is destructive to the supplier master.

```vba
Attribute VB_Name = "ImportSuppliersFromJSON"
Option Explicit

Sub ImportSuppliersFromJSON()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    
    ' Clear old data (keep headers)
    ws.Rows("2:" & ws.Rows.count).Clear
    
    rowNum = 2
    
    ' ============================================================
    ' COLUMNS: B=NO, C=TIN#, D=PARTICULARS, E=ADDRESS, F=CHART OF ACCOUNT
    ' ============================================================
    
    ' 1
    ws.Cells(rowNum, "B").Value = 1
    ws.Cells(rowNum, "C").Value = "008-888-659"
    ws.Cells(rowNum, "D").Value = "VPHARMA HEALTH AND WELNESS INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 2
    ws.Cells(rowNum, "B").Value = 2
    ws.Cells(rowNum, "C").Value = "246-099-058"
    ws.Cells(rowNum, "D").Value = "TRAVELLERS INTL HOTEL INC"
    ws.Cells(rowNum, "E").Value = "NEWPORT, PASAY"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 3
    ws.Cells(rowNum, "B").Value = 3
    ws.Cells(rowNum, "C").Value = "008-120-744"
    ws.Cells(rowNum, "D").Value = "RUSTAN"
    ws.Cells(rowNum, "E").Value = "SAN LORENZO, MAKATI"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 4
    ws.Cells(rowNum, "B").Value = 4
    ws.Cells(rowNum, "C").Value = "008-625-341"
    ws.Cells(rowNum, "D").Value = "THE MESSHALL"
    ws.Cells(rowNum, "E").Value = "MAGALLANES, MAKATI"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 5
    ws.Cells(rowNum, "B").Value = 5
    ws.Cells(rowNum, "C").Value = "252-126-611"
    ws.Cells(rowNum, "D").Value = "EUROFUEL INC"
    ws.Cells(rowNum, "E").Value = "183 PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 6
    ws.Cells(rowNum, "B").Value = 6
    ws.Cells(rowNum, "C").Value = "008-909-992"
    ws.Cells(rowNum, "D").Value = "SOUTHEASTASIA RETAIL INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 7
    ws.Cells(rowNum, "B").Value = 7
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE 39TH"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 8
    ws.Cells(rowNum, "B").Value = 8
    ws.Cells(rowNum, "C").Value = "207-961-175"
    ws.Cells(rowNum, "D").Value = "SAVEMORE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 9
    ws.Cells(rowNum, "B").Value = 9
    ws.Cells(rowNum, "C").Value = "001-799-123"
    ws.Cells(rowNum, "D").Value = "UNIGLOBE TRAVELWARE INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 10
    ws.Cells(rowNum, "B").Value = 10
    ws.Cells(rowNum, "C").Value = "148-855-506"
    ws.Cells(rowNum, "D").Value = "MINDWERKS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 11
    ws.Cells(rowNum, "B").Value = 11
    ws.Cells(rowNum, "C").Value = "007-851-409"
    ws.Cells(rowNum, "D").Value = "PRIMERA CLASE GAS STATION"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 12
    ws.Cells(rowNum, "B").Value = 12
    ws.Cells(rowNum, "C").Value = "200-035-311"
    ws.Cells(rowNum, "D").Value = "ACE HARDWARE"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 13
    ws.Cells(rowNum, "B").Value = 13
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE BLUEBAY WALK"
    ws.Cells(rowNum, "E").Value = "76 PASAY"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 14
    ws.Cells(rowNum, "B").Value = 14
    ws.Cells(rowNum, "C").Value = "246-969-491"
    ws.Cells(rowNum, "D").Value = "S AND R"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 15
    ws.Cells(rowNum, "B").Value = 15
    ws.Cells(rowNum, "C").Value = "648-700-952"
    ws.Cells(rowNum, "D").Value = "GOOD MORNING FOOD"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 16
    ws.Cells(rowNum, "B").Value = 16
    ws.Cells(rowNum, "C").Value = "008-087-158"
    ws.Cells(rowNum, "D").Value = "PNC GARBROS FOODS INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 17
    ws.Cells(rowNum, "B").Value = 17
    ws.Cells(rowNum, "C").Value = "208-006-656"
    ws.Cells(rowNum, "D").Value = "SM STORE"
    ws.Cells(rowNum, "E").Value = "ORANBO, PASIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 18
    ws.Cells(rowNum, "B").Value = 18
    ws.Cells(rowNum, "C").Value = "214-706-591"
    ws.Cells(rowNum, "D").Value = "WATSONS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 19
    ws.Cells(rowNum, "B").Value = 19
    ws.Cells(rowNum, "C").Value = "608-395-184"
    ws.Cells(rowNum, "D").Value = "GIANETO PIZZERIA"
    ws.Cells(rowNum, "E").Value = "ORANBO, PASIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 20
    ws.Cells(rowNum, "B").Value = 20
    ws.Cells(rowNum, "C").Value = "010-624-401"
    ws.Cells(rowNum, "D").Value = "MIT SUKOSHI FRESH"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 21
    ws.Cells(rowNum, "B").Value = 21
    ws.Cells(rowNum, "C").Value = "655-591-162"
    ws.Cells(rowNum, "D").Value = "MINISO"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 22
    ws.Cells(rowNum, "B").Value = 22
    ws.Cells(rowNum, "C").Value = "000-405-340"
    ws.Cells(rowNum, "D").Value = "ROBINSONS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 23
    ws.Cells(rowNum, "B").Value = 23
    ws.Cells(rowNum, "C").Value = "000-123-826"
    ws.Cells(rowNum, "D").Value = "GRAND UNION SUPERMARKET INC"
    ws.Cells(rowNum, "E").Value = "FILINVEST, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 24
    ws.Cells(rowNum, "B").Value = 24
    ws.Cells(rowNum, "C").Value = "008-188-059"
    ws.Cells(rowNum, "D").Value = "KNOXPORT INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 25
    ws.Cells(rowNum, "B").Value = 25
    ws.Cells(rowNum, "C").Value = "009-754-702"
    ws.Cells(rowNum, "D").Value = "GOOD SMALL SHEEP INC"
    ws.Cells(rowNum, "E").Value = "DAANG BAKAL, MANDALUYONG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 26
    ws.Cells(rowNum, "B").Value = 26
    ws.Cells(rowNum, "C").Value = "163-955-615"
    ws.Cells(rowNum, "D").Value = "725 TAGTEAM FUEL STATION"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 27
    ws.Cells(rowNum, "B").Value = 27
    ws.Cells(rowNum, "C").Value = "007-875-333"
    ws.Cells(rowNum, "D").Value = "SOUTHERN ONGVILLE CORP"
    ws.Cells(rowNum, "E").Value = "CANDELARIA, QUEZON"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 28
    ws.Cells(rowNum, "B").Value = 28
    ws.Cells(rowNum, "C").Value = "905-035-955"
    ws.Cells(rowNum, "D").Value = "N MART GENERAL MERCHANDIZE"
    ws.Cells(rowNum, "E").Value = "BAMBANG, NUEVA ECIJA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 29
    ws.Cells(rowNum, "B").Value = 29
    ws.Cells(rowNum, "C").Value = "241-838-060"
    ws.Cells(rowNum, "D").Value = "FACIAL CARE CENTRE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 30
    ws.Cells(rowNum, "B").Value = 30
    ws.Cells(rowNum, "C").Value = "009-149--884"
    ws.Cells(rowNum, "D").Value = "LANDMARK SUPERMARKET"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 31
    ws.Cells(rowNum, "B").Value = 31
    ws.Cells(rowNum, "C").Value = "240-940-697"
    ws.Cells(rowNum, "D").Value = "INENGS"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 32
    ws.Cells(rowNum, "B").Value = 32
    ws.Cells(rowNum, "C").Value = "009-271-384"
    ws.Cells(rowNum, "D").Value = "OTADER DENTAL AND MEDICAL"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 33
    ws.Cells(rowNum, "B").Value = 33
    ws.Cells(rowNum, "C").Value = "000-340-727"
    ws.Cells(rowNum, "D").Value = "ALPHADENT CORPORATION"
    ws.Cells(rowNum, "E").Value = "SAN ANTONIO, QUEZON"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 34
    ws.Cells(rowNum, "B").Value = 34
    ws.Cells(rowNum, "C").Value = "006-614-702"
    ws.Cells(rowNum, "D").Value = "PROS APAC CORP"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 35
    ws.Cells(rowNum, "B").Value = 35
    ws.Cells(rowNum, "C").Value = "003-935-334"
    ws.Cells(rowNum, "D").Value = "ST PATRICK MEDICAL SYSTEMS INC"
    ws.Cells(rowNum, "E").Value = "SHAW BLVD., MANDALUYONG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 36
    ws.Cells(rowNum, "B").Value = 36
    ws.Cells(rowNum, "C").Value = "216-037-865"
    ws.Cells(rowNum, "D").Value = "DENTAL DOMAIN CORP"
    ws.Cells(rowNum, "E").Value = "CULIAT, QUEZON CITY"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 37
    ws.Cells(rowNum, "B").Value = 37
    ws.Cells(rowNum, "C").Value = "226-527-915"
    ws.Cells(rowNum, "D").Value = "METRO MARKET MARKET"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 38
    ws.Cells(rowNum, "B").Value = 38
    ws.Cells(rowNum, "C").Value = "010-176-614"
    ws.Cells(rowNum, "D").Value = "GREEN DENTAL LAB"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 39
    ws.Cells(rowNum, "B").Value = 39
    ws.Cells(rowNum, "C").Value = "440-275-225"
    ws.Cells(rowNum, "D").Value = "AAHA GAS RETAILING"
    ws.Cells(rowNum, "E").Value = "PAMPLONA UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 40
    ws.Cells(rowNum, "B").Value = 40
    ws.Cells(rowNum, "C").Value = "010-119-810"
    ws.Cells(rowNum, "D").Value = "RJD MERRY HARVEST CORP"
    ws.Cells(rowNum, "E").Value = "SILANG, CAVITE"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 41
    ws.Cells(rowNum, "B").Value = 41
    ws.Cells(rowNum, "C").Value = "000-522-680"
    ws.Cells(rowNum, "D").Value = "ALBA INTERNATIONAL INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 42
    ws.Cells(rowNum, "B").Value = 42
    ws.Cells(rowNum, "C").Value = "007-551-518"
    ws.Cells(rowNum, "D").Value = "NUTRITION FOR LIFE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 43
    ws.Cells(rowNum, "B").Value = 43
    ws.Cells(rowNum, "C").Value = "008-664-864"
    ws.Cells(rowNum, "D").Value = "GRANDWEN GAS CORP"
    ws.Cells(rowNum, "E").Value = "BACLARAN, PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 44
    ws.Cells(rowNum, "B").Value = 44
    ws.Cells(rowNum, "C").Value = "009-195-748"
    ws.Cells(rowNum, "D").Value = "BIG SKY NATION INC"
    ws.Cells(rowNum, "E").Value = "183 PASAY"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 45
    ws.Cells(rowNum, "B").Value = 45
    ws.Cells(rowNum, "C").Value = "241-437-266"
    ws.Cells(rowNum, "D").Value = "YELOW CAB PIZZA"
    ws.Cells(rowNum, "E").Value = "STA. ROSA, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 46
    ws.Cells(rowNum, "B").Value = 46
    ws.Cells(rowNum, "C").Value = "636-985956"
    ws.Cells(rowNum, "D").Value = "WILSON MOA"
    ws.Cells(rowNum, "E").Value = "MOA, PASAY"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 47
    ws.Cells(rowNum, "B").Value = 47
    ws.Cells(rowNum, "C").Value = "206-114-170"
    ws.Cells(rowNum, "D").Value = "PETRON EXPRESS CENTER"
    ws.Cells(rowNum, "E").Value = "BALAGTAS, BULACAN"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 48
    ws.Cells(rowNum, "B").Value = 48
    ws.Cells(rowNum, "C").Value = "005-148-299"
    ws.Cells(rowNum, "D").Value = "ERNA TRADINGCORP"
    ws.Cells(rowNum, "E").Value = "TALON 4, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 49
    ws.Cells(rowNum, "B").Value = 49
    ws.Cells(rowNum, "C").Value = "008-367-582"
    ws.Cells(rowNum, "D").Value = "JAVELLANA ADVANCED AESTHETICS CLINIC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 50
    ws.Cells(rowNum, "B").Value = 50
    ws.Cells(rowNum, "C").Value = "226-784-068-000"
    ws.Cells(rowNum, "D").Value = "MDI Group Holdings Inc"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG CITY"
    ws.Cells(rowNum, "F").Value = "Rental"
    rowNum = rowNum + 1
    
    ' 51
    ws.Cells(rowNum, "B").Value = 51
    ws.Cells(rowNum, "C").Value = "000-405-340"
    ws.Cells(rowNum, "D").Value = "ROBINSONS"
    ws.Cells(rowNum, "E").Value = "TALON UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 52
    ws.Cells(rowNum, "B").Value = 52
    ws.Cells(rowNum, "C").Value = "007-189-834"
    ws.Cells(rowNum, "D").Value = "KENNY ROGERS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 53
    ws.Cells(rowNum, "B").Value = 53
    ws.Cells(rowNum, "C").Value = "001-096-140"
    ws.Cells(rowNum, "D").Value = "ALMON MARINA INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 54
    ws.Cells(rowNum, "B").Value = 54
    ws.Cells(rowNum, "C").Value = "236-414-781"
    ws.Cells(rowNum, "D").Value = "YOKI FARM"
    ws.Cells(rowNum, "E").Value = "PALOCPOC, CAVITE"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 55
    ws.Cells(rowNum, "B").Value = 55
    ws.Cells(rowNum, "C").Value = "203-301-457"
    ws.Cells(rowNum, "D").Value = "CINNABON"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 56
    ws.Cells(rowNum, "B").Value = 56
    ws.Cells(rowNum, "C").Value = "000-388-474"
    ws.Cells(rowNum, "D").Value = "MERCURY"
    ws.Cells(rowNum, "E").Value = "TALON 5, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 57
    ws.Cells(rowNum, "B").Value = 57
    ws.Cells(rowNum, "C").Value = "000-123-826"
    ws.Cells(rowNum, "D").Value = "SOUTH SUPERMARKET"
    ws.Cells(rowNum, "E").Value = "FILINVEST, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 58
    ws.Cells(rowNum, "B").Value = 58
    ws.Cells(rowNum, "C").Value = "241-498-356"
    ws.Cells(rowNum, "D").Value = "HONEYBON"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 59
    ws.Cells(rowNum, "B").Value = 59
    ws.Cells(rowNum, "C").Value = "000-388-771"
    ws.Cells(rowNum, "D").Value = "JOLLIBEE"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 60
    ws.Cells(rowNum, "B").Value = 60
    ws.Cells(rowNum, "C").Value = "610-884-579"
    ws.Cells(rowNum, "D").Value = "T RIPLE MLD CORP"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 61
    ws.Cells(rowNum, "B").Value = 61
    ws.Cells(rowNum, "C").Value = "607-237-567"
    ws.Cells(rowNum, "D").Value = "RUSTIC PIZZA PASTA RESTAURANT INC"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 62
    ws.Cells(rowNum, "B").Value = 62
    ws.Cells(rowNum, "C").Value = "005-148-299"
    ws.Cells(rowNum, "D").Value = "ERNA TRADING CORP"
    ws.Cells(rowNum, "E").Value = "TALON 4, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 63
    ws.Cells(rowNum, "B").Value = 63
    ws.Cells(rowNum, "C").Value = "010-892-033"
    ws.Cells(rowNum, "D").Value = "ELEVATE PILATES COMPANY LIMITED INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 64
    ws.Cells(rowNum, "B").Value = 64
    ws.Cells(rowNum, "C").Value = "200-005-881"
    ws.Cells(rowNum, "D").Value = "ARISTOCRAT BAKESHOP"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 65
    ws.Cells(rowNum, "B").Value = 65
    ws.Cells(rowNum, "C").Value = "000-060-696"
    ws.Cells(rowNum, "D").Value = "RUSTANS"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Representation"
    rowNum = rowNum + 1
    
    ' 66
    ws.Cells(rowNum, "B").Value = 66
    ws.Cells(rowNum, "C").Value = "909-079-780"
    ws.Cells(rowNum, "D").Value = "INGUSAN GASOLINE STATION"
    ws.Cells(rowNum, "E").Value = "TALON UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 67
    ws.Cells(rowNum, "B").Value = 67
    ws.Cells(rowNum, "C").Value = "000-299-299"
    ws.Cells(rowNum, "D").Value = "ABACUS BOOK AND CARD CORP"
    ws.Cells(rowNum, "E").Value = "POBLACION, MAKATI"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 68
    ws.Cells(rowNum, "B").Value = 68
    ws.Cells(rowNum, "C").Value = "003-058-789"
    ws.Cells(rowNum, "D").Value = "SM PRIME HOLDINGS INC"
    ws.Cells(rowNum, "E").Value = "ALMANZA UNO, LAS PI�AS"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 69
    ws.Cells(rowNum, "B").Value = 69
    ws.Cells(rowNum, "C").Value = "183-378-131"
    ws.Cells(rowNum, "D").Value = "MG PETRON CORDON SERVICE STATION"
    ws.Cells(rowNum, "E").Value = "CORDON, ISABELA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 70
    ws.Cells(rowNum, "B").Value = 70
    ws.Cells(rowNum, "C").Value = "630-167-414"
    ws.Cells(rowNum, "D").Value = "NAIL GARDEN"
    ws.Cells(rowNum, "E").Value = "76 PASAY"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 71
    ws.Cells(rowNum, "B").Value = 71
    ws.Cells(rowNum, "C").Value = "230-577-870"
    ws.Cells(rowNum, "D").Value = "LALONG FUEL STATION"
    ws.Cells(rowNum, "E").Value = "SAN PABLO, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Fuel and Oil"
    rowNum = rowNum + 1
    
    ' 72
    ws.Cells(rowNum, "B").Value = 72
    ws.Cells(rowNum, "C").Value = "008-909-992"
    ws.Cells(rowNum, "D").Value = "LANDERS"
    ws.Cells(rowNum, "E").Value = "TAMBO, PARA�AQUE"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 73
    ws.Cells(rowNum, "B").Value = 73
    ws.Cells(rowNum, "C").Value = "207-961-175"
    ws.Cells(rowNum, "D").Value = "FESTIVAL MALL"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 74
    ws.Cells(rowNum, "B").Value = 74
    ws.Cells(rowNum, "C").Value = "010-262-979"
    ws.Cells(rowNum, "D").Value = "LIN ET VIE PELLE MERCHANDISING"
    ws.Cells(rowNum, "E").Value = "ALABANG, MUNTINLUPA"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 75
    ws.Cells(rowNum, "B").Value = 75
    ws.Cells(rowNum, "C").Value = "005-695-758"
    ws.Cells(rowNum, "D").Value = "UNICITY"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 76
    ws.Cells(rowNum, "B").Value = 76
    ws.Cells(rowNum, "C").Value = "010-176-614"
    ws.Cells(rowNum, "D").Value = "GREEN DELTA LAB PHILIPPINES INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 77
    ws.Cells(rowNum, "B").Value = 77
    ws.Cells(rowNum, "C").Value = "260-019-682"
    ws.Cells(rowNum, "D").Value = "BURDA MEDICAL APPAREL"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 78
    ws.Cells(rowNum, "B").Value = 78
    ws.Cells(rowNum, "C").Value = "278-353-148"
    ws.Cells(rowNum, "D").Value = "D AND L MEDICAL APPAREL"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 79
    ws.Cells(rowNum, "B").Value = 79
    ws.Cells(rowNum, "C").Value = "778-866-700"
    ws.Cells(rowNum, "D").Value = "GREEN WAVE COMPUTER SYSTEM"
    ws.Cells(rowNum, "E").Value = "SANTA CRUZ, MANILA"
    ws.Cells(rowNum, "F").Value = "Repairs and Maintenance"
    rowNum = rowNum + 1
    
    ' 80
    ws.Cells(rowNum, "B").Value = 80
    ws.Cells(rowNum, "C").Value = "010-559-959"
    ws.Cells(rowNum, "D").Value = "FRONTIER DENTAL PRODUCTS CORP"
    ws.Cells(rowNum, "E").Value = "SAN ANTONIO, MAKATI"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 81
    ws.Cells(rowNum, "B").Value = 81
    ws.Cells(rowNum, "C").Value = "000-324-286"
    ws.Cells(rowNum, "D").Value = "GOLDEN PEAK SALES CORP"
    ws.Cells(rowNum, "E").Value = "SAN NICOLAS, MANILA"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 82
    ws.Cells(rowNum, "B").Value = 82
    ws.Cells(rowNum, "C").Value = "225-527-914"
    ws.Cells(rowNum, "D").Value = "METRO RETAIL STORES GROUP INC"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 83
    ws.Cells(rowNum, "B").Value = 83
    ws.Cells(rowNum, "C").Value = "208-544-458"
    ws.Cells(rowNum, "D").Value = "MKYR TRANSPORT SERVICES"
    ws.Cells(rowNum, "E").Value = "BATASAN HILLS, QUEZON CITY"
    ws.Cells(rowNum, "F").Value = "Transportation and Travel"
    rowNum = rowNum + 1
    
    ' 84
    ws.Cells(rowNum, "B").Value = 84
    ws.Cells(rowNum, "C").Value = "006-614-702"
    ws.Cells(rowNum, "D").Value = "PROS-APAC CORP"
    ws.Cells(rowNum, "E").Value = "MALATE, MANILA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' 85
    ws.Cells(rowNum, "B").Value = 85
    ws.Cells(rowNum, "C").Value = "201-619-152"
    ws.Cells(rowNum, "D").Value = "VILLACORTA DENTAL LABORATY INC"
    ws.Cells(rowNum, "E").Value = "SANTA ROSA, LAGUNA"
    ws.Cells(rowNum, "F").Value = "Clinic Materials and Supplies"
    rowNum = rowNum + 1
    
    ' ROW 86 REMOVED - DUPLICATE
    
    ' 87
    ws.Cells(rowNum, "B").Value = 87
    ws.Cells(rowNum, "C").Value = "120-568-073"
    ws.Cells(rowNum, "D").Value = "7-ELEVEN"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 88
    ws.Cells(rowNum, "B").Value = 88
    ws.Cells(rowNum, "C").Value = "000-388-474"
    ws.Cells(rowNum, "D").Value = "MERCURY DRUG"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 89
    ws.Cells(rowNum, "B").Value = 89
    ws.Cells(rowNum, "C").Value = "233-251-708"
    ws.Cells(rowNum, "D").Value = "FEDERAL BRENT RETAIL INC"
    ws.Cells(rowNum, "E").Value = "076 PASAY"
    ws.Cells(rowNum, "F").Value = "Supplies"
    rowNum = rowNum + 1
    
    ' 90
    ws.Cells(rowNum, "B").Value = 90
    ws.Cells(rowNum, "C").Value = "214-706-591"
    ws.Cells(rowNum, "D").Value = "WATSONS"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Miscellaneous"
    rowNum = rowNum + 1
    
    ' 91
    ws.Cells(rowNum, "B").Value = 91
    ws.Cells(rowNum, "C").Value = "226-784-068"
    ws.Cells(rowNum, "D").Value = "MDI Group Holdings Inc 2"
    ws.Cells(rowNum, "E").Value = "BGC, TAGUIG"
    ws.Cells(rowNum, "F").Value = "Association Dues"
    rowNum = rowNum + 1
    
    ' 92
    ws.Cells(rowNum, "B").Value = 92
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Rosalina Ignacio"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 93
    ws.Cells(rowNum, "B").Value = 93
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Daisy Cornejo"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 94
    ws.Cells(rowNum, "B").Value = 94
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Janina Tayag"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 95
    ws.Cells(rowNum, "B").Value = 95
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Jasper Tago"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Corporators"
    rowNum = rowNum + 1
    
    ' 96
    ws.Cells(rowNum, "B").Value = 96
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Ana Ramona Locsin"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 97
    ws.Cells(rowNum, "B").Value = 97
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Virgilio Malenab Jr"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 98
    ws.Cells(rowNum, "B").Value = 98
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "SBM Consulting Inc"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Professional Fee"
    rowNum = rowNum + 1
    
    ' 99
    ws.Cells(rowNum, "B").Value = 99
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Christian Anthony Ermita"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' 100
    ws.Cells(rowNum, "B").Value = 100
    ws.Cells(rowNum, "C").Value = "TIN NOT FOUND"
    ws.Cells(rowNum, "D").Value = "Vergilyn Montales"
    ws.Cells(rowNum, "E").Value = "�"
    ws.Cells(rowNum, "F").Value = "Clinician's Fee-Visiting"
    rowNum = rowNum + 1
    
    ' ---- VALIDATE FOR DUPLICATES ----
    ' Check: Same Name + Same Address = NOT ALLOWED
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    Dim duplicateFound As Boolean
    duplicateFound = False
    
    For i = 2 To lastRow
        Dim name As String
        Dim address As String
        Dim j As Long
        
        name = ws.Cells(i, "D").Value
        address = ws.Cells(i, "E").Value
        
        If name <> "" And address <> "" And address <> "�" Then
            For j = i + 1 To lastRow
                Dim name2 As String
                Dim address2 As String
                
                name2 = ws.Cells(j, "D").Value
                address2 = ws.Cells(j, "E").Value
                
                If name = name2 And address = address2 Then
                    MsgBox "DUPLICATE FOUND:" & vbCrLf & _
                           "Row " & i & ": " & name & " | " & address & vbCrLf & _
                           "Row " & j & ": " & name2 & " | " & address2 & vbCrLf & vbCrLf & _
                           "Same Name + Address is NOT allowed!", vbCritical
                    duplicateFound = True
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    If Not duplicateFound Then
        MsgBox "Import complete! " & (rowNum - 2) & " suppliers imported." & vbCrLf & _
               "All suppliers have unique Name + Address combinations.", vbInformation
    End If
End Sub
```

## 7. Repository provenance

| Path | What it contains | Use in this analysis |
|---|---|---|
| `build/v8_5_vba/` | Full v8.5 VBA export, one file per component | **Canonical source for this document** |
| `vba/` | v8.4.1 copies of `modGLGate`, `modGLRefresh`, `modGLScanEngine`, `modGLWorkbookMap` | Same logic as the v8.5 export except the export header; not reproduced separately |
| `build/baseline_vba/` | Baseline export used by the v8.5 build diff | Historical comparison only |
| `build/post_vba/` | Post-patch Sheet11/Sheet12 exports | Historical build artifact |
| `patch/claude code/` | v7.9.5 `modEngine` patch and PJ sheet patch | Historical patch source; not the current superset |
| `_inspect/all_vba.txt` | Older single-file VBA dump | Superseded by the per-component v8.5 export |

### Build/verification artifacts referenced by the code

- `build/verify_summary.json`: v8.5 changed modules (`Sheet11`, `Sheet12`, `modEngine`) and workbook SHA-256.
- `build/artifacts/compile.log`: recorded `compile PASS`.
- `GL_REBUILD_VERIFICATION_REPORT.md`: v8.4.1 live refresh returned `HOLD` on the duplicate GL title and unmapped Sundry amounts.
- `scripts/verify_full_gl.py` and `tests/`: source-level checks for scope and duplicate-account guards.

## 8. Quick reference: what to inspect for a specific question

| Question | Start here |
|---|---|
| Why did a GL refresh HOLD? | `modGLScanEngine.ScanGL` and `WriteScanAudit`; then inspect `GL_AUDIT` |
| Where is a CDJ/CRJ row posted? | `modEngine.SyncPJToCDJ`, `SyncPJNonVatToCDJ`, `SyncSJToCRJ`; `modPJSync.SyncToCDJ_NonVat` for legacy |
| Where is a new CDJ row placed? | `modEngine.FindCDJFreeRow` and `CDJReserveRow` |
| Where are supplier TIN/name matches resolved? | `modEngine.GetSupplierByTIN`, `GetSupplierByName`, `AddSupplierRecord` |
| Where are journal labels mapped to GL accounts? | `modGLWorkbookMap.V8_ResolveAccount`, `V8_CDJMap`, `V8_CRJMap`, `EnsureMaps` |
| Which macros appear in Alt+F8? | Zero-argument `Public` procedures, especially `modGLRefresh`, `modGLGate.RunV8Gate`, `modEngine` status/reset helpers, `modGraphify.Graphify`, `ImportSuppliersFromJSON` |
| Which code is legacy and not on the active GL path? | `modEngine.GLScanGJ` / `Legacy_*`, `BuildV8CalcSheet`, Sheet9 and `modPJSync` |
