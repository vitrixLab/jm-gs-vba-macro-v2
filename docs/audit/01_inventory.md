# Phase 1 - VBA inventory

Canonical source: `build/v8_5_vba/`. Counts and declarations were parsed from the exported text, not from Excel.

## Component counts

| Module | Type | File | Lines | Nonblank | Procedures | Event handlers |
|---|---|---|---:|---:|---:|---|
| `ThisWorkbook` | document | `ThisWorkbook.txt` | 10 | 8 | 1 | Workbook_Open |
| `Sheet12` | document | `Sheet12.txt` | 290 | 255 | 9 | Worksheet_Change |
| `Sheet11` | document | `Sheet11.txt` | 294 | 260 | 9 | Worksheet_Change |
| `Sheet9` | document | `Sheet9.txt` | 295 | 204 | 1 | Worksheet_Change |
| `Sheet3` | document | `Sheet3.txt` | 175 | 147 | 6 | Worksheet_Activate, Worksheet_Change |
| `Sheet8` | document | `Sheet8.txt` | 88 | 74 | 2 | Worksheet_Change |
| `Sheet10` | document | `Sheet10.txt` | 34 | 31 | 1 | - |
| `modEngine` | standard | `modEngine.bas` | 1781 | 1668 | 76 | - |
| `modPJSync` | standard | `modPJSync.bas` | 532 | 257 | 8 | - |
| `modGLGate` | standard | `modGLGate.bas` | 51 | 43 | 3 | - |
| `modGLRefresh` | standard | `modGLRefresh.bas` | 28 | 22 | 4 | - |
| `modGLScanEngine` | standard | `modGLScanEngine.bas` | 335 | 292 | 9 | - |
| `modGLWorkbookMap` | standard | `modGLWorkbookMap.bas` | 211 | 189 | 12 | - |
| `modGraphify` | standard | `modGraphify.bas` | 229 | 201 | 4 | - |
| `ImportSuppliersFromJSON` | standard | `ImportSuppliersFromJSON.bas` | 856 | 741 | 1 | - |

Total: 15 non-empty components, 146 declared procedures, 8 standard modules, 6 worksheet modules and `ThisWorkbook`.

## Event entry points

| Handler | Line | Trigger |
|---|---:|---|
| `ThisWorkbook.Workbook_Open` | 3 | Open |
| `Sheet12.Worksheet_Change` | 192 | Change |
| `Sheet11.Worksheet_Change` | 195 | Change |
| `Sheet9.Worksheet_Change` | 2 | Change |
| `Sheet3.Worksheet_Activate` | 21 | Activate |
| `Sheet3.Worksheet_Change` | 39 | Change |
| `Sheet8.Worksheet_Change` | 30 | Change |

No `UserForm_*`, `Class_*`, `Auto_*` or `Application.OnTime` entry points were found.

## Procedure inventory

| Module | Line | Kind | Scope | Procedure |
|---|---:|---|---|---|
| `ThisWorkbook` | 3 | Sub | Private | `Workbook_Open` |
| `Sheet12` | 19 | Function | Private | `GetCOAList` |
| `Sheet12` | 34 | Function | Private | `SelectCOA` |
| `Sheet12` | 55 | Sub | Private | `LookupByTIN` |
| `Sheet12` | 101 | Sub | Private | `LookupByName` |
| `Sheet12` | 144 | Sub | Private | `ClearSupplierRow` |
| `Sheet12` | 150 | Sub | Private | `CalculateTotals` |
| `Sheet12` | 162 | Function | Private | `FindBlockStart` |
| `Sheet12` | 174 | Sub | Implicit | `ManualSyncToCDJ` |
| `Sheet12` | 192 | Sub | Private | `Worksheet_Change` |
| `Sheet11` | 19 | Function | Private | `GetCOAList` |
| `Sheet11` | 34 | Function | Private | `SelectCOA` |
| `Sheet11` | 55 | Sub | Private | `LookupByTIN` |
| `Sheet11` | 101 | Sub | Private | `LookupByName` |
| `Sheet11` | 144 | Sub | Private | `ClearSupplierRow` |
| `Sheet11` | 153 | Sub | Private | `CalculateTotals` |
| `Sheet11` | 165 | Function | Private | `FindBlockStart` |
| `Sheet11` | 177 | Sub | Implicit | `ManualSyncToCDJ` |
| `Sheet11` | 195 | Sub | Private | `Worksheet_Change` |
| `Sheet9` | 2 | Sub | Private | `Worksheet_Change` |
| `Sheet3` | 21 | Sub | Private | `Worksheet_Activate` |
| `Sheet3` | 27 | Sub | Private | `LookupCustomer` |
| `Sheet3` | 34 | Sub | Private | `ClearCustomerRow` |
| `Sheet3` | 39 | Sub | Private | `Worksheet_Change` |
| `Sheet3` | 137 | Sub | Implicit | `ManualSyncToCRJ` |
| `Sheet3` | 171 | Function | Private | `SheetExists` |
| `Sheet8` | 23 | Sub | Private | `SyncToCRJ` |
| `Sheet8` | 30 | Sub | Private | `Worksheet_Change` |
| `Sheet10` | 1 | Sub | Implicit | `CheckDuplicates` |
| `modEngine` | 129 | Sub | Public | `InitializeEngine` |
| `modEngine` | 145 | Sub | Public | `InitializeCRJ` |
| `modEngine` | 158 | Function | Private | `CellText` |
| `modEngine` | 168 | Function | Private | `NormName` |
| `modEngine` | 177 | Sub | Private | `NewDupDict` |
| `modEngine` | 182 | Function | Private | `IsDupKey` |
| `modEngine` | 191 | Sub | Private | `RegisterSupplier` |
| `modEngine` | 213 | Sub | Private | `LoadSuppliers` |
| `modEngine` | 236 | Sub | Public | `ReloadSuppliers` |
| `modEngine` | 250 | Function | Public | `NormTIN` |
| `modEngine` | 261 | Function | Public | `AmbiguousTINCount` |
| `modEngine` | 268 | Function | Public | `AmbiguousTINList` |
| `modEngine` | 278 | Function | Public | `IsAmbiguousTIN` |
| `modEngine` | 284 | Function | Public | `AmbiguousNamesFor` |
| `modEngine` | 291 | Sub | Private | `LoadCDJHeaders` |
| `modEngine` | 330 | Sub | Private | `LoadCDJReferences` |
| `modEngine` | 362 | Function | Private | `CDJRowIsFree` |
| `modEngine` | 372 | Function | Public | `CDJMarkerMonth` |
| `modEngine` | 380 | Function | Public | `CDJMonthAbbr` |
| `modEngine` | 398 | Function | Public | `CDJTotalRow` |
| `modEngine` | 410 | Function | Public | `FindCDJFreeRow` |
| `modEngine` | 436 | Function | Public | `CDJReserveRow` |
| `modEngine` | 466 | Sub | Private | `ShiftCDJRefCache` |
| `modEngine` | 475 | Function | Public | `PJMonthOf` |
| `modEngine` | 484 | Sub | Private | `LoadSequence` |
| `modEngine` | 494 | Sub | Private | `LoadNextRow` |
| `modEngine` | 502 | Sub | Private | `LoadCRJReferences` |
| `modEngine` | 516 | Sub | Private | `LoadCRJNextRow` |
| `modEngine` | 533 | Sub | Private | `LoadCRJSequence` |
| `modEngine` | 546 | Function | Public | `GetNextPJNumber` |
| `modEngine` | 557 | Function | Public | `GetNextInvoiceNumber` |
| `modEngine` | 568 | Function | Public | `GetNextSeq` |
| `modEngine` | 582 | Function | Public | `GetSupplierByTIN` |
| `modEngine` | 598 | Function | Public | `GetSupplierByName` |
| `modEngine` | 615 | Sub | Public | `AddSupplierToCache` |
| `modEngine` | 631 | Function | Public | `AddSupplierRecord` |
| `modEngine` | 714 | Function | Public | `FindExpenseColumn` |
| `modEngine` | 766 | Function | Public | `CalculateWithholdingTax` |
| `modEngine` | 784 | Function | Public | `FindCDJRow` |
| `modEngine` | 818 | Sub | Public | `UpdateCDJReference` |
| `modEngine` | 824 | Function | Public | `FindCRJRow` |
| `modEngine` | 831 | Sub | Public | `UpdateCRJReference` |
| `modEngine` | 840 | Sub | Public | `SyncPJToCDJ` |
| `modEngine` | 966 | Sub | Public | `SyncSJToCRJ` |
| `modEngine` | 1041 | Function | Public | `EngineStatus` |
| `modEngine` | 1060 | Sub | Public | `ShowCDJStatus` |
| `modEngine` | 1087 | Sub | Public | `ShowCRJStatus` |
| `modEngine` | 1110 | Sub | Public | `ShowAmbiguousTINs` |
| `modEngine` | 1126 | Sub | Public | `RepairEngine` |
| `modEngine` | 1163 | Sub | Public | `ResetCRJ` |
| `modEngine` | 1189 | Sub | Public | `ResetCDJ` |
| `modEngine` | 1218 | Function | Public | `CalculateEWT` |
| `modEngine` | 1238 | Sub | Public | `SyncPJNonVatToCDJ` |
| `modEngine` | 1334 | Function | Private | `GLN` |
| `modEngine` | 1343 | Function | Private | `GLAcct` |
| `modEngine` | 1396 | Function | Private | `GLMonth` |
| `modEngine` | 1416 | Sub | Private | `GLAdd` |
| `modEngine` | 1426 | Sub | Private | `GLAudit` |
| `modEngine` | 1433 | Sub | Private | `GLScanCDJ` |
| `modEngine` | 1472 | Sub | Private | `GLScanCRJ` |
| `modEngine` | 1494 | Sub | Private | `GLScanGJ` |
| `modEngine` | 1514 | Function | Private | `GLAccounts` |
| `modEngine` | 1537 | Sub | Private | `Legacy_RefreshGL` |
| `modEngine` | 1618 | Function | Private | `SheetNamesList` |
| `modEngine` | 1632 | Function | Private | `DetectGLYear` |
| `modEngine` | 1679 | Function | Private | `Legacy_CalculateMonthlyNet` |
| `modEngine` | 1700 | Function | Private | `Legacy_CalculateEndingBalanceFull` |
| `modEngine` | 1710 | Function | Private | `Legacy_ValidateGLConsistency` |
| `modEngine` | 1721 | Sub | Private | `Legacy_RefreshAllGL` |
| `modEngine` | 1725 | Sub | Public | `InitializeCOALookup` |
| `modEngine` | 1728 | Function | Public | `GetAccountFromCOA` |
| `modEngine` | 1735 | Function | Public | `GetAccountNameFromTIN` |
| `modEngine` | 1751 | Function | Public | `GetFinancialMonth` |
| `modEngine` | 1755 | Function | Public | `GetFinancialYear` |
| `modEngine` | 1767 | Function | Public | `ValidateGLConsistency` |
| `modEngine` | 1771 | Function | Public | `BuildV8CalcSheet` |
| `modPJSync` | 86 | Function | Private | `PJSheet` |
| `modPJSync` | 104 | Function | Public | `GetNextSequence` |
| `modPJSync` | 142 | Sub | Public | `LookupByTIN_NonVat` |
| `modPJSync` | 187 | Sub | Public | `LookupByName_NonVat` |
| `modPJSync` | 221 | Sub | Public | `ClearSupplierRow` |
| `modPJSync` | 241 | Sub | Public | `CalculateTotals` |
| `modPJSync` | 291 | Sub | Public | `SyncToCDJ_NonVat` |
| `modPJSync` | 475 | Function | Private | `FindExpenseCol` |
| `modGLGate` | 6 | Function | Private | `SheetExists` |
| `modGLGate` | 12 | Function | Public | `ValidateV8Structure` |
| `modGLGate` | 39 | Sub | Public | `RunV8Gate` |
| `modGLRefresh` | 13 | Sub | Public | `RefreshGL` |
| `modGLRefresh` | 17 | Sub | Public | `RefreshAllGL` |
| `modGLRefresh` | 21 | Sub | Public | `RefreshGLIntoSheet` |
| `modGLRefresh` | 25 | Sub | Public | `RefreshGLForYear` |
| `modGLScanEngine` | 20 | Sub | Public | `ScanGL` |
| `modGLScanEngine` | 71 | Sub | Private | `AddPost` |
| `modGLScanEngine` | 82 | Sub | Private | `ScanCDJ` |
| `modGLScanEngine` | 169 | Sub | Private | `ScanCRJ` |
| `modGLScanEngine` | 254 | Function | Private | `ValidateGLAccountTitles` |
| `modGLScanEngine` | 284 | Sub | Private | `WriteScannedGL` |
| `modGLScanEngine` | 311 | Sub | Private | `WriteScanAudit` |
| `modGLScanEngine` | 324 | Function | Private | `SafeText` |
| `modGLScanEngine` | 329 | Function | Private | `MaxD` |
| `modGLWorkbookMap` | 22 | Sub | Public | `V8_MapReset` |
| `modGLWorkbookMap` | 32 | Function | Public | `V8_Key` |
| `modGLWorkbookMap` | 44 | Sub | Private | `EnsureMaps` |
| `modGLWorkbookMap` | 89 | Function | Public | `V8_Norm` |
| `modGLWorkbookMap` | 94 | Function | Public | `V8_Month` |
| `modGLWorkbookMap` | 111 | Function | Public | `V8_GLAccounts` |
| `modGLWorkbookMap` | 122 | Function | Public | `V8_ResolveAccount` |
| `modGLWorkbookMap` | 155 | Function | Public | `V8_CDJMap` |
| `modGLWorkbookMap` | 174 | Function | Public | `GLTitleExists` |
| `modGLWorkbookMap` | 182 | Function | Public | `GLTitleOf` |
| `modGLWorkbookMap` | 193 | Function | Public | `V8_CRJMap` |
| `modGLWorkbookMap` | 203 | Function | Public | `V8_Number` |
| `modGraphify` | 37 | Sub | Implicit | `Graphify` |
| `modGraphify` | 173 | Function | Private | `PickNumericColumn` |
| `modGraphify` | 205 | Function | Private | `FirstNumericRow` |
| `modGraphify` | 225 | Function | Private | `ColumnLetter` |
| `ImportSuppliersFromJSON` | 4 | Sub | Implicit | `ImportSuppliersFromJSON` |

## Pattern inventory

| Pattern | Result | Evidence |
|---|---|---|
| `Option Explicit` | Missing in `Sheet9.txt` and `Sheet10.txt`; present in the other 13 files. | `Sheet9.txt:1`, `Sheet10.txt:1` |
| External file / workbook I/O | None. No `Workbooks.Open`, `Dir`, `Environ`, `FileSystemObject`, `Open #` or `ActiveWorkbook` found. | full-source scan |
| UI selection APIs | No `.Select` / `.Activate`; `modGraphify` uses `ActiveSheet` at `:52`,`:56` and `Selection` at `:66-70` only. | `modGraphify.bas:52-70` |
| Application state changes | `Sheet9` saves and restores prior state; `Sheet11`/`Sheet12` force automatic calculation and true events; `Sheet8` restores true/automatic; `CDJReserveRow` restores only on its normal/failed path; `modGLScanEngine.WriteScannedGL` restores `ScreenUpdating` and re-raises. | `Sheet9.txt:25-27,276-278`; `Sheet11.txt:204-206,259-262`; `Sheet12.txt:201-203,252-255`; `Sheet8.txt:37-40,84-87`; `modEngine.bas:443-463`; `modGLScanEngine.bas:288-306` |
| `On Error` handlers | Empty handlers exist in three public `modEngine` sync procedures, but `On Error GoTo 0` disables them before the write bodies; errors after that point are not swallowed. | `modEngine.bas:847-851,960-961`; `:974-978,1035-1036`; `:1245-1249,1328-1329` |
| Hard-coded sheet names | `"CDJ"`, `"CRJ"`, `"GL"`, `"COA_MAP"`, `"GL_AUDIT"`, `"GJ"`, `"PJ Non-Vat2"`, `"SUPPLIERS DATA"` are referenced literally; no named-range lookup. | source scan |
| Duplicate local declarations | `ValidateV8Structure` declares `n` twice. | `modGLGate.bas:13,29` |
| Undeclared identifiers | `Sheet9` uses `START_ROW`, `COL_TIN`, `COL_NAME`, `COL_AMOUNT`, `COL_DEBIT`, `COL_DEBIT_MIRROR`, `COL_REF`, `COL_COA`, `COL_FULL_ID` without declarations and without `Option Explicit`. | `Sheet9.txt:72,78,135,163,169,187,205,212,226,249,252` |

## Not present in the canonical source

- No named ranges, no `UserForms`, no external data connections, no `OnTime`, no `Auto_Open`.
- No `GL_V8_CALC` writer: `ScanGL` uses `GL` and `GL_AUDIT` only (`modGLScanEngine.bas:55-58`).
- No GJ postings in the active path; `GLScanGJ` exists only in legacy `modEngine` code (`modEngine.bas:1494`).
