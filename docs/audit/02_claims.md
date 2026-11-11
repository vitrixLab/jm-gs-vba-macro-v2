# Phase 2 - Claim verification

Material claims from `VBA_CODEBASE.md` and `VBA_CODE_FLOW.md`, checked against `build/v8_5_vba/`. The procedure index and the 15 full-source reproductions were verified in bulk and are treated as one claim each.

Status counts: **33 VERIFIED, 7 PARTIAL, 8 WRONG**.

## `VBA_CODEBASE.md`

| ID | Doc claim | Evidence | Status |
|---|---|---|---|
| C01 | Workbook SHA-256 is `D35BF...` (`:5`) | Recomputed SHA-256 matches; also in `build/verify_summary.json` | VERIFIED |
| C02 | 15 non-empty components and 146 declared procedures (`:24`) | Inventory: 8 standard + 6 sheet + `ThisWorkbook`; 146 procedures | VERIFIED |
| C03 | Per-module line/nonblank/procedure counts (`30-56`) | Recomputed; all 15 rows match | VERIFIED |
| C04 | Full procedure index with line numbers (`:101-...`) | All 146 entries match declaration lines in source | VERIFIED |
| C05 | Full source reproduced for every non-empty module | All 15 code fences match the canonical files line-for-line | VERIFIED |
| C06 | `vba/` is byte-identical except the `Attribute VB_Name` header (`:24`, `:5770`) | Also differs by UTF-8 BOM and 1-2 trailing blank lines; body logic is otherwise the same | WRONG |
| C07 | Empty-module sentence lists all empty sheet modules (`:73`) | Sentence omits `CDJ` and `GJ`, although both appear in the table above it | PARTIAL |
| C08 | v8.5 has a recorded compile PASS (`:99`, `:5779`) | `compile.log` exists, but mtime 17:42 predates `modGLGate.bas` 18:28 and workbook 18:25; not evidence for current text | PARTIAL |
| C09 | `Sheet9` undeclared constants are the listed subset (`:81`) | Also uses undeclared `START_ROW` and `COL_DEBIT_MIRROR`; `Sheet9.txt:72,169,226,249` | PARTIAL |
| C10 | COA `MAPPED` test accepts `UNMAPPED` (`:82`) | `InStr(st, "MAPPED") > 0` at `modGLWorkbookMap.bas:75`; `UNMAPPED` contains the substring | VERIFIED |
| C11 | Hard-coded HOLD aliases are executable (`:83`) | `modGLWorkbookMap.bas:139-143` maps five extra labels, including `COST OF REVENUE`, `SUPPLIES`, `BANK CHARGE`, `CHARGES` | VERIFIED |
| C12 | Scan year is not enforced (`:84`) | `ScanCDJ`/`ScanCRJ` accept `yr` but never use it; `modGLScanEngine.bas:82,169` | VERIFIED |
| C13 | GL ending balance resets per block (`:85`) | `bal = 0#` inside the block loop at `modGLScanEngine.bas:293` | VERIFIED |
| C14 | Gate is before GL write-back (`:89`) | `ScanGL` writes audit and exits before `WriteScannedGL` on gate failure; `modGLScanEngine.bas:42-68` | VERIFIED |
| C15 | Titles validated twice, including duplicates (`:90`) | `ValidateV8Structure` + `ValidateGLAccountTitles`; `modGLGate.bas:12-37`, `modGLScanEngine.bas:254-282` | VERIFIED |
| C16 | Sundry T/U disagreements are audit-only (`:91`) | Numeric mismatch only adds detail (`:157-162`), but nonnumeric T/U increments `inv` (`:153,156`) and can force HOLD | PARTIAL |
| C17 | Sync errors are swallowed by empty ErrorHandler (`:92`) | `On Error GoTo 0` at `modEngine.bas:851,978,1249` disables the handler before writes; later write errors propagate | WRONG |
| C18 | Supplier import clears before validation (`:93`) | Rows 2+ cleared at `ImportSuppliersFromJSON.bas:13`; duplicates checked after writes at `:816-847` | VERIFIED |
| C19 | Legacy GJ/calc code remains compiled (`:97`) | All listed procedures exist at `modEngine.bas:1494,1537,1679,1710,1721,1771` | VERIFIED |
| C20 | Duplicate implementations exist (`:98`) | `Sheet11`/`Sheet12` and `Sheet3`/`Sheet8` duplicate logic; `Sheet9` duplicates `modPJSync` | VERIFIED |
| C21 | Snapshot cannot certify compile/run (`:99`) | Correct boundary; no workbook was opened or compiled in this audit | VERIFIED |
| C22 | Active path is PJ/PNV->CDJ, SJ->CRJ, gate->GL (`flow:15-23`) | Call graph and write targets match | VERIFIED |
| C23 | Entry-point catalogue (`flow:29-43`) | Mostly correct, but SJ2 is not the “same sales formulas” and Sheet9 remains wired; rows 35 and 33 are partial as written | PARTIAL |
| C24 | VAT vs Non-VAT differences (`flow:179-189`) | Trigger columns, amount treatment, `PJ-`/`PNV-` prefixes, sync targets, EWT and supplier-add all match | VERIFIED |
| C25 | CDJ placement inside month partition (`flow:191-218`) | `FindCDJFreeRow` scans the month block; `CDJReserveRow` inserts above TOTAL; both syncs use it | VERIFIED |
| C26 | Both syncs write “date, day, particulars…” (`flow:208`) | `SyncPJToCDJ` writes day at `modEngine.bas:884,926`; `SyncPJNonVatToCDJ` never writes `CDJ_COL_DAY` (`:1267-1286,1299-1319`) | WRONG |
| C27 | Legacy lookup order AA/AB/Y (`flow:220-233`) | `LoadCDJReferences` populates those keys and `FindCDJRow` checks them in order | VERIFIED |
| C28 | Active non-VAT path is Sheet11 (`flow:235`) | `Sheet9.Worksheet_Change` remains wired and callable; it is broken, not disconnected | PARTIAL |
| C29 | SJ invoice stamp requires date + amount (`flow:248`, `:1095`) | `Sheet3.txt:62-63,78` checks amount/TOTAL only; `Sheet8.txt:67-69` checks date too | WRONG |
| C30 | SJ updates existing CRJ row or appends and advances pointer (`flow:257-265`) | Matches `SyncSJToCRJ` at `modEngine.bas:966-1039` | VERIFIED |
| C31 | SJ2 uses “same sales formulas” (`flow:35`) | `Sheet8.txt:58-64` writes only M/O; `Sheet3.txt:72-75` also writes L/N mirrors | WRONG |
| C32 | SJ/SJ2 layout table (`flow:829-846`) | Conflates layouts: Sheet8 column I is auto invoice and U is reference; Sheet3 I is manual physical invoice, U auto invoice, V full ID | WRONG |
| C33 | GL entry points and gate sequence (`flow:296-329`) | Wrappers call `ScanGL`; audit on HOLD before `WriteScannedGL`; audit PASS after write | VERIFIED |
| C34 | CDJ scan details: allocation, Sundry, T/U correction (`flow:340-420`) | `ScanCDJ` derives Sundry from balanced row and records T/U correction; numeric mismatch is audit-only | VERIFIED |
| C35 | CRJ scan details: invoice boundaries, 12 blocks, invalid partitions (`flow:430-455`) | Matches `ScanCRJ` at `modGLScanEngine.bas:169-252` | VERIFIED |
| C36 | GL layout mapping F:I and six value columns (`flow:455-483`) | `WriteScannedGL` writes those columns and `GL_AUDIT` rows | VERIFIED |
| C37 | Account resolution: aliases, exact, punctuation-normalized (`flow:122-153`) | `V8_ResolveAccount` order matches; alias table wins over built-ins | VERIFIED |
| C38 | Fixed CDJ/CRJ column maps (`flow:493-495`) | `V8_CDJMap` also maps column 6; `V8_CRJMap` also maps column 8; alias list omits four labels | PARTIAL |
| C39 | Supplier lookup ambiguity and cache behavior (`flow:499-527`) | `GetSupplierByTIN/Name` returns nothing on ambiguous TIN; add refreshes cache | VERIFIED |
| C40 | Cross-module dependency matrix (`flow:744-760`) | Direct calls match for all listed rows | VERIFIED |
| C41 | Sync write error is swallowed and diverges silently (`flow:895-899`) | Contradicted by `On Error GoTo 0` before write bodies; errors surface | WRONG |
| C42 | B.3 invoice stamp requires date + amount (`flow:1095`) | Same Sheet3 condition error as C29 | WRONG |
| C43 | Initialization facts (`flow:122-127`) | Idempotent rebuild, CRJ early return, ignored AA14 pointer, lazy CRJ caches all match | VERIFIED |
| C44 | Audit is written for PASS and HOLD (`flow:1067-1071`) | `ScanGL` writes `GL_AUDIT` on both paths | VERIFIED |
| C45 | `vba/` has the same logic as the v8.5 export (`:5770`) | Body logic matches after header/BOM/trailing-blank normalization | VERIFIED |
| C46 | `verify_summary.json` lists changed modules and workbook hash (`:5778`) | File lists `Sheet11`, `Sheet12`, `modEngine`; hash matches | VERIFIED |
| C47 | No external I/O / named ranges / UserForms | Full-source pattern scan found none | VERIFIED |
| C48 | `Sheet9` has no `Option Explicit` (`:81`) | Confirmed at `Sheet9.txt:1` | VERIFIED |

## Claims intentionally not rowed

The following were checked in bulk rather than as individual rows: every procedure-index line, every full-source code block, all 22 Mermaid blocks for syntactic balance, and every line/nonblank/procedure count in Section 3. No other material contradiction was found in the code reproductions themselves.
