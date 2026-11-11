# Audit state

Current phase: **Step 3 complete.**
Branch: `audit/vba-docs-sync` at HEAD `1a7cde6`.
Phases 0-3 and Step 3 are committed and pushed to origin. Canonical `build/v8_5_vba/` is tracked. Held code findings are recorded in `docs/audit/blocked.md`. No live Excel compile or run has been performed.

## Phase status

| Phase | Status | Output |
|---|---|---|
| 0 - Files and provenance | Complete | `00_files.md` |
| 1 - Inventory | Complete | `01_inventory.md` |
| 2 - Claim verification | Complete | `02_claims.md` |
| 3 - Findings | Complete | `03_findings.md` |
| 4 - Documentation fixes | Complete | `VBA_CODEBASE.md` commit `daf0e52`; `VBA_CODE_FLOW.md` commit `0a9d203` |
| 5 - VBA source fixes | Complete | commits C-01 da6c790, C-14 f75bb4f, C-03 97c9c16, C-04 ae2c755, C-07 49ff9dd, C-11 b1c4569, C-12 5fce068, C-13 5a63ce5, C-02 145145e, C-15 56bc67b, C-08 ab1b04d, C-09 26a6da3 |
| 6 - Checker and report | Complete | `scripts/check_docs.py`, `scripts/check_vba_static.py`, `docs/audit/06_report.md` |

## Counts

| Measure | Count |
|---|---:|
| Non-empty components | 15 |
| Declared procedures | 146 |
| Event handlers | 7 |
| Material claims checked | 48 |
| Claims verified | 33 |
| Claims partial | 7 |
| Claims wrong | 8 |
| Findings | 28 |
| Code findings | 15 |
| Documentation findings | 13 |
| Critical | 1 |
| High | 4 |
| Medium | 13 |
| Low | 9 |
| Info | 1 |

## Open decisions

Approved behavioural fixes: C-02, C-03, C-04, C-08, C-09. Held: C-05, C-06, C-10. The remaining code findings are mechanical fixes. Documentation findings D-01 through D-13 are fixed.

## Evidence boundary

- Canonical source is `build/v8_5_vba/`; the workbook hash is verified. The export is tracked in `dc9dbfa`.
- Privacy scan: 96 tax-ID-like spans across 95 lines; see `privacy_note.md` (no values recorded).
- `build/artifacts/compile.log` is stale and is not accepted as evidence that the current export compiles.
- No live Excel compile or run has been performed. The open workbook/embedded-source boundary remains unverified.
- `DRR_deep-research-report.md` remains an untracked hypothesis list, not a finding source.
