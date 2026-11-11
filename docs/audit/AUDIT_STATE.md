# Audit state

Current phase: **Step 0 complete; Phase 4 documentation fixes next.**
Branch: `audit/vba-docs-sync` at HEAD `dc9dbfa`.
Phase 3 is committed and published. Canonical `build/v8_5_vba/` is tracked. The privacy scan is recorded. No VBA source or workbook was modified.

## Phase status

| Phase | Status | Output |
|---|---|---|
| 0 - Files and provenance | Complete | `00_files.md` |
| 1 - Inventory | Complete | `01_inventory.md` |
| 2 - Claim verification | Complete | `02_claims.md` |
| 3 - Findings | Complete | `03_findings.md` |
| 4 - Documentation fixes | Complete | `VBA_CODEBASE.md` commit `daf0e52`; `VBA_CODE_FLOW.md` commit `0a9d203` |
| 5 - VBA source fixes | Not started | `build/v8_5_vba/` |
| 6 - Checker and report | Not started | `scripts/`, `docs/audit/06_report.md` |

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
| Findings | 27 |
| Code findings | 14 |
| Documentation findings | 13 |
| Critical | 1 |
| High | 4 |
| Medium | 12 |
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
