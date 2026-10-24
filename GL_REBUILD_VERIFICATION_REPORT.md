# GL Automation v2 — Evidence Status

**Status: HOLD — not certified.**

## Scope boundary

- CDJ and CRJ are the only permitted posting inputs.
- GJ must not be read or posted.
- Write-back target is the existing GL sheet.
- Do not create GL_V8_CALC or restore deleted V*GL sheets.
- The intended output layout is 46 distinct account blocks × 12 months.

## Source-level findings

1. The previous scan module contained a ScanGJ procedure that read the GJ worksheet. The proposed source removes that procedure; the active scan path calls CDJ and CRJ only.
2. The scan engine previously allowed repeated GL account titles to receive the same normalized account/month postings in multiple blocks. The proposed source validates all 46 account-title rows and holds write-back on blank or duplicate titles.
3. The mapping update script in the prior baseline hard-coded ten COA mappings and marked them MAPPED without retaining accounting approval evidence. The proposed change neutralizes that script; no such mapping is represented here as approved.
4. The source calls an account resolver for CDJ sundry labels. Resolver output is not proof of accounting authorization; only operator-approved COA_MAP rows may be used.

## Not verified

- Embedded VBA/source equivalence for the candidate workbook.
- Successful VBA compile or Excel Desktop/COM execution.
- Source-row completeness and duplicate-transaction resolution.
- Twelve-month CDJ/CRJ totals, account-level roll-forward, and ending balances.
- Whether all 46 account blocks in the actual workbook are distinct and correctly titled.

## Acceptance gates

| Gate | Required evidence | Current status |
|---|---|---|
| Scope | No GJ worksheet read or post in active scan | Source change proposed; workbook binary not re-injected |
| GL structure | 46 nonblank, unique account titles | Guard added; live workbook not runtime-tested |
| Mapping | Every nonzero sundry label has explicit approved mapping | HOLD pending accounting evidence |
| Reconciliation | Source journal totals and GL postings agree within tolerance | NOT RUN |
| Roll-forward | Each month’s ending balance equals prior ending + debit − credit | NOT RUN |
| Runtime | Compile and refresh in Excel, inspect output and audit row | NOT RUN |

No monthly ending balances or overall PASS claim are certified by this report.
