# Global-Smile 2026 — GL Automation v2

## Current status: HOLD

This repository contains VBA source and a candidate workbook, but the current evidence does **not** justify calling the workbook certified or production-ready. A green audit row, a source-level test, or a balanced grand total alone is not sufficient proof.

## Required scope

- **Inputs:** CDJ and CRJ only.
- **Excluded:** GJ must not be read or posted.
- **Output:** existing GL worksheet only.
- **No generated calculation sheets:** do not recreate GL_V8_CALC or any deleted V*GL sheet.
- **Coverage:** 46 distinct account titles × 12 months, subject to validation against the actual GL account blocks.
- **Roll-forward:** ending balance = prior ending balance + current debit − current credit.

## Fail-closed behavior

The scan must stop before GL write-back when any of these conditions occur:

1. An account title is blank or duplicated across the 46 GL blocks.
2. A CDJ/CRJ record is invalid or cannot be assigned to a verified month.
3. A nonzero posting amount has no approved account mapping.
4. Debit and credit totals exceed the configured tolerance.

No accounting mapping should be marked approved solely to make a gate pass. The ten sundry labels in the supplied mapping script are **not assumed approved** by this repository; they require explicit accounting authorization.

## Verification boundary

The current PR adds a source-level guard for duplicate/missing GL account titles and removes the GJ scan procedure. It does **not** certify that the embedded VBA in either XLSM workbook matches the exported source, compiles in Excel, or passes a live refresh. Those require workbook-level verification in Excel Desktop/COM.

Run source-level checks:

    python -m unittest discover -s tests -v

Do not run scripts/apply_full_fix.ps1 from earlier commits: it writes accounting mappings into COA_MAP without attached approval evidence. The file is neutralized by this change to prevent accidental use.
