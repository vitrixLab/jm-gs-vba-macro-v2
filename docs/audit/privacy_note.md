# Privacy scan - tax-ID-like patterns

Scan date: 2026-11-11
Scope: `build/v8_5_vba/` text files only.
Pattern: `\d{3}-?\d{3}-?\d{3}-?\d{0,4}`
No matched values are reproduced in this note.

## Count

**96 spans across 95 lines.**

| File | Matching lines | Spans |
|---|---:|---:|
| `build/v8_5_vba/ImportSuppliersFromJSON.bas` | 89 | 89 |
| `build/v8_5_vba/modEngine.bas` | 2 | 2 |
| `build/v8_5_vba/modPJSync.bas` | 1 | 2 |
| `build/v8_5_vba/Sheet3.txt` | 1 | 1 |
| `build/v8_5_vba/Sheet8.txt` | 2 | 2 |

## File:line evidence

`build/v8_5_vba/ImportSuppliersFromJSON.bas` (one span each):
23, 31, 39, 47, 55, 63, 71, 79, 87, 95, 103, 111, 119, 127, 135, 143, 151, 159, 167, 175, 183, 191, 199, 207, 215, 223, 231, 239, 247, 263, 271, 279, 287, 295, 303, 311, 319, 327, 335, 343, 351, 359, 367, 375, 383, 391, 399, 407, 415, 423, 431, 439, 447, 455, 463, 471, 479, 487, 495, 503, 511, 519, 527, 535, 543, 551, 559, 567, 575, 583, 591, 599, 607, 615, 623, 631, 639, 647, 655, 663, 671, 679, 687, 695, 705, 713, 721, 729, 737

`build/v8_5_vba/modEngine.bas` (one span each):
246, 247

`build/v8_5_vba/modPJSync.bas` (two spans):
157

`build/v8_5_vba/Sheet3.txt` (one span):
72

`build/v8_5_vba/Sheet8.txt` (one span each):
57, 60

## Interpretation

The supplier seeder accounts for 89 of the 96 spans. The remaining matches are in comments or numeric formula/constant text; the broad pattern can match number fragments in formulas. The scan is a pattern inventory, not proof that every span is a real tax ID.

No values were printed, and no files were edited as part of this scan.
