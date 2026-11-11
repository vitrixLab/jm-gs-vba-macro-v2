# Blocked code findings

Held per decision. Do not edit the listed logic until the question is answered.

## C-05 - Sheet9 undeclared constants (held)
Question: What are the confirmed column values for Sheet9's undeclared constants (`START_ROW`, `COL_TIN`, `COL_NAME`, `COL_AMOUNT`, `COL_DEBIT`, `COL_DEBIT_MIRROR`, `COL_REF`, `COL_COA`, `COL_FULL_ID`)? Do not guess; confirm from the live Sheet9 layout or operator before declaring constants and adding `Option Explicit`.

## C-06 - SyncPJNonVatToCDJ missing CDJ day (held)
Question: What is the source column for the day value on Sheet11/PJ Non-Vat rows (likely column C), and should a blank day be cleared on update? Confirm before writing `CDJ_COL_DAY` on both update and insert paths.

## C-10 - LookupByTIN_NonVat partial write (held)
Question: On a duplicate TIN shared by different supplier names in `LookupByTIN_NonVat`, should any partial supplier row already written be rolled back/cleared, or should the lookup be aborted before any sheet writes? Confirm rollback/abort behavior before changing `modPJSync.bas:165-181`.
