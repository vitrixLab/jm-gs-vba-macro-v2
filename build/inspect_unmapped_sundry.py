"""
inspect_unmapped_sundry.py
Reads the CDJ's Sundry Account column (col S = col 19) from the applied workbook,
shows every unique label and whether V8_ResolveAccount / GL title would match it.
"""
import os, sys
import win32com.client

ROOT   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WB     = os.path.join(ROOT, "Global-Smile_2026-v8.5-applied.xlsm")

# Hard-coded aliases from modGLWorkbookMap.V8_ResolveAccount
ALIASES = {
    "TRANSPORATION AND TRAVEL":  "Transportation and Travel",
    "TRANSPORTATION AND TRAVEL":  "Transportation and Travel",
    "ADVANCES TO EMPLOYEE":      "Advances to Employees",
    "REPAIRS AND MAINTENANCE":   "Repair and Maintenance",
    "OUTPUT VAT PAYABLE":        "VAT Payable",
    "CLINIC SUPPLIES":           "Clinic Material and Supplies",
    "MEDICAL EQUIPMENT":         "Dental Equipment",
    "COST OF REVENUE":           "Clinic Material and Supplies",
    "SUPPLIES":                  "Clinic Material and Supplies",
    "BANK CHARGE":               "Miscellaneous",
    "CHARGES":                   "Miscellaneous",
}

def norm(v):
    import re
    return re.sub(r"\s+", " ", str(v).upper().replace("\u00a0", " ")).strip()

def v8_key(v):
    return "".join(c for c in str(v).upper() if c.isalnum())

excel = win32com.client.DispatchEx("Excel.Application")
excel.AutomationSecurity = 3
excel.DisplayAlerts = False
excel.Visible = False
excel.EnableEvents = False

try:
    wb   = excel.Workbooks.Open(WB, ReadOnly=True, UpdateLinks=0)
    cdj  = wb.Worksheets("CDJ")
    gl   = wb.Worksheets("GL")

    # Build GL title set
    gl_norm = set()
    gl_key  = set()
    last_gl = gl.Cells(gl.Rows.Count, 6).End(-4162).Row  # xlUp
    for r in range(13, last_gl + 1):
        t = str(gl.Cells(r, 6).Value2 or "").strip()
        if t:
            gl_norm.add(norm(t))
            gl_key.add(v8_key(t))

    # Scan CDJ col 19 (S) = Sundry Account
    last_cdj = cdj.Cells(cdj.Rows.Count, 5).End(-4162).Row
    labels   = {}   # norm_label -> (original, cash_total, rows)

    for r in range(15, last_cdj + 1):
        raw = str(cdj.Cells(r, 19).Value2 or "").strip()
        if not raw:
            continue
        n = norm(raw)
        if n in ("TOTAL", "SUNDRY ACCOUNT", ""):
            continue
        try:
            cash = float(cdj.Cells(r, 6).Value2 or 0)
        except Exception:
            cash = 0.0
        if n not in labels:
            labels[n] = [raw, 0.0, 0]
        labels[n][1] += abs(cash)
        labels[n][2] += 1

    wb.Close(False)
finally:
    try:
        excel.Quit()
    except Exception:
        pass

print(f"\n{'='*72}")
print(f"CDJ Sundry Account labels found ({len(labels)} unique)")
print(f"{'='*72}")
print(f"{'Label':<45} {'Rows':>4} {'Cash':>14} {'Status'}")
print(f"{'-'*72}")

unmapped = []
for n, (orig, cash, rows) in sorted(labels.items(), key=lambda x: -x[1][1]):
    if n in ALIASES:
        status = f"-> {ALIASES[n]}"
    elif n in gl_norm:
        status = "GL exact match"
    elif v8_key(n) in gl_key:
        status = "GL key match"
    else:
        status = "*** UNMAPPED ***"
        unmapped.append((orig, cash, rows))
    print(f"{orig:<45} {rows:>4} {cash:>14,.2f}  {status}")

print(f"\n{'='*72}")
print(f"UNMAPPED labels ({len(unmapped)}) — these are blocking WriteScannedGL:")
print(f"{'='*72}")
total_um = 0
for orig, cash, rows in sorted(unmapped, key=lambda x: -x[1]):
    print(f"  {orig:<45} ₱{cash:>14,.2f}  ({rows} rows)")
    total_um += cash
print(f"  {'TOTAL':45} ₱{total_um:>14,.2f}")

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unmapped_sundry.txt")
with open(out, "w", encoding="utf-8") as f:
    f.write(f"UNMAPPED SUNDRY LABELS\n")
    for orig, cash, rows in sorted(unmapped, key=lambda x: -x[1]):
        f.write(f"{orig}\t{cash:.2f}\t{rows}\n")
print(f"\nSaved → {out}")
