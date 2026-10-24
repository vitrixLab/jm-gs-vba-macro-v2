import openpyxl

wb = openpyxl.load_workbook(r'D:\citrixlabph\globalsmile_v2\Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm', data_only=True)
cdj = wb['CDJ']
crj = wb['CRJ']
gl = wb['GL']

MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

def v8_norm(v):
    if v is None: return ""
    return str(v).replace('\u00a0', ' ').strip().upper()

def v8_month(v):
    norm = v8_norm(v)
    m_map = {
        'JAN': 1, 'JANUARY': 1,
        'FEB': 2, 'FEBRUARY': 2,
        'MAR': 3, 'MARCH': 3,
        'APR': 4, 'APRIL': 4,
        'MAY': 5, 'MAY': 5,
        'JUN': 6, 'JUNE': 6,
        'JUL': 7, 'JULY': 7,
        'AUG': 8, 'AUGUST': 8,
        'SEP': 9, 'SEPTEMBER': 9,
        'OCT': 10, 'OCTOBER': 10,
        'NOV': 11, 'NOVEMBER': 11,
        'DEC': 12, 'DECEMBER': 12
    }
    return m_map.get(norm, 0)

# 1. CDJ Source Inspection: Check visible TOTAL rows
cdj_total_rows = {}
for r in range(15, 700):
    c5 = str(cdj.cell(row=r, column=5).value or "").strip().upper()
    if "TOTAL" in c5:
        # Get month from col C or preceding
        m_val = cdj.cell(row=r, column=3).value
        cash_tot = float(cdj.cell(row=r, column=6).value or 0.0)
        cdj_total_rows[r] = (c5, m_val, cash_tot)

print("=== CDJ VISIBLE TOTAL ROWS ===")
for r, (c5, m_val, cash_tot) in cdj_total_rows.items():
    print(f"Row {r:3d}: '{c5}' (Col C='{m_val}') Cash Total = {cash_tot:>14,.2f}")

# 2. CDJ Data Rows Sum vs Visible TOTAL rows
cdj_month_cash = [0.0] * 13
cdj_month_alloc_db = [0.0] * 13
cdj_month_alloc_cr = [0.0] * 13
cdj_month_sundry_net = [0.0] * 13
cur_m = 0

for r in range(15, 675):
    c3 = cdj.cell(row=r, column=3).value
    m = v8_month(c3)
    if m > 0: cur_m = m
    if cur_m == 0: continue
    
    title = str(cdj.cell(row=r, column=5).value or "").strip()
    if len(title) == 0 or "TOTAL" in title.upper():
        continue
    
    cash = float(cdj.cell(row=r, column=6).value or 0.0)
    cdj_month_cash[cur_m] += cash
    
    # Allocations
    for c in range(7, 19):
        val = float(cdj.cell(row=r, column=c).value or 0.0)
        if val > 0: cdj_month_alloc_db[cur_m] += val
        elif val < 0: cdj_month_alloc_cr[cur_m] += (-val)
        
    # Sundry
    raw = str(cdj.cell(row=r, column=19).value or "").strip()
    if len(raw) > 0 and raw.upper() not in ["TOTAL", "SUNDRY ACCOUNT"]:
        sundry_db = float(cdj.cell(row=r, column=20).value or 0.0)
        sundry_cr = float(cdj.cell(row=r, column=21).value or 0.0)
        cdj_month_sundry_net[cur_m] += (sundry_db - sundry_cr)

print("\n=== CDJ MONTHLY TRANSACTION TOTALS (SUM OF DATA ROWS) ===")
for m in range(1, 7):
    print(f"Month {m} ({MONTH_NAMES[m-1]}): Cash (Col F) = {cdj_month_cash[m]:>14,.2f} | Allocations Dr = {cdj_month_alloc_db[m]:>14,.2f} | Sundry Net = {cdj_month_sundry_net[m]:>14,.2f}")

# 3. CRJ Monthly Totals
crj_month_cash = [0.0] * 13
crj_month_vat = [0.0] * 13
crj_month_exempt = [0.0] * 13
crj_month_sales = [0.0] * 13

startRow = 10
lastMonth = 0
for r in range(10, 583):
    m = v8_month(crj.cell(row=r, column=4).value)
    f_val = str(crj.cell(row=r, column=6).value or "").strip()
    g_val = str(crj.cell(row=r, column=7).value or "").strip()
    if m > 0 and len(f_val) == 0 and len(g_val) == 0:
        for rr in range(startRow, r):
            if str(crj.cell(row=rr, column=7).value or "").strip():
                crj_month_cash[m] += float(crj.cell(row=rr, column=8).value or 0.0)
                crj_month_vat[m] += float(crj.cell(row=rr, column=10).value or 0.0)
                crj_month_exempt[m] += float(crj.cell(row=rr, column=11).value or 0.0)
                crj_month_sales[m] += float(crj.cell(row=rr, column=12).value or 0.0)
        startRow = r + 1

print("\n=== CRJ MONTHLY TRANSACTION TOTALS ===")
for m in range(1, 7):
    cr_sum = crj_month_vat[m] + crj_month_exempt[m] + crj_month_sales[m]
    print(f"Month {m} ({MONTH_NAMES[m-1]}): Cash Dr (Col H) = {crj_month_cash[m]:>14,.2f} | Credits (Vat+Sales) = {cr_sum:>14,.2f} | Diff = {crj_month_cash[m] - cr_sum:>8.2f}")
