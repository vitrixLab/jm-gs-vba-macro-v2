$ErrorActionPreference = "Stop"

$xlFile = "D:\citrixlabph\globalsmile_v2\Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening $xlFile..."
    $wb = $excel.Workbooks.Open($xlFile)

    # --- 1. Update COA_MAP rows 13 to 22 directly ---
    $wsMap = $wb.Sheets.Item("COA_MAP")
    $mapUpdates = @{
        13 = @("Representation", "Meals, Gifts & Entertainment Expense")
        14 = @("Partners Clinicians Fee", "Clinician's Fee-Corporators")
        15 = @("Clinicians Fee- Corporatos", "Clinician's Fee-Corporators")
        16 = @("Visiting Clinicians Fee", "Clinician's Fee-Visiting")
        17 = @("SSS,HDMF AND PHIC ER SHARE", "Government Contributions ER Share")
        18 = @("Clinacians Fee- Corporators", "Clinician's Fee-Corporators")
        19 = @("Partners Clinicians Fee- Corporators", "Clinician's Fee-Corporators")
        20 = @("Incentives", "De Minimis")
        21 = @("Outside Services", "Common Use Service Area")
        22 = @("Training and Seminar", "Miscellaneous")
    }

    foreach ($r in $mapUpdates.Keys) {
        $lbl = $mapUpdates[$r][0]
        $acct = $mapUpdates[$r][1]
        $wsMap.Cells.Item($r, 1).Value2 = $lbl
        $wsMap.Cells.Item($r, 4).Value2 = $acct
        $wsMap.Cells.Item($r, 6).Value2 = "MAPPED"
        Write-Host "COA_MAP row $r updated: '$lbl' -> '$acct' | MAPPED"
    }

    # Clear rows 25 to 45
    for ($r = 25; $r -le 45; $r++) {
        for ($c = 1; $c -le 8; $c++) {
            $wsMap.Cells.Item($r, $c).Value2 = $null
        }
    }
    Write-Host "Cleared extra rows 25-45 in COA_MAP."

    # --- 2. Update modGLWorkbookMap EnsureMaps to use For r = 5 To lastRow ---
    $comp = $wb.VBProject.VBComponents.Item("modGLWorkbookMap")
    $cm = $comp.CodeModule
    $totalLines = $cm.CountOfLines
    Write-Host "modGLWorkbookMap line count: $totalLines"

    for ($i = 1; $i -le $totalLines; $i++) {
        $line = $cm.Lines($i, 1)
        if ($line -match "^\s*Do While Len\(Trim\$\(CStr\(ws\.Cells\(r,\s*1\)\.Value2\)\)\) > 0") {
            Write-Host ('Found Do While at line ' + $i + ': ' + $line)
            $cm.ReplaceLine($i, "        For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row")
            Write-Host "  Replaced with For loop."
        }
        if ($line -match "^\s*r = r \+ 1\s*$") {
            # In the original: r = r + 1 inside Do While
            $nextLine = $cm.Lines($i + 1, 1)
            if ($nextLine -match "^\s*Loop\s*$") {
                Write-Host ('Found r = r + 1 at line ' + $i)
                $cm.ReplaceLine($i, "            ' (increment handled by For loop)")
                $cm.ReplaceLine($i + 1, "        Next r")
                Write-Host "  Replaced Loop with Next r."
            }
        }
    }

    # --- 3. Save workbook before running RefreshGL ---
    $wb.Save()
    Write-Host "Saved workbook changes."

    # --- 4. Run RefreshGL ---
    Write-Host "`nRunning modGLRefresh.RefreshGL..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("modGLRefresh.RefreshGL")
    $sw.Stop()
    Write-Host ("RefreshGL finished in " + $sw.ElapsedMilliseconds + " ms")

    # --- 5. Save again to persist calculated GL ---
    $wb.Save()
    Write-Host "Saved workbook after RefreshGL."

    # --- 6. Inspect GL_AUDIT last row ---
    $auditWs = $wb.Sheets.Item("GL_AUDIT")
    $lastAuditRow = $auditWs.Cells($auditWs.Rows.Count, 1).End(-4162).Row
    Write-Host "`n=== GL_AUDIT (Last Row $lastAuditRow) ==="
    Write-Host ("Timestamp: " + $auditWs.Cells.Item($lastAuditRow, 1).Text)
    Write-Host ("Status:    " + $auditWs.Cells.Item($lastAuditRow, 2).Text)
    Write-Host ("Summary:   " + $auditWs.Cells.Item($lastAuditRow, 3).Text)

    # --- 7. Inspect GL sheet Block 1 (Cash in Bank) ---
    $glWs = $wb.Sheets.Item("GL")
    Write-Host "`n=== GL Sheet (Cash in Bank, Rows 13-24) ==="
    for ($r = 13; $r -le 24; $r++) {
        $m = $glWs.Cells.Item($r, 2).Text
        $title = $glWs.Cells.Item($r, 6).Text
        $dr = $glWs.Cells.Item($r, 7).Text
        $cr = $glWs.Cells.Item($r, 8).Text
        $bal = $glWs.Cells.Item($r, 9).Text
        Write-Host ("Row " + $r + " | Month: " + $m + " | Dr: " + $dr + " | Cr: " + $cr + " | Bal: " + $bal)
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
