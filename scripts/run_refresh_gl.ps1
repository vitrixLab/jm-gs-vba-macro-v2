$ErrorActionPreference = "Stop"

$xlFile = "D:\citrixlabph\globalsmile_v2\Global-Smile_2026-v8.4-NO-SHIT-FIXED.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening $xlFile..."
    $wb = $excel.Workbooks.Open($xlFile)
    Write-Host "Running modGLRefresh.RefreshGL..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("modGLRefresh.RefreshGL")
    $sw.Stop()
    Write-Host ("RefreshGL finished in " + $sw.ElapsedMilliseconds + " ms")

    Write-Host "Saving workbook..."
    $wb.Save()
    Write-Host "Saved successfully."

    # Inspect GL_AUDIT last row
    $auditWs = $wb.Sheets.Item("GL_AUDIT")
    $lastAuditRow = $auditWs.Cells($auditWs.Rows.Count, 1).End(-4162).Row
    Write-Host "`n=== GL_AUDIT (Last Row $lastAuditRow) ==="
    Write-Host ("Timestamp: " + $auditWs.Cells.Item($lastAuditRow, 1).Text)
    Write-Host ("Status:    " + $auditWs.Cells.Item($lastAuditRow, 2).Text)
    Write-Host ("Summary:   " + $auditWs.Cells.Item($lastAuditRow, 3).Text)

    # Inspect GL sheet Block 1 (Cash in Bank)
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
