$ErrorActionPreference = "Stop"

$root = "D:\citrixlabph\globalsmile_v2"
$src = Join-Path $root "Global-Smile_2026-v8.4-NO-SHIT.xlsm"
$dest = Join-Path $root "Global-Smile_2026-v8.4.1.xlsm"

Write-Host "=== Rebuilding Global-Smile 2026 v8.4.1 ==="
Write-Host "Source:      $src"
Write-Host "Destination: $dest"

Copy-Item -Path $src -Destination $dest -Force
Write-Host "Workbook copied successfully."

$modules = @("modGLScanEngine", "modGLWorkbookMap", "modGLGate", "modGLRefresh")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "`nOpening $dest..."
    $wb = $excel.Workbooks.Open($dest)
    Write-Host "Workbook opened."

    foreach ($m in $modules) {
        $basPath = Join-Path $root "vba\$m.bas"
        if (-not (Test-Path $basPath)) {
            throw "Missing module file: $basPath"
        }
        $rawCode = [System.IO.File]::ReadAllText($basPath, [System.Text.Encoding]::UTF8)
        # Strip UTF-8 BOM if present
        $cleanCode = $rawCode.TrimStart([char]0xFEFF)

        $comp = $wb.VBProject.VBComponents.Item($m)
        $cm = $comp.CodeModule
        $lineCount = $cm.CountOfLines
        if ($lineCount -gt 0) {
            $cm.DeleteLines(1, $lineCount)
        }
        $cm.AddFromString($cleanCode)
        Write-Host "Injected $m : $($cm.CountOfLines) lines (from $basPath)"
    }

    Write-Host "`nSaving rebuilt v8.4.1 workbook..."
    $wb.Save()
    Write-Host "Saved successfully."

    # Test running RefreshGL
    Write-Host "`nRunning modGLRefresh.RefreshGL in live workbook..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("modGLRefresh.RefreshGL")
    $sw.Stop()
    Write-Host ("RefreshGL finished in " + $sw.ElapsedMilliseconds + " ms")

    $wb.Save()

    # Read latest GL_AUDIT row
    $wsAudit = $wb.Sheets.Item("GL_AUDIT")
    $lastRow = $wsAudit.Cells($wsAudit.Rows.Count, 1).End(-4162).Row # xlUp
    Write-Host "`n=== Latest GL_AUDIT Entry (Row $lastRow) ==="
    Write-Host ("Timestamp: " + $wsAudit.Cells.Item($lastRow, 1).Text)
    Write-Host ("Status:    " + $wsAudit.Cells.Item($lastRow, 2).Text)
    Write-Host ("Summary:   " + $wsAudit.Cells.Item($lastRow, 3).Text)
    $detail = $wsAudit.Cells.Item($lastRow, 4).Text
    if ($detail.Length -gt 0) {
        Write-Host "Details preview (first 500 chars):"
        Write-Host ($detail.Substring(0, [Math]::Min(500, $detail.Length)))
    }

    $wb.Close($false)
    Write-Host "`nv8.4.1 rebuild and verification complete."
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
