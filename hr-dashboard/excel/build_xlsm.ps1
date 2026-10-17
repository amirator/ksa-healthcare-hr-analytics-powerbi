# Builds "KSA Healthcare HR Analytics.xlsm": opens the data workbook, injects the VBA
# engine, runs the full dashboard build once, verifies, and saves.
$ErrorActionPreference = "Stop"
$root  = Split-Path -Parent $PSScriptRoot
$ex    = Join-Path $root "excel"
$dataX = Join-Path $ex "HR_Data.xlsx"
$out   = Join-Path $root "KSA Healthcare HR Analytics.xlsm"

# --- temporarily allow programmatic access to the VBA project (restored below)
$changed = @{}
Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d$' } | ForEach-Object {
        if (Test-Path (Join-Path $_.PSPath "Excel")) {
            $sec = Join-Path $_.PSPath "Excel\Security"
            if (-not (Test-Path $sec)) { New-Item $sec -Force | Out-Null }
            $old = (Get-ItemProperty $sec -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
            $changed[$sec] = $old
            Set-ItemProperty $sec -Name AccessVBOM -Value 1 -Type DWord
        }
    }

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($dataX)

    $vbp = $wb.VBProject
    Get-ChildItem (Join-Path $ex "vba\*.bas") | ForEach-Object {
        [void]$vbp.VBComponents.Import($_.FullName)
        Write-Host "imported  $($_.Name)"
    }
    $tw = $vbp.VBComponents.Item("ThisWorkbook").CodeModule
    $tw.AddFromString((Get-Content (Join-Path $ex "vba\ThisWorkbook.vba") -Raw))
    Write-Host "imported  ThisWorkbook events"

    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled

    Write-Host "running   modMain.BuildAll ..."
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $xl.Run("modMain.BuildAll")
    $sw.Stop()
    Write-Host ("BuildAll completed in {0:n1}s" -f $sw.Elapsed.TotalSeconds)
    $wb.Save()

    # --- verification ---------------------------------------------------------
    $snapRows = $wb.Sheets("Snapshots").UsedRange.Rows.Count
    $calcRows = $wb.Sheets("_Calc").UsedRange.Rows.Count
    $nCharts  = $wb.Sheets("Dashboard").ChartObjects().Count
    $nPivots  = $wb.Sheets("PivotExplorer").PivotTables().Count
    $nAlerts  = $wb.Sheets("Alerts").UsedRange.Rows.Count - 3
    $nShapes  = $wb.Sheets("Dashboard").Shapes.Count
    Write-Host "verify: snapshots=$($snapRows-1) rows, _Calc=$calcRows rows, charts=$nCharts, pivots=$nPivots, alerts=$nAlerts, dashboard shapes=$nShapes"
    if ($nCharts -ne 4 -or $nPivots -ne 3 -or $snapRows -lt 1000) { throw "verification failed" }

    $wb.Close($true)
    $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # restore the trust-access setting exactly as it was
    foreach ($k in $changed.Keys) {
        if ($null -eq $changed[$k]) {
            Remove-ItemProperty $k -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty $k -Name AccessVBOM -Value $changed[$k] -Type DWord
        }
    }
}
Write-Host "BUILT: $out"
