# Regenerates van-lookup.js from "Van Allocation.xlsx" (sheet "CAL").
# Run this whenever the xlsx has been replaced with an updated export
# (vans reassigned or added), then commit + push the result.

$ErrorActionPreference = "Stop"
$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxPath = Join-Path $folder "Van Allocation.xlsx"
$outPath = Join-Path $folder "van-lookup.js"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($xlsxPath, $true, $true)
    $ws = $wb.Worksheets.Item("CAL")
    $rows = $ws.UsedRange.Rows.Count

    $map = [ordered]@{}
    for ($r = 2; $r -le $rows; $r++) {
        $van = ([string]$ws.Cells.Item($r, 2).Text).Trim()
        $reg = (([string]$ws.Cells.Item($r, 3).Text) -replace '\s', '').ToUpper()
        if ($van -match '^\d{1,3}$' -and $reg) {
            $van = $van.PadLeft(3, '0')
            $map[$van] = $reg   # later rows overwrite earlier ones - most recent allocation wins
        }
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$pairs = $map.Keys | ForEach-Object { '"' + $_ + '":"' + ($map[$_] -replace '"', '\"') + '"' }
$jsObject = $pairs -join ","
$content = "var VAN_LOOKUP = {$jsObject};`n"
[System.IO.File]::WriteAllText($outPath, $content, (New-Object System.Text.UTF8Encoding($false)))

Write-Output "Wrote $($map.Count) van allocations to van-lookup.js"
