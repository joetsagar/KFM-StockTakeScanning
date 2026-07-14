# Regenerates tech-lookup.js from "Resource ID Info (Active).xlsx".
# Run this whenever the xlsx has been replaced with an updated export
# (technicians added or removed), then commit + push the result.

$ErrorActionPreference = "Stop"
$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxPath = Join-Path $folder "Resource ID Info (Active).xlsx"
$outPath = Join-Path $folder "tech-lookup.js"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($xlsxPath, $true, $true)
    $ws = $wb.Worksheets.Item("Resource ID Info (Active)")
    $rows = $ws.UsedRange.Rows.Count

    $names = New-Object System.Collections.Generic.List[string]
    for ($r = 2; $r -le $rows; $r++) {
        $active = [string]$ws.Cells.Item($r, 5).Text
        $name = [string]$ws.Cells.Item($r, 2).Text
        if ($name -and $active -eq "TRUE") { $names.Add($name) }
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$sorted = $names | Sort-Object -Unique
$jsArray = ($sorted | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join ","
$content = "var TECH_NAMES = [$jsArray];`n"
[System.IO.File]::WriteAllText($outPath, $content, (New-Object System.Text.UTF8Encoding($false)))

Write-Output "Wrote $($sorted.Count) technician names to tech-lookup.js"
