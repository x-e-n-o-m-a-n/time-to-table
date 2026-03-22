[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outFile = Join-Path (Get-Location) "tools\export_analysis.txt"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$filePath = (Resolve-Path "EXPORT.XLSX").Path
$wb = $excel.Workbooks.Open($filePath)
$ws = $wb.Sheets.Item(1)
$ur = $ws.UsedRange
$rowCount = $ur.Rows.Count
$colCount = $ur.Columns.Count

$lines = @()
$lines += "Sheet: $($ws.Name)  Rows: $rowCount  Cols: $colCount"
$lines += ""
$lines += "=== HEADERS ==="
for ($c = 1; $c -le $colCount; $c++) {
    $letter = ""
    if ($c -le 26) { $letter = [char](64 + $c) }
    else { $letter = "A" + [char](64 + $c - 26) }
    $val = $ws.Cells.Item(1, $c).Text
    $lines += "  Col $letter ($c): $val"
}

$lines += ""
$lines += "=== SAMPLE DATA (rows 2-5) ==="
for ($r = 2; $r -le 5; $r++) {
    $lines += "--- Row $r ---"
    for ($c = 1; $c -le $colCount; $c++) {
        $letter = ""
        if ($c -le 26) { $letter = [char](64 + $c) }
        else { $letter = "A" + [char](64 + $c - 26) }
        $hdr = $ws.Cells.Item(1, $c).Text
        $val = $ws.Cells.Item($r, $c).Text
        $lines += "  $letter ($hdr): $val"
    }
}

$lines += ""
$lines += "=== UNIQUE VALUES ==="
$col2Vals = @{}
$col13Vals = @{}
for ($r = 2; $r -le $rowCount; $r++) {
    $v2 = $ws.Cells.Item($r, 2).Text
    $v13 = $ws.Cells.Item($r, 13).Text
    if ($v2 -ne "") { $col2Vals[$v2] = ($col2Vals[$v2] + 1) }
    if ($v13 -ne "") { $col13Vals[$v13] = 1 }
}
$lines += "Col B (unique count): $($col2Vals.Count)"
foreach ($k in ($col2Vals.Keys | Select-Object -First 10)) {
    $lines += "  '$k': $($col2Vals[$k]) rows"
}
$lines += "Col M (unique count): $($col13Vals.Count)"
foreach ($k in $col13Vals.Keys) {
    $lines += "  '$k'"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

$lines | Out-File -FilePath $outFile -Encoding UTF8
Write-Host "Done: $outFile"
