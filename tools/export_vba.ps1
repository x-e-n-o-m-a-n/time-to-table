<#
  Exports modTimeToTable VBA module from TimeToTable_VBA.xlsm -> modTimeToTable_manual.bas
  Invoked automatically as a PostToolUse hook after generate_vba_workbook.ps1
#>
param()
$ErrorActionPreference = 'Stop'

$root     = Split-Path $PSScriptRoot -Parent
$xlsmPath = Join-Path $root 'TimeToTable_VBA.xlsm'
$basPath  = Join-Path $PSScriptRoot 'modTimeToTable_manual.bas'

$key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$old = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -ErrorAction SilentlyContinue

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false

    $wb = $xl.Workbooks.Open($xlsmPath)
    $wb.VBProject.VBComponents.Item('modTimeToTable').Export($basPath)
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

    Write-Host "Exported: modTimeToTable -> modTimeToTable_manual.bas"
} finally {
    if ($null -ne $old) {
        Set-ItemProperty  -Path $key -Name AccessVBOM -Value $old -ErrorAction SilentlyContinue
    } else {
        Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}
