<#
  Exports modTimeToTable VBA module and document VBA modules
  from TimeToTable_VBA.xlsm into the tools folder.
  Invoked automatically as a PostToolUse hook after generate_vba_workbook.ps1
#>
param()
$ErrorActionPreference = 'Stop'

$root     = Split-Path $PSScriptRoot -Parent
$xlsmPath = Join-Path $root 'TimeToTable_VBA.xlsm'
$basPath  = Join-Path $PSScriptRoot 'modTimeToTable_manual.bas'

function Get-SafeFileStem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $Name.ToCharArray()) {
        if ($invalidChars -contains $char) {
            [void]$builder.Append('_')
        } else {
            [void]$builder.Append($char)
        }
    }

    $safeName = $builder.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return 'unnamed'
    }

    return $safeName
}

function Write-CodeModuleText {
    param(
        [Parameter(Mandatory = $true)]
        $Component,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $codeModule = $Component.CodeModule
    $lineCount = $codeModule.CountOfLines
    if ($lineCount -gt 0) {
        $text = $codeModule.Lines(1, $lineCount)
    } else {
        $text = ''
    }

    $text = [regex]::Replace(
        $text,
        '(?ms)\A\s*VERSION [^\r\n]+\r?\nBEGIN\r?\n.*?^\s*END\r?\n',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?m)^\s*Attribute VB_[^\r\n]*\r?\n',
        ''
    )
    $text = $text.TrimStart("`r", "`n")

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $text, $utf8NoBom)
}

$key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$old = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -ErrorAction SilentlyContinue

$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false

    $wb = $xl.Workbooks.Open($xlsmPath)
    $exported = New-Object System.Collections.Generic.List[string]

    $mainComponent = $wb.VBProject.VBComponents.Item('modTimeToTable')
    Write-CodeModuleText -Component $mainComponent -Path $basPath
    $exported.Add("modTimeToTable -> $(Split-Path $basPath -Leaf)") | Out-Null

    $legacyThisWorkbookPath = Join-Path $PSScriptRoot 'ThisWorkbook_manual.cls'
    if (Test-Path $legacyThisWorkbookPath) {
        Remove-Item -LiteralPath $legacyThisWorkbookPath -Force
    }
    $thisWorkbookPath = Join-Path $PSScriptRoot 'ThisWorkbook_manual.bas'
    $thisWorkbookComponent = $wb.VBProject.VBComponents.Item($wb.CodeName)
    Write-CodeModuleText -Component $thisWorkbookComponent -Path $thisWorkbookPath
    $exported.Add("$($wb.CodeName) -> $(Split-Path $thisWorkbookPath -Leaf)") | Out-Null

    foreach ($ws in $wb.Worksheets) {
        $component = $wb.VBProject.VBComponents.Item($ws.CodeName)
        $legacySheetFileName = 'sheet_{0}_manual.cls' -f (Get-SafeFileStem -Name $ws.Name)
        $legacySheetPath = Join-Path $PSScriptRoot $legacySheetFileName
        if (Test-Path $legacySheetPath) {
            Remove-Item -LiteralPath $legacySheetPath -Force
        }
        $sheetFileName = 'sheet_{0}_manual.bas' -f (Get-SafeFileStem -Name $ws.Name)
        $sheetPath = Join-Path $PSScriptRoot $sheetFileName
        Write-CodeModuleText -Component $component -Path $sheetPath
        $exported.Add("$($ws.CodeName) [$($ws.Name)] -> $sheetFileName") | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
    }

    Write-Host 'Exported VBA components:'
    foreach ($item in $exported) {
        Write-Host "  $item"
    }
} finally {
    if ($null -ne $wb) {
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($null -ne $xl) {
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    }

    if ($null -ne $old) {
        Set-ItemProperty  -Path $key -Name AccessVBOM -Value $old -ErrorAction SilentlyContinue
    } else {
        Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
