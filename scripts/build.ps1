<#
.SYNOPSIS
    Builds APIProductIntelligenceDemo.xlsm by importing VBA source files into Excel.

.DESCRIPTION
    This script:
      1. Opens a blank Excel workbook.
      2. Imports all .bas and .cls modules from the src\ folder.
      3. Adds a "Refresh Data" button to the Dashboard sheet.
      4. Saves the result as APIProductIntelligenceDemo.xlsm.

.PREREQUISITES
    • Microsoft Excel must be installed.
    • In Excel: File -> Options -> Trust Center -> Trust Center Settings ->
      Macro Settings -> check "Trust access to the VBA project object model".

.USAGE
    .\scripts\build.ps1
    .\scripts\build.ps1 -OutputPath "C:\MyFolder\demo.xlsm"
#>

param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "APIProductIntelligenceDemo.xlsm")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$srcDir = Join-Path (Split-Path $PSScriptRoot -Parent) "src"

Write-Host "`n=== APIProductIntelligenceDemo build ===" -ForegroundColor Cyan
Write-Host "Source dir : $srcDir"
Write-Host "Output     : $OutputPath`n"

if (-not (Test-Path $srcDir)) {
    Write-Error "Source directory not found: $srcDir"
    exit 1
}

# ---------------------------------------------------------------------------
# Launch Excel (hidden)
# ---------------------------------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible        = $false
$excel.DisplayAlerts  = $false

try {
    $wb  = $excel.Workbooks.Add()
    $vba = $wb.VBProject

    # -----------------------------------------------------------------------
    # Import standard modules (.bas)
    # -----------------------------------------------------------------------
    $basModules = @(
        "ModernJsonInVBA.bas",
        "ProductDataFetcher.bas",
        "DashboardController.bas",
        "WorkbookSetup.bas"
    )

    foreach ($mod in $basModules) {
        $path = Join-Path $srcDir $mod
        if (Test-Path $path) {
            $null = $vba.VBComponents.Import($path)
            Write-Host "  Imported  $mod" -ForegroundColor Green
        } else {
            Write-Warning "  Not found: $path"
        }
    }

    # -----------------------------------------------------------------------
    # Merge ThisWorkbook.cls into the existing ThisWorkbook component.
    # We cannot import it as a new component – Excel already has one.
    # -----------------------------------------------------------------------
    $thisWbPath = Join-Path $srcDir "ThisWorkbook.cls"
    if (Test-Path $thisWbPath) {
        $lines    = Get-Content $thisWbPath
        $codeBody = @()
        $inCode   = $false

        foreach ($line in $lines) {
            # Skip .cls header lines (everything before Option Explicit / first Sub)
            if (-not $inCode) {
                if ($line -match "^Option Explicit" -or $line -match "^Private Sub" -or $line -match "^Public Sub") {
                    $inCode = $true
                }
            }
            if ($inCode) { $codeBody += $line }
        }

        $codeStr   = $codeBody -join "`n"
        $thisWbCmp = $vba.VBComponents.Item("ThisWorkbook")
        $thisWbCmp.CodeModule.AddFromString($codeStr)
        Write-Host "  Merged    ThisWorkbook.cls" -ForegroundColor Green
    }

    # -----------------------------------------------------------------------
    # Add a "Refresh Data" button to Sheet1 (will become Dashboard after open)
    # We place it high up so it's visible once the Dashboard initialises.
    # -----------------------------------------------------------------------
    $sheet = $wb.Worksheets.Item(1)

    $btn       = $sheet.Buttons.Add(10, 10, 130, 28)
    $btn.Text  = "Refresh Data"
    $btn.Name  = "btnRefresh"
    $btn.OnAction = "RefreshAll"

    # -----------------------------------------------------------------------
    # Save as macro-enabled workbook  (52 = xlOpenXMLWorkbookMacroEnabled)
    # -----------------------------------------------------------------------
    $wb.SaveAs($OutputPath, 52)
    Write-Host "`nSaved -> $OutputPath" -ForegroundColor Cyan
    Write-Host "Build complete.  Open the .xlsm file in Excel to run the demo.`n" -ForegroundColor Green

} finally {
    try { $wb.Close($false) }  catch {}
    try { $excel.Quit() }      catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
