[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path (Join-Path $PSScriptRoot '..\releases') 'MTR_Control.xlsm')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-RepositoryRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

function Import-VbaSources {
    param(
        [Parameter(Mandatory = $true)] $Workbook,
        [Parameter(Mandatory = $true)] [string] $SourceRoot
    )

    $extensions = @('*.bas', '*.cls', '*.frm')
    $sourceFiles = foreach ($extension in $extensions) {
        Get-ChildItem -Path $SourceRoot -Filter $extension -Recurse -File
    }

    $sourceFiles = $sourceFiles | Sort-Object FullName
    if (-not $sourceFiles) {
        throw "No VBA source files were found in $SourceRoot."
    }

    foreach ($file in $sourceFiles) {
        Write-Host "Importing $($file.FullName)"
        [void]$Workbook.VBProject.VBComponents.Import($file.FullName)
    }
}

$repoRoot = Resolve-RepositoryRoot
$srcRoot = Join-Path $repoRoot 'src'
$outputFullPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath))
$outputDirectory = Split-Path -Parent $outputFullPath

if (-not (Test-Path $srcRoot)) {
    throw "Source directory not found: $srcRoot"
}

if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

if (Test-Path $outputFullPath) {
    Remove-Item -Path $outputFullPath -Force
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    Import-VbaSources -Workbook $workbook -SourceRoot $srcRoot

    Write-Host 'Creating Dashboard worksheet and buttons'
    $excel.Run('SetupDashboard', $workbook)

    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($outputFullPath, $xlOpenXMLWorkbookMacroEnabled)
    Write-Host "Built $outputFullPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
