param(
    [string]$Path = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dashboardName = Split-Path -Leaf $PSScriptRoot
$outputDir = Join-Path $projectRoot "outputs\$dashboardName"
$safeComDir = Join-Path $projectRoot ".tmp\excel_com_materia_prima_cleanup"

if ([string]::IsNullOrWhiteSpace($Path)) {
    $outputFile = Get-ChildItem -LiteralPath $outputDir -Filter "*Dashboard.xlsm" | Select-Object -First 1
    if ($null -eq $outputFile) { throw "Arquivo final nao encontrado em: $outputDir" }
    $Path = $outputFile.FullName
}

$excel = $null
$wb = $null

function Close-WorkbookQuietly($workbook) {
    if ($null -eq $workbook) { return }
    try { $workbook.Close($true) | Out-Null } catch {}
}

function Quit-ExcelQuietly($application) {
    if ($null -eq $application) { return }
    try { $application.Quit() | Out-Null } catch {}
}

function Copy-FileWithRetry([string]$From, [string]$To) {
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Copy-Item -LiteralPath $From -Destination $To -Force
            return
        } catch {
            if ($attempt -eq 5) { throw }
            Start-Sleep -Milliseconds (300 * $attempt)
        }
    }
}

function Save-WorkbookWithFallback($workbook, [string]$Path) {
    try { $workbook.CheckCompatibility = $false } catch {}
    try {
        $workbook.Save()
        return
    } catch {
        $firstError = $_
    }

    try {
        $workbook.SaveAs($Path, 52)
        return
    } catch {}

    try {
        $workbook.SaveCopyAs($Path)
        return
    } catch {}

    throw $firstError
}

try {
    New-Item -ItemType Directory -Force -Path $safeComDir | Out-Null
    $workingFile = Join-Path $safeComDir "materia-prima-dashboard-cleanup.xlsm"
    Copy-Item -LiteralPath $Path -Destination $workingFile -Force
    $resolved = (Resolve-Path -LiteralPath $workingFile).Path
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch {}

    $wb = $excel.Workbooks.Open($resolved, 0, $false, 5, "", "", $true, 1, "", $false, $false, 0, $false, $false, 1)
    foreach ($name in @("__GraficoTemp", "Planilha4")) {
        try { $wb.Worksheets.Item($name).Delete() } catch {}
    }

    $ws = $wb.Worksheets.Item("Dashboard2")
    $ws.Move($wb.Worksheets.Item(1))
    $ws = $wb.Worksheets.Item("Dashboard2")
    $ws.Activate() | Out-Null
    try { $ws.Unprotect() | Out-Null } catch {}

    $excel.CalculateFullRebuild()
    Start-Sleep -Milliseconds 500

    $details = @()
    for ($i = 1; $i -le $ws.ChartObjects().Count; $i++) {
        $co = $ws.ChartObjects($i)
        $left = $co.Left
        $top = $co.Top
        $co.Locked = $false
        $co.Placement = 3
        try { $co.ShapeRange.Locked = $false } catch {}
        $co.Left = $left + 4
        $co.Top = $top + 4
        $moved = ([math]::Abs($co.Left - ($left + 4)) -lt 0.2) -and ([math]::Abs($co.Top - ($top + 4)) -lt 0.2)
        $co.Left = $left
        $co.Top = $top
        $details += "chart$i moved=$moved locked=$($co.Locked) placement=$($co.Placement)"
    }

    $errorCount = 0
    try {
        $errorCount = $ws.UsedRange.SpecialCells(-4123, 16).Count
    } catch {
        $errorCount = 0
    }

    $firstSheet = $wb.Worksheets.Item(1).Name
    $chartCount = $ws.ChartObjects().Count
    Save-WorkbookWithFallback $wb $workingFile
    Close-WorkbookQuietly $wb
    $wb = $null
    Copy-FileWithRetry $workingFile $Path

    Write-Output "file=$Path"
    Write-Output "firstSheet=$firstSheet"
    Write-Output "charts=$chartCount"
    Write-Output ($details -join "; ")
    Write-Output "formulaErrors=$errorCount"
}
finally {
    Close-WorkbookQuietly $wb
    Quit-ExcelQuietly $excel
    Start-Sleep -Milliseconds 500
    try {
        if (Test-Path -LiteralPath $safeComDir) {
            Remove-Item -LiteralPath $safeComDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
