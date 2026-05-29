param(
    [string]$Source = "",
    [string]$Output = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dashboardName = Split-Path -Leaf $PSScriptRoot
$workDir = Join-Path $projectRoot ".tmp\$dashboardName"
$outputDir = Join-Path $projectRoot "outputs\$dashboardName"
$inputDir = Join-Path $projectRoot "inputs\$dashboardName"
$safeComDir = Join-Path $projectRoot ".tmp\excel_com_procyfloc"

if ([string]::IsNullOrWhiteSpace($Source)) {
    $sourceFile = Get-ChildItem -LiteralPath $workDir -Filter "*base-sem-graficos.xlsm" | Select-Object -First 1
    if ($null -eq $sourceFile) { throw "Arquivo base sem graficos nao encontrado em: $workDir" }
    $Source = $sourceFile.FullName
}

if ([string]::IsNullOrWhiteSpace($Output)) {
    $inputFile = Get-ChildItem -LiteralPath $inputDir -Filter "*.xlsm" | Select-Object -First 1
    if ($null -eq $inputFile) { throw "Planilha de entrada nao encontrada em: $inputDir" }
    $Output = Join-Path $outputDir "$($inputFile.BaseName) - Dashboard.xlsm"
}

function XlColor([string]$Hex) {
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return $r + (256 * $g) + (65536 * $b)
}

function Unlock-ChartObject($chartObject) {
    try { $chartObject.Locked = $false } catch {}
    try { $chartObject.Placement = 3 } catch {}
    try { $chartObject.ShapeRange.Locked = $false } catch {}
    try { $chartObject.ShapeRange.LockAspectRatio = 0 } catch {}
}

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

function Invoke-ComRetry([scriptblock]$Action) {
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        try {
            return & $Action
        } catch {
            if ($attempt -eq 8) { throw }
            Start-Sleep -Milliseconds (350 * $attempt)
        }
    }
}

function Wait-ExcelReady($application, [int]$Seconds = 45) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        try {
            if ($application.CalculationState -eq 0) { return }
        } catch {
            return
        }
        Start-Sleep -Milliseconds 500
    }
}

function Convert-ToAbsoluteRange([string]$Range) {
    $items = @()
    foreach ($part in $Range.Split(":")) {
        if ($part -match "^([A-Z]+)([0-9]+)$") {
            $items += ('$' + $matches[1] + '$' + $matches[2])
        }
    }
    return ($items -join ":")
}

function Get-SeriesRangeFormula($ws, [string]$Range) {
    $sheetName = "'" + $ws.Name.Replace("'", "''") + "'!"
    return $sheetName + (Convert-ToAbsoluteRange $Range)
}

function New-NativeChartObject($targetWs, $tempWs, $box, [int]$chartType) {
    $left = $box.Left + 8
    $top = $box.Top + 8
    $width = $box.Width - 16
    $height = $box.Height - 16

    $co = Invoke-ComRetry { $tempWs.ChartObjects().Add(20, 20, $width, $height) }
    Invoke-ComRetry { $co.Chart.ChartType = $chartType } | Out-Null
    Invoke-ComRetry { $co.Chart.Location(2, $targetWs.Name) } | Out-Null
    Start-Sleep -Milliseconds 300

    $co = Invoke-ComRetry { $targetWs.ChartObjects($targetWs.ChartObjects().Count) }
    $co.Left = $left
    $co.Top = $top
    $co.Width = $width
    $co.Height = $height
    Unlock-ChartObject $co
    return $co
}

function Format-ChartBase($chart) {
    $chart.HasTitle = $false
    $chart.HasLegend = $false
    $chart.PlotVisibleOnly = $false
    $chart.ChartArea.Font.Name = "Segoe UI"
    $chart.ChartArea.Font.Size = 8
    try { $chart.ChartArea.Format.Fill.ForeColor.RGB = XlColor "FFFFFF" } catch {}
    try { $chart.ChartArea.Format.Line.ForeColor.RGB = XlColor "D7DEE8" } catch {}
    try { $chart.PlotArea.Format.Fill.ForeColor.RGB = XlColor "FFFFFF" } catch {}
}

function Add-QuantityChart($ws, $tempWs, [string]$anchorRange, [string]$categoryRange, [string]$valueRange, [string]$colorHex, [int]$chartType) {
    $box = $ws.Range($anchorRange)
    $co = New-NativeChartObject $ws $tempWs $box $chartType
    $chart = $co.Chart
    Invoke-ComRetry { $chart.ChartType = $chartType } | Out-Null
    Format-ChartBase $chart

    $series = Invoke-ComRetry { $chart.SeriesCollection().NewSeries() }
    $categories = Get-SeriesRangeFormula $ws $categoryRange
    $values = Get-SeriesRangeFormula $ws $valueRange
    Invoke-ComRetry { $series.Formula = "=SERIES(""Quantidade (kg)"",$categories,$values,1)" } | Out-Null
    try { $series.Format.Fill.ForeColor.RGB = XlColor $colorHex } catch {}
    try { $series.Format.Line.Visible = 0 } catch {}

    Invoke-ComRetry { $series.ApplyDataLabels() } | Out-Null
    $labels = $series.DataLabels()
    try { $labels.ShowSeriesName = $false } catch {}
    try { $labels.ShowCategoryName = $false } catch {}
    try { $labels.ShowValue = $true } catch {}
    try { $labels.NumberFormatLocal = '#.##0 "kg"' } catch {}
    try { $labels.Font.Name = "Segoe UI" } catch {}
    try { $labels.Font.Size = 8 } catch {}
    try { $labels.Font.Color = XlColor "334155" } catch {}
    try { $labels.Position = 2 } catch {}

    $pointCount = 0
    try { $pointCount = Invoke-ComRetry { $series.Points().Count } } catch {}
    for ($i = 1; $i -le $pointCount; $i++) {
        try {
            $value = [double]$ws.Range($valueRange).Cells($i, 1).Value2
            if ($value -le 0) { $series.Points($i).HasDataLabel = $false }
        } catch {}
    }

    try {
        $chart.Axes(1).HasTitle = $false
        $chart.Axes(1).TickLabels.Font.Name = "Segoe UI"
        $chart.Axes(1).TickLabels.Font.Size = 8
        $chart.Axes(1).TickLabels.Font.Color = XlColor "334155"
    } catch {}
    try {
        $chart.Axes(2).HasTitle = $false
        $chart.Axes(2).TickLabels.Font.Name = "Segoe UI"
        $chart.Axes(2).TickLabels.Font.Size = 8
        $chart.Axes(2).TickLabels.Font.Color = XlColor "667085"
        $chart.Axes(2).TickLabels.NumberFormatLocal = '#.##0'
        $chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = XlColor "E7EDF4"
    } catch {}

    Unlock-ChartObject $co
    return $co
}

function Add-StatusChart($ws, $tempWs) {
    $box = $ws.Range("E48:K61")
    $co = New-NativeChartObject $ws $tempWs $box -4120
    $co.Left = $box.Left + 42
    $co.Top = $box.Top + 12
    $co.Width = $box.Width - 84
    $co.Height = $box.Height - 24
    Unlock-ChartObject $co

    $chart = $co.Chart
    Invoke-ComRetry { $chart.ChartType = -4120 } | Out-Null
    Format-ChartBase $chart

    $series = Invoke-ComRetry { $chart.SeriesCollection().NewSeries() }
    $categories = Get-SeriesRangeFormula $ws "M51:M53"
    $values = Get-SeriesRangeFormula $ws "N51:N53"
    Invoke-ComRetry { $series.Formula = "=SERIES(""Recebimentos"",$categories,$values,1)" } | Out-Null

    $colors = @("16A34A", "D97706", "BE123C")
    for ($i = 1; $i -le $colors.Count; $i++) {
        try { $series.Points($i).Format.Fill.ForeColor.RGB = XlColor $colors[$i - 1] } catch {}
    }

    Invoke-ComRetry { $series.ApplyDataLabels() } | Out-Null
    $labels = $series.DataLabels()
    try { $labels.ShowSeriesName = $false } catch {}
    try { $labels.ShowCategoryName = $true } catch {}
    try { $labels.ShowPercentage = $true } catch {}
    try { $labels.ShowValue = $false } catch {}
    try { $labels.Separator = " " } catch {}
    try { $labels.Font.Name = "Segoe UI" } catch {}
    try { $labels.Font.Size = 8 } catch {}
    try { $labels.Font.Color = XlColor "334155" } catch {}

    $pointCount = 0
    try { $pointCount = Invoke-ComRetry { $series.Points().Count } } catch {}
    $pointLimit = [Math]::Min(3, $pointCount)
    for ($i = 1; $i -le $pointLimit; $i++) {
        try {
            $value = [double]$ws.Cells(50 + $i, 14).Value2
            if ($value -le 0) { $series.Points($i).HasDataLabel = $false }
        } catch {}
    }

    Unlock-ChartObject $co
    return $co
}

$excel = $null
$wb = $null

try {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Output) | Out-Null
    if (Test-Path -LiteralPath $safeComDir) {
        Remove-Item -LiteralPath $safeComDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Force -Path $safeComDir | Out-Null
    $workingFile = Join-Path $safeComDir "procyfloc-dashboard.xlsm"
    Copy-Item -LiteralPath $Source -Destination $workingFile -Force
    $resolved = (Resolve-Path -LiteralPath $workingFile).Path

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $true
    try { $excel.AutomationSecurity = 3 } catch {}

    $wb = $excel.Workbooks.Open($resolved, 0, $false, 5, "", "", $true, 1, "", $false, $false, 0, $false, $false, 1)
    try { $wb.Unprotect() | Out-Null } catch {}
    $ws = $wb.Worksheets.Item("Dashboard2")
    try { $wb.Windows.Item(1).Activate() | Out-Null } catch {}
    try { $ws.Select($true) | Out-Null } catch {}
    $ws.Activate() | Out-Null
    try { $ws.Unprotect() | Out-Null } catch {}

    Start-Sleep -Milliseconds 1000

    while ((Invoke-ComRetry { $ws.ChartObjects().Count }) -gt 0) {
        Invoke-ComRetry { $ws.ChartObjects(1).Delete() } | Out-Null
    }

    $tempName = "__GraficoTemp"
    try { $ws.Select($true) | Out-Null } catch {}
    try { $wb.Worksheets.Item($tempName).Delete() } catch {}
    $tempWs = $wb.Worksheets.Add()
    $tempWs.Name = $tempName

    Add-QuantityChart $ws $tempWs "E14:N26" "Q16:Q25" "R16:R25" "2563EB" 57 | Out-Null
    Add-QuantityChart $ws $tempWs "E31:N43" "P33:P44" "R33:R44" "0F766E" 51 | Out-Null
    Add-StatusChart $ws $tempWs | Out-Null

    try { $tempWs.Delete() } catch {}

    $ws.Activate() | Out-Null

    foreach ($shape in $ws.Shapes) {
        try { $shape.Locked = $false } catch {}
        try { $shape.Placement = 3 } catch {}
    }
    try { $ws.EnableSelection = 0 } catch {}

    $movableChartCount = Invoke-ComRetry { $ws.ChartObjects().Count }
    for ($i = 1; $i -le $movableChartCount; $i++) {
        $co = Invoke-ComRetry { $ws.ChartObjects($i) }
        $left = $co.Left
        $top = $co.Top
        $co.Left = $left + 5
        $co.Top = $top + 5
        $co.Left = $left
        $co.Top = $top
        Unlock-ChartObject $co
    }

    $chartCount = Invoke-ComRetry { $ws.ChartObjects().Count }
    Save-WorkbookWithFallback $wb $workingFile
    Close-WorkbookQuietly $wb
    $wb = $null
    Copy-FileWithRetry $workingFile $Output

    Write-Output "file=$Output"
    Write-Output "charts=$chartCount"
    Write-Output "movementTest=passed"
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
