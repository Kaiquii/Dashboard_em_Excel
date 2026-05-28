param(
    [string]$Source = "$((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))\.tmp\CQ_Produto_Acabado_Procytrat_2026\CQ Produto Acabado - Procytrat 2026 - Dashboard2-base-sem-graficos-sem-simulados.xlsm",
    [string]$Output = "$((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))\outputs\CQ_Produto_Acabado_Procytrat_2026\CQ Produto Acabado - Procytrat 2026 - Dashboard2-sem-simulados.xlsm"
)

$ErrorActionPreference = "Stop"

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

function New-NativeChartObject($targetWs, $tempWs, $box, [int]$chartType) {
    $left = $box.Left + 8
    $top = $box.Top + 8
    $width = $box.Width - 16
    $height = $box.Height - 16

    $co = $tempWs.ChartObjects().Add(20, 20, $width, $height)
    $co.Chart.ChartType = $chartType
    $co.Chart.Location(2, $targetWs.Name) | Out-Null
    Start-Sleep -Milliseconds 300

    $newCo = $targetWs.ChartObjects($targetWs.ChartObjects().Count)
    $newCo.Left = $left
    $newCo.Top = $top
    $newCo.Width = $width
    $newCo.Height = $height
    return $newCo
}

function Add-BarChart($ws, $tempWs, [string]$anchorRange, [string]$categoryRange, [string]$valueRange, [string]$colorHex, [int]$chartType) {
    $box = $ws.Range($anchorRange)
    $co = New-NativeChartObject $ws $tempWs $box $chartType
    Unlock-ChartObject $co

    $chart = $co.Chart
    $chart.ChartType = $chartType
    $chart.HasTitle = $false
    $chart.HasLegend = $false
    $chart.PlotVisibleOnly = $false
    $chart.ChartArea.Font.Name = "Segoe UI"
    $chart.ChartArea.Font.Size = 8
    try { $chart.ChartArea.Format.Fill.ForeColor.RGB = XlColor "FFFFFF" } catch {}
    try { $chart.ChartArea.Format.Line.ForeColor.RGB = XlColor "D7DEE8" } catch {}
    try { $chart.PlotArea.Format.Fill.ForeColor.RGB = XlColor "FFFFFF" } catch {}

    $series = $chart.SeriesCollection().NewSeries()
    $series.Name = "Quantidade (kg)"
    $series.Values = $ws.Range($valueRange)
    $series.XValues = $ws.Range($categoryRange)
    try { $series.Format.Fill.ForeColor.RGB = XlColor $colorHex } catch {}
    try { $series.Format.Line.Visible = 0 } catch {}
    $series.ApplyDataLabels()
    $series.DataLabels().NumberFormatLocal = '#.##0 "kg"'
    $series.DataLabels().Font.Name = "Segoe UI"
    $series.DataLabels().Font.Size = 8
    $series.DataLabels().Font.Color = XlColor "334155"
    try { $series.DataLabels().Position = 2 } catch {}

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
    $chart.ChartType = -4120
    $chart.HasTitle = $false
    $chart.HasLegend = $true
    $chart.Legend.Position = -4152
    $chart.ChartArea.Font.Name = "Segoe UI"
    $chart.ChartArea.Font.Size = 8
    try { $chart.ChartArea.Format.Fill.ForeColor.RGB = XlColor "FFFFFF" } catch {}
    try { $chart.ChartArea.Format.Line.ForeColor.RGB = XlColor "D7DEE8" } catch {}

    $series = $chart.SeriesCollection().NewSeries()
    $series.Name = "Lotes"
    $series.Values = $ws.Range("N51:N53")
    $series.XValues = $ws.Range("M51:M53")

    $colors = @("16A34A", "D97706", "BE123C")
    for ($i = 1; $i -le $colors.Count; $i++) {
        try { $series.Points($i).Format.Fill.ForeColor.RGB = XlColor $colors[$i - 1] } catch {}
    }

    $series.ApplyDataLabels()
    $series.DataLabels().ShowCategoryName = $true
    $series.DataLabels().ShowPercentage = $true
    $series.DataLabels().ShowValue = $false
    $series.DataLabels().Separator = " "
    $series.DataLabels().Font.Name = "Segoe UI"
    $series.DataLabels().Font.Size = 8
    $series.DataLabels().Font.Color = XlColor "334155"
    for ($i = 1; $i -le 3; $i++) {
        try {
            $value = [double]$ws.Cells(50 + $i, 14).Value2
            if ($value -le 0) { $series.Points($i).HasDataLabel = $false }
        } catch {}
    }

    return $co
}

$excel = $null
$wb = $null

try {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Output) | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Output -Force
    $resolved = (Resolve-Path -LiteralPath $Output).Path

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $true
    try { $excel.AutomationSecurity = 3 } catch {}

    $wb = $excel.Workbooks.Open($resolved)
    Start-Sleep -Milliseconds 1000
    try { $wb.Unprotect() | Out-Null } catch {}
    $ws = $wb.Worksheets.Item("Dashboard2")
    $ws.Activate() | Out-Null
    try { $ws.Unprotect() | Out-Null } catch {}

    $excel.CalculateFullRebuild()
    Start-Sleep -Milliseconds 500

    while ($ws.ChartObjects().Count -gt 0) {
        try { $ws.ChartObjects(1).Delete() } catch { break }
    }

    $tempName = "__GraficoTemp"
    try {
        $oldTemp = $wb.Worksheets.Item($tempName)
        $oldTemp.Delete()
    } catch {}
    $tempWs = $wb.Worksheets.Add()
    $tempWs.Name = $tempName

    $top = Add-BarChart $ws $tempWs "E14:N26" "Q16:Q25" "R16:R25" "2563EB" 57
    $monthly = Add-BarChart $ws $tempWs "E31:N43" "P33:P44" "R33:R44" "0F766E" 51
    $status = Add-StatusChart $ws $tempWs

    $tempWs.Delete()
    $ws.Activate() | Out-Null

    foreach ($shape in $ws.Shapes) {
        try { $shape.Locked = $false } catch {}
        try { $shape.Placement = 3 } catch {}
    }
    $ws.EnableSelection = 0

    # Practical movement test: move each chart 5px and move it back. If this fails,
    # the object is not really movable in Excel.
    for ($i = 1; $i -le $ws.ChartObjects().Count; $i++) {
        $co = $ws.ChartObjects($i)
        $left = $co.Left
        $topPos = $co.Top
        $co.Left = $left + 5
        $co.Top = $topPos + 5
        $co.Left = $left
        $co.Top = $topPos
        Unlock-ChartObject $co
    }

    $errorCount = 0
    try {
        $errorCount = $ws.UsedRange.SpecialCells(-4123, 16).Count
    } catch {
        $errorCount = 0
    }

    $details = @()
    for ($i = 1; $i -le $ws.ChartObjects().Count; $i++) {
        $co = $ws.ChartObjects($i)
        $details += "chart$i locked=$($co.Locked) placement=$($co.Placement) left=$([math]::Round($co.Left,1)) top=$([math]::Round($co.Top,1))"
    }

    $wb.Save()
    Write-Output "file=$resolved"
    Write-Output "charts=$($ws.ChartObjects().Count)"
    Write-Output ($details -join "; ")
    Write-Output "protected=$($ws.ProtectContents)"
    Write-Output "formulaErrors=$errorCount"
    Write-Output "movementTest=passed"
}
finally {
    if ($wb -ne $null) { $wb.Close($true) | Out-Null }
    if ($excel -ne $null) { $excel.Quit() | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
