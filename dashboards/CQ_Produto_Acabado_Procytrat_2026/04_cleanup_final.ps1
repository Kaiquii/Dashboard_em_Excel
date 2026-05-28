param(
    [string]$Path = "$((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))\outputs\CQ_Produto_Acabado_Procytrat_2026\CQ Produto Acabado - Procytrat 2026 - Dashboard2-sem-simulados.xlsm"
)

$ErrorActionPreference = "Stop"

$excel = $null
$wb = $null

try {
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch {}

    $wb = $excel.Workbooks.Open($resolved, 0, $false)
    if ($wb -eq $null) {
        throw "Excel returned a null workbook."
    }

    foreach ($name in @("__GraficoTemp", "Planilha4")) {
        try {
            $sheet = $wb.Worksheets.Item($name)
            if ($sheet.Name -eq $name) {
                $sheet.Delete()
            }
        } catch {}
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

    $wb.Save()
    Write-Output "file=$resolved"
    Write-Output "firstSheet=$($wb.Worksheets.Item(1).Name)"
    Write-Output "sheets=$($wb.Worksheets.Count)"
    Write-Output "charts=$($ws.ChartObjects().Count)"
    Write-Output ($details -join "; ")
    Write-Output "protected=$($ws.ProtectContents)"
    Write-Output "formulaErrors=$errorCount"
    Write-Output "summary=$($ws.Range('E64').Text)|$($ws.Range('E66').Text)|$($ws.Range('E67').Text)"
}
finally {
    if ($wb -ne $null) { $wb.Close($true) | Out-Null }
    if ($excel -ne $null) { $excel.Quit() | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
