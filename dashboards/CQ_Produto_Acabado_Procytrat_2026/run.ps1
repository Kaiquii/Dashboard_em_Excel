param(
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$dashboardDir = $PSScriptRoot
$projectRoot = Split-Path -Parent (Split-Path -Parent $dashboardDir)
$codexPython = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if ([string]::IsNullOrWhiteSpace($Python)) {
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $Python = "python"
    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $Python = "py"
    } elseif (Test-Path -LiteralPath $codexPython) {
        $Python = $codexPython
    } else {
        throw "Python nao encontrado. Passe o caminho com -Python."
    }
}

Write-Output "1/4 Gerando dashboard base..."
& $Python "$dashboardDir\01_build_dashboard.py"

Write-Output "2/4 Removendo graficos openpyxl..."
& $Python "$dashboardDir\02_strip_charts.py"

Write-Output "3/4 Recriando graficos nativos..."
& powershell -ExecutionPolicy Bypass -File "$dashboardDir\03_recreate_native_charts.ps1"

Write-Output "4/4 Finalizando arquivo..."
& powershell -ExecutionPolicy Bypass -File "$dashboardDir\04_cleanup_final.ps1"

Write-Output "Concluido. Saida em: $projectRoot\outputs\CQ_Produto_Acabado_Procytrat_2026"
