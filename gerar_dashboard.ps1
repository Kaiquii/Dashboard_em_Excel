param(
    [string]$Dashboard = "CQ_Produto_Acabado_Procytrat_2026",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$dashboardDir = Join-Path $PSScriptRoot "dashboards\$Dashboard"
$runner = Join-Path $dashboardDir "run.ps1"

if (-not (Test-Path -LiteralPath $runner)) {
    throw "Dashboard nao encontrado ou sem run.ps1: $Dashboard"
}

& powershell -ExecutionPolicy Bypass -File $runner -Python $Python
