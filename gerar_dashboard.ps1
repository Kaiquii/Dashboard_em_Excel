param(
    [string]$Dashboard = ("CQ_Mat" + [char]0x00E9 + "ria_Prima_2026"),
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$dashboardDir = Join-Path $PSScriptRoot "dashboards\$Dashboard"
$runner = Join-Path $dashboardDir "run.ps1"

if (-not (Test-Path -LiteralPath $runner)) {
    throw "Dashboard nao encontrado ou sem run.ps1: $Dashboard"
}

if ([string]::IsNullOrWhiteSpace($Python)) {
    & powershell -ExecutionPolicy Bypass -File $runner
} else {
    & powershell -ExecutionPolicy Bypass -File $runner -Python $Python
}

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
