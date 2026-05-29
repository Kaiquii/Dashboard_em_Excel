param(
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$dashboard = "CQ_Produto_Acabado_Procyfloc_2026"
$dashboardDir = $PSScriptRoot
$projectRoot = Split-Path -Parent (Split-Path -Parent $dashboardDir)
$runner = Join-Path $projectRoot "gerar_dashboard.py"
$codexPython = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if ([string]::IsNullOrWhiteSpace($Python)) {
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $Python = "python"
    } elseif (Test-Path -LiteralPath $codexPython) {
        $Python = $codexPython
    } else {
        throw "Python nao encontrado. Passe o caminho com -Python."
    }
}

& $Python $runner $dashboard
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
