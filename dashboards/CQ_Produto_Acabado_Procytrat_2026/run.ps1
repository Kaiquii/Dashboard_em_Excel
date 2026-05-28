param(
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$dashboardDir = $PSScriptRoot
$projectRoot = Split-Path -Parent (Split-Path -Parent $dashboardDir)
$codexPython = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$inputFile = "$projectRoot\inputs\CQ_Produto_Acabado_Procytrat_2026\CQ Produto Acabado - Procytrat 2026.xlsm"
$workDir = "$projectRoot\.tmp\CQ_Produto_Acabado_Procytrat_2026"
$outputDir = "$projectRoot\outputs\CQ_Produto_Acabado_Procytrat_2026"
$finalFile = "$outputDir\CQ Produto Acabado - Procytrat 2026 - Dashboard2-sem-simulados.xlsm"

trap {
    if (Test-Path -LiteralPath $finalFile) {
        Remove-Item -LiteralPath $finalFile -Force
    }
    if (Test-Path -LiteralPath $workDir) {
        Remove-Item -LiteralPath $workDir -Recurse -Force
    }
    throw $_
}

function Invoke-Step($Description, $Command, $Arguments) {
    Write-Output $Description
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Etapa falhou: $Description"
    }
}

function Test-PythonCommand($Command) {
    try {
        & $Command -c "import sys; import openpyxl; print(sys.executable)" *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

if ([string]::IsNullOrWhiteSpace($Python)) {
    if ((Test-Path -LiteralPath $codexPython) -and (Test-PythonCommand $codexPython)) {
        $Python = $codexPython
    } elseif ((Get-Command py -ErrorAction SilentlyContinue) -and (Test-PythonCommand "py")) {
        $Python = "py"
    } elseif ((Get-Command python -ErrorAction SilentlyContinue) -and (Test-PythonCommand "python")) {
        $Python = "python"
    } else {
        throw "Python com openpyxl nao encontrado. Instale Python + openpyxl ou passe o caminho com -Python."
    }
} elseif (-not (Test-PythonCommand $Python)) {
    throw "O Python informado nao funcionou ou nao tem openpyxl: $Python"
}

if (-not (Test-Path -LiteralPath $inputFile)) {
    throw "Planilha original nao encontrada. Coloque o arquivo em: $inputFile"
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
foreach ($oldFile in @(
    $finalFile,
    "$outputDir\CQ Produto Acabado - Procytrat 2026 - Dashboard2-profissional-sem-simulados.xlsm",
    "$outputDir\CQ Produto Acabado - Procytrat 2026 - Dashboard2-base-sem-graficos-sem-simulados.xlsm"
)) {
    if (Test-Path -LiteralPath $oldFile) {
        Remove-Item -LiteralPath $oldFile -Force
    }
}
if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}

Invoke-Step "1/4 Gerando dashboard base..." $Python @("$dashboardDir\01_build_dashboard.py")

Invoke-Step "2/4 Removendo graficos openpyxl..." $Python @("$dashboardDir\02_strip_charts.py")

Invoke-Step "3/4 Recriando graficos nativos..." "powershell" @("-ExecutionPolicy", "Bypass", "-File", "$dashboardDir\03_recreate_native_charts.ps1")

Invoke-Step "4/4 Finalizando arquivo..." "powershell" @("-ExecutionPolicy", "Bypass", "-File", "$dashboardDir\04_cleanup_final.ps1")

if (-not (Test-Path -LiteralPath $finalFile)) {
    throw "Arquivo final nao foi gerado: $finalFile"
}

if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}

Write-Output "Concluido. Saida em: $outputDir"
