$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$SpecPath = Join-Path $ProjectRoot "StatusManager.spec"

if (-not (Test-Path -LiteralPath $SpecPath)) {
    throw "StatusManager.spec nao encontrado em: $SpecPath"
}

Push-Location $ProjectRoot
try {
    python -m PyInstaller $SpecPath
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "Build concluido. Executavel esperado em:" -ForegroundColor Green
Write-Host (Join-Path $ProjectRoot "dist\StatusManager.exe")
