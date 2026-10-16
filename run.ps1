$scriptPath = Join-Path $PSScriptRoot 'protocol_app.py'
$bundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if (Test-Path -LiteralPath $bundledPython) {
    & $bundledPython $scriptPath
    exit $LASTEXITCODE
}

Write-Host 'The local Python runtime was not found.' -ForegroundColor Yellow
Write-Host 'Open this folder in Codex and run Запустить.bat again.'
Read-Host 'Press Enter to close this window'
