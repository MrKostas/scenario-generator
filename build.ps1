param(
    [string]$Python = (Join-Path $PSScriptRoot '.venv\Scripts\python.exe')
)

$ErrorActionPreference = 'Stop'
$env:SCENARIO_BUILD_MODE = 'onefile'
try {
    & $Python -m PyInstaller --noconfirm --clean `
        --distpath (Join-Path $PSScriptRoot 'dist') `
        --workpath (Join-Path $PSScriptRoot 'build') `
        (Join-Path $PSScriptRoot 'scenario-generator.spec')
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller exited with code $LASTEXITCODE" }
} finally {
    Remove-Item Env:\SCENARIO_BUILD_MODE -ErrorAction SilentlyContinue
}
