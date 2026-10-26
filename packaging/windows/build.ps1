param(
    [string]$Python = 'python',
    [switch]$SkipInstall,
    [switch]$Console
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Build this distribution on Windows.' }
$buildRoot = $PSScriptRoot
$buildPython = Join-Path $buildRoot '.build-venv\Scripts\python.exe'
$previousConsole = $env:PPT_TRANSLATOR_BUILD_CONSOLE

function Invoke-Checked {
    param([string]$Executable, [string[]]$Arguments)
    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Command failed (exit $LASTEXITCODE): $Executable" }
}

Push-Location $buildRoot
try {
    if (-not (Test-Path -LiteralPath $buildPython)) {
        Invoke-Checked -Executable $Python -Arguments @('-m', 'venv', (Join-Path $buildRoot '.build-venv'))
    }
    if (-not $SkipInstall) {
        Invoke-Checked -Executable $buildPython -Arguments @('-m', 'pip', 'install', '-r', (Join-Path $buildRoot 'requirements-build.txt'))
    }
    # A skipped install must not silently produce a build missing PDF support.
    Invoke-Checked -Executable $buildPython -Arguments @((Join-Path $buildRoot 'check_dependencies.py'))
    $env:PPT_TRANSLATOR_BUILD_CONSOLE = if ($Console) { '1' } else { '0' }
    Invoke-Checked -Executable $buildPython -Arguments @(
        '-m', 'PyInstaller', '--noconfirm', '--clean',
        '--distpath', (Join-Path $buildRoot 'dist'),
        '--workpath', (Join-Path $buildRoot 'build'),
        (Join-Path $buildRoot 'PPTTranslator.spec')
    )
    $bundle = Join-Path $buildRoot 'dist\PPTTranslator'
    Copy-Item -LiteralPath (Join-Path $buildRoot 'DISTRIBUTION.txt') -Destination $bundle
    $zip = Join-Path $buildRoot 'dist\PPTTranslator-windows.zip'
    Invoke-Checked -Executable $buildPython -Arguments @((Join-Path $buildRoot 'archive.py'), $bundle, $zip)
    Write-Host "Executable: $bundle\PPTTranslator.exe"
    Write-Host "Distribute the complete folder, or: $zip"
} finally {
    $env:PPT_TRANSLATOR_BUILD_CONSOLE = $previousConsole
    Pop-Location
}
