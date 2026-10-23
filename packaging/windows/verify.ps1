# Offline smoke check: bundled resources and GUI, with Python removed from PATH.
$ErrorActionPreference = 'Stop'
$bundle = Join-Path $PSScriptRoot 'dist\PPTTranslator'
$exe = Join-Path $bundle 'PPTTranslator.exe'
$resource = Join-Path $bundle '_internal\pptx_pipeline\render_powerpoint.ps1'
$source = Join-Path $PSScriptRoot '..\..\pptx_pipeline\render_powerpoint.ps1'
if (-not (Test-Path -LiteralPath $exe)) { throw 'Build the executable first.' }
if (-not (Test-Path -LiteralPath $resource)) { throw 'Bundled renderer is missing.' }
if ((Get-FileHash -LiteralPath $source).Hash -ne (Get-FileHash -LiteralPath $resource).Hash) {
    throw 'Bundled renderer differs from source.'
}

$smokeDir = Join-Path $PSScriptRoot 'build'
New-Item -ItemType Directory -Path $smokeDir -Force | Out-Null
$screenshot = Join-Path $smokeDir ('smoke-' + [guid]::NewGuid().ToString('N') + '.png')
$previousPath = $env:PATH
$previousPythonHome = $env:PYTHONHOME
$previousPythonPath = $env:PYTHONPATH
$previousQtPlatform = $env:QT_QPA_PLATFORM
try {
    $env:PATH = ''
    $env:PYTHONHOME = ''
    $env:PYTHONPATH = ''
    $env:QT_QPA_PLATFORM = 'offscreen'
    $process = Start-Process -FilePath $exe -ArgumentList @(
        '--demo', '--screenshot', ('"' + $screenshot + '"'), '--quit-after', '2200'
    ) -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(30000)) {
        $process.Kill()
        throw 'GUI smoke check timed out.'
    }
    $process.Refresh()
    if ($process.ExitCode -ne 0) { throw "GUI exited with code $($process.ExitCode)." }
    if (-not (Test-Path -LiteralPath $screenshot)) { throw 'GUI screenshot was not created.' }
    Write-Host 'Bundled renderer and offline GUI check: OK'
    Write-Host "Screenshot: $screenshot"
} finally {
    $env:PATH = $previousPath
    $env:PYTHONHOME = $previousPythonHome
    $env:PYTHONPATH = $previousPythonPath
    $env:QT_QPA_PLATFORM = $previousQtPlatform
}
