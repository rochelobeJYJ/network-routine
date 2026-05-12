param(
    [string]$Python = "py",
    [string]$PythonVersion = "-3.13"
)

$ErrorActionPreference = "Stop"

$pythonArgs = @()
if ($PythonVersion) {
    $pythonArgs += $PythonVersion
}

& $Python @pythonArgs -m pip install -r .\requirements-dev.txt

New-Item -ItemType Directory -Force -Path .\build_final | Out-Null
New-Item -ItemType Directory -Force -Path .\release_final | Out-Null

$buildStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$workPath = Join-Path .\build_final "pyinstaller_$buildStamp"
$specPath = Join-Path $workPath "spec"
$distPath = Join-Path $workPath "dist"
$outputExe = Join-Path .\release_final "NetworkRoutine.exe"
$builtExe = Join-Path $distPath "NetworkRoutine.exe"

New-Item -ItemType Directory -Force -Path $workPath | Out-Null
New-Item -ItemType Directory -Force -Path $specPath | Out-Null
New-Item -ItemType Directory -Force -Path $distPath | Out-Null

& $Python @pythonArgs -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --noconsole `
    --uac-admin `
    --name NetworkRoutine `
    --distpath $distPath `
    --workpath $workPath `
    --specpath $specPath `
    .\network_routine.py

for ($attempt = 0; $attempt -lt 10; $attempt++) {
    try {
        if (Test-Path $outputExe) {
            Remove-Item -LiteralPath $outputExe -Force
        }
        Copy-Item -LiteralPath $builtExe -Destination $outputExe -Force
        break
    }
    catch {
        if ($attempt -eq 9) {
            throw
        }
        Start-Sleep -Seconds 1
    }
}
