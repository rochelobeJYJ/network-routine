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

& $Python @pythonArgs -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --noconsole `
    --uac-admin `
    --name NetworkRoutine `
    --distpath .\release_final `
    --workpath .\build_final `
    --specpath .\build_final `
    .\network_routine.py
