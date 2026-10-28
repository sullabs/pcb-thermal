# Make a folder to install the page on a PC without internet: the ansys\ files, offline\wheels (every package in
# requirements.txt, for Python 3.10 on 64-bit Windows) and offline\python-3.10.11-amd64.exe, plus a zip of it.
#   .\ansys\make_offline.ps1 [out folder]     default: dist\pcb-thermal next to ansys\
# Copies the ansys\ files git tracks as they are on disk now, so uncommitted edits go along.
param([string]$Out = (Join-Path (Split-Path $PSScriptRoot) "dist\pcb-thermal"))
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$py = Get-ChildItem Env: | Where-Object Name -match '^ANSYSEM_ROOT\d{3}$' | Sort-Object Name | Select-Object -Last 1 |
    ForEach-Object { Join-Path $_.Value "commonfiles\CPython\3_10\winx64\Release\python\python.exe" }
if (-not $py -or -not (Test-Path $py)) { $py = "python" }   # pip download only needs some Python with pip

if (Test-Path $Out) { throw "$Out exists; move it away first." }
New-Item -ItemType Directory -Force (Join-Path $Out "offline\wheels") | Out-Null
git -C $here ls-files | ForEach-Object {
    $dst = Join-Path $Out $_
    New-Item -ItemType Directory -Force (Split-Path $dst) | Out-Null
    Copy-Item (Join-Path $here $_) $dst
}
# The sample board start.bat opens by default (git does not track it)
New-Item -ItemType Directory -Force (Join-Path $Out "example") | Out-Null
Copy-Item (Join-Path $here "example\power_board.aedb") (Join-Path $Out "example") -Recurse

& $py -m pip download --disable-pip-version-check -r (Join-Path $here "requirements.txt") -d (Join-Path $Out "offline\wheels") `
    --only-binary=:all: --platform win_amd64 --python-version 3.10 --implementation cp
if ($LASTEXITCODE -ne 0) { throw "pip download failed." }
$installer = Join-Path $Out "offline\python-3.10.11-amd64.exe"
Invoke-WebRequest https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe -OutFile $installer
if ((Get-AuthenticodeSignature $installer).Status -ne "Valid") { throw "The Python installer's signature is not valid." }

@"
PCB thermal page (Ansys AEDT route): install on a new PC

1. Install Ansys Electronics Desktop 2025 R2 or later. It brings its own Python 3.10, which the page uses.
   offline\python-3.10.11-amd64.exe is only needed if AEDT's commonfiles\CPython folder is missing:
   install it for the current user with the default options.
2. Copy this folder anywhere and double-click start.bat. The first run installs the Python packages from
   offline\wheels into .venv (no internet needed), starts AEDT on gRPC port 50051 and opens the page.
   With more than one AEDT installed it uses the newest; set AEDT_RELEASE=252 to force 2025 R2.
3. Drag a board folder (with its .aedb or Altium .edb) onto start.bat to open that board.
"@ | Set-Content (Join-Path $Out "INSTALL.txt")

Compress-Archive -Path $Out -DestinationPath "$Out.zip" -Force
"{0}`n{1}.zip, {2:N0} MB" -f $Out, $Out, ((Get-Item "$Out.zip").Length / 1MB)
