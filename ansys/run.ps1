# Run an ansys/ script with Ansys' own Python and the packages in ansys\.venv.
#   .\ansys\run.ps1 serve.py <folder>
#   .\ansys\run.ps1 setup_study.py <folder>
# Uses the newest AEDT installed (ANSYSEM_ROOT252, ANSYSEM_ROOT261, ...); set AEDT_RELEASE=252 to pick one.
param([Parameter(Mandatory)][string]$Script, [Parameter(ValueFromRemainingArguments)]$Rest)
$here = $PSScriptRoot
$rel = $env:AEDT_RELEASE
if (-not $rel) {
    $rel = Get-ChildItem Env: | Where-Object Name -match '^ANSYSEM_ROOT\d{3}$' |
        ForEach-Object { $_.Name.Substring(12) } | Sort-Object | Select-Object -Last 1
}
$root = if ($rel) { [Environment]::GetEnvironmentVariable("ANSYSEM_ROOT$rel") }
if (-not $root) { throw "No ANSYSEM_ROOTxxx variable is set. Install AEDT (2025 R2 or later), or set ANSYSEM_ROOT252 to the AnsysEM folder." }
$env:AEDT_VERSION = "20$($rel.Substring(0, 2)).$($rel.Substring(2))"   # 252 -> 2025.2, read by the scripts
$py = Get-ChildItem (Join-Path $root "commonfiles\CPython\3_*\winx64\Release\python\python.exe") |
    Sort-Object { [version]($_.FullName -replace '.*CPython\\3_(\d+).*', '3.$1') } | Select-Object -Last 1 -ExpandProperty FullName
# No Python in AEDT: a Python 3.10 installed from offline\python-3.10.11-amd64.exe
if (-not $py) {
    $py = @("$env:LOCALAPPDATA\Programs\Python\Python310\python.exe", "$env:ProgramFiles\Python310\python.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $py) { throw "No Python found under $root\commonfiles\CPython. Install offline\python-3.10.11-amd64.exe." }
Write-Host "AEDT $env:AEDT_VERSION, $root"
$env:PYTHONPATH = Join-Path $here ".venv\Lib\site-packages"
# First run on a PC: install requirements.txt into ansys\.venv (pip --target, the venv launcher itself is broken),
# from offline\wheels when make_offline.ps1 made that folder, else from the internet.
$wheels = Join-Path $here "offline\wheels"
$src = if (Test-Path $wheels) { @("--no-index", "--find-links", $wheels) } else { @() }
& $py -c "import importlib.metadata as m; m.distribution('pyaedt')" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyAEDT is missing, installing requirements.txt into .venv..."
    & $py -m pip install --disable-pip-version-check --target $env:PYTHONPATH @src -r (Join-Path $here "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "The install failed. Without offline\wheels it needs internet access; see the pip output above." }
}
& $py -u (Join-Path $here $Script) @Rest
