# Run an ansys/ script with Ansys' own Python and the packages in ansys\.venv.
#   .\ansys\run.ps1 serve.py <folder>
#   .\ansys\run.ps1 setup_study.py <folder>
param([Parameter(Mandatory)][string]$Script, [Parameter(ValueFromRemainingArguments)]$Rest)
$here = $PSScriptRoot
$root = $env:ANSYSEM_ROOT261
if (-not $root) { throw "ANSYSEM_ROOT261 is not set. Install AEDT 2026 R1 or set it to the AnsysEM folder." }
$py = Join-Path $root "commonfiles\CPython\3_10\winx64\Release\python\python.exe"
$env:PYTHONPATH = Join-Path $here ".venv\Lib\site-packages"
# First run on a PC: install requirements.txt into ansys\.venv (pip --target, the venv launcher itself is broken).
& $py -c "import importlib.metadata as m; m.distribution('pyaedt')" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyAEDT is missing, installing requirements.txt into .venv..."
    & $py -m pip install --target $env:PYTHONPATH -r (Join-Path $here "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "The install failed. It needs internet access; see the pip output above." }
}
& $py -u (Join-Path $here $Script) @Rest
