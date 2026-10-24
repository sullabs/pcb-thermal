# Run an ansys/ script with Ansys' own Python and the packages in ansys\.venv.
#   .\ansys\run.ps1 serve.py <folder>
#   .\ansys\run.ps1 setup_study.py <folder>
param([Parameter(Mandatory)][string]$Script, [Parameter(ValueFromRemainingArguments)]$Rest)
$here = $PSScriptRoot
$root = $env:ANSYSEM_ROOT261
if (-not $root) { throw "ANSYSEM_ROOT261 is not set. Install AEDT 2026 R1 or set it to the AnsysEM folder." }
$py = Join-Path $root "commonfiles\CPython\3_10\winx64\Release\python\python.exe"
$env:PYTHONPATH = Join-Path $here ".venv\Lib\site-packages"
& $py -u (Join-Path $here $Script) @Rest
