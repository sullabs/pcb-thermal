@echo off
rem Starts everything for the page: AEDT (with its gRPC port, if it is not running yet), the local server, the browser.
rem   start.bat [folder]     default: ansys\example. Drag a board folder onto this file to open that one.
rem Close this window to stop the server. AEDT stays open.
rem Uses the newest AEDT installed (ANSYSEM_ROOT252, ANSYSEM_ROOT261, ...); set AEDT_RELEASE=252 to pick one.
setlocal EnableDelayedExpansion
set "HERE=%~dp0"
set "REL=%AEDT_RELEASE%"
if "%REL%"=="" (
  rem "set" lists the variables sorted, so the last ANSYSEM_ROOTnnn is the newest release
  for /f "tokens=1 delims==" %%v in ('set ANSYSEM_ROOT 2^>nul') do (
    set "NAME=%%v"
    if "!NAME:~15!"=="" if not "!NAME:~14!"=="" set "REL=!NAME:~12,3!"
  )
)
set "ROOT="
if not "%REL%"=="" set "ROOT=!ANSYSEM_ROOT%REL%!"
if "%ROOT%"=="" (
  echo No ANSYSEM_ROOTxxx variable is set. Install AEDT ^(2025 R2 or later^), or set ANSYSEM_ROOT252 to the AnsysEM folder.
  pause
  exit /b 1
)
rem 252 -> 2025.2, read by the scripts
set "AEDT_VERSION=20%REL:~0,2%.%REL:~2,1%"
set "PY=%ROOT%\commonfiles\CPython\3_10\winx64\Release\python\python.exe"
if not exist "%PY%" for /d %%d in ("%ROOT%\commonfiles\CPython\3_*") do if exist "%%d\winx64\Release\python\python.exe" set "PY=%%d\winx64\Release\python\python.exe"
rem No Python in AEDT: a Python 3.10 installed from offline\python-3.10.11-amd64.exe
if not exist "%PY%" set "PY=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not exist "%PY%" set "PY=%ProgramFiles%\Python310\python.exe"
if not exist "%PY%" (
  echo No Python found under %ROOT%\commonfiles\CPython. Install offline\python-3.10.11-amd64.exe.
  pause
  exit /b 1
)
echo AEDT %AEDT_VERSION%, %ROOT%
set "FOLDER=%~1"
if "%FOLDER%"=="" set "FOLDER=%HERE%example"

tasklist /FI "IMAGENAME eq ansysedt.exe" | find /I "ansysedt.exe" >nul
if errorlevel 1 (
  echo Starting AEDT on gRPC port 50051...
  start "" "%ROOT%\ansysedt.exe" -grpcsrv 50051
) else (
  echo AEDT is already running.
)

set "PYTHONPATH=%HERE%.venv\Lib\site-packages"
rem First run on a PC: install requirements.txt into ansys\.venv (pip --target, the venv launcher itself is broken),
rem from offline\wheels when make_offline.ps1 made that folder, else from the internet.
set "SRC="
if exist "%HERE%offline\wheels" set "SRC=--no-index --find-links "%HERE%offline\wheels""
"%PY%" -c "import importlib.metadata as m; m.distribution('pyaedt')" 2>nul
if errorlevel 1 (
  echo PyAEDT is missing, installing requirements.txt into .venv...
  "%PY%" -m pip install --disable-pip-version-check --target "%PYTHONPATH%" %SRC% -r "%HERE%requirements.txt"
  if errorlevel 1 (
    echo The install failed. Without offline\wheels it needs internet access; see the pip output above.
    pause
    exit /b 1
  )
)
"%PY%" -u "%HERE%serve.py" "%FOLDER%"
pause
