@echo off
rem Starts everything for the page: AEDT (with its gRPC port, if it is not running yet), the local server, the browser.
rem   start.bat [folder]     default: ansys\example. Drag a board folder onto this file to open that one.
rem Close this window to stop the server. AEDT stays open.
setlocal
set "HERE=%~dp0"
if "%ANSYSEM_ROOT261%"=="" (
  echo ANSYSEM_ROOT261 is not set. Install AEDT 2026 R1, or set it to the AnsysEM folder.
  pause
  exit /b 1
)
set "FOLDER=%~1"
if "%FOLDER%"=="" set "FOLDER=%HERE%example"

tasklist /FI "IMAGENAME eq ansysedt.exe" | find /I "ansysedt.exe" >nul
if errorlevel 1 (
  echo Starting AEDT on gRPC port 50051...
  start "" "%ANSYSEM_ROOT261%\ansysedt.exe" -grpcsrv 50051
) else (
  echo AEDT is already running.
)

set "PYTHONPATH=%HERE%.venv\Lib\site-packages"
"%ANSYSEM_ROOT261%\commonfiles\CPython\3_10\winx64\Release\python\python.exe" -u "%HERE%serve.py" "%FOLDER%"
pause
