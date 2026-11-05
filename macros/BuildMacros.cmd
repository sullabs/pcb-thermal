@echo off
rem Builds PCBThermal.swp from the .bas files: imports, compiles, saves. Starts SolidWorks when it is not running.
rem Result in %TEMP%\PCBThermal_build.log. Option -nopause: no "Press any key" at the end.
cscript //nologo "%~dp0BuildMacros.vbs"
set rc=%errorlevel%
if not "%1"=="-nopause" pause
exit /b %rc%
