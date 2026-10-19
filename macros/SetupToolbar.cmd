@echo off
rem Makes the PCB Thermal tab in SolidWorks. Close SolidWorks first. Options are passed on, for example: SetupToolbar.cmd -Version 2025
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0SetupToolbar.ps1" %*
pause
