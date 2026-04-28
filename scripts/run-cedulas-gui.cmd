@echo off
setlocal
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-cedulas-gui-v2.ps1"
endlocal
