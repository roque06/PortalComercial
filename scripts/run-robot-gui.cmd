@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "GUI_SCRIPT=%SCRIPT_DIR%run-robot-gui.ps1"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%GUI_SCRIPT%" (
  echo [ERROR] No se encontro el archivo:
  echo         %GUI_SCRIPT%
  echo.
  pause
  exit /b 1
)

if not exist "%PS_EXE%" (
  set "PS_EXE=powershell.exe"
)

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%GUI_SCRIPT%"
set "RC=%ERRORLEVEL%"
if "%RC%"=="5" set "RC=0"

if not "%RC%"=="0" (
  echo.
  echo [ERROR] No se pudo abrir el runner. ExitCode=%RC%
  echo.
  pause
)

endlocal & exit /b %RC%
