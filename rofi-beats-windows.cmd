@echo off
rem SPDX-License-Identifier: GPL-3.0-only
rem Rofi Beats - Windows launcher
setlocal
set "WINPS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0rofi-beats-windows.ps1"
set "TARGET=%~dp0windows\rofi-beats-windows.ps1"
set "STATIONS=%~dp0windows\stations.json"
set "FOREGROUND=0"
set "RESET="

if not exist "%TARGET%" goto missing_files
if not exist "%STATIONS%" goto missing_files

:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--foreground" (
  set "FOREGROUND=1"
  shift
  goto parse_args
)
if /I "%~1"=="--reset-profile" (
  set "RESET=-ResetProfile"
  shift
  goto parse_args
)
echo Unknown argument: %~1
echo Usage: rofi-beats-windows.cmd [--foreground] [--reset-profile]
exit /b 1

:args_done
if "%FOREGROUND%"=="1" goto run_foreground

if exist "%WINPS%" (
  start "" "%WINPS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%SCRIPT%" %RESET%
) else (
  start "" powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%SCRIPT%" %RESET%
)
exit /b 0

:run_foreground
if exist "%WINPS%" (
  "%WINPS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" %RESET%
) else (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" %RESET%
)
exit /b 0

:missing_files
echo Rofi Beats - Windows needs the full release package.
echo Please download and extract the .zip release so the "windows" folder stays next to this file.
pause
exit /b 1
