@echo off
setlocal
if not exist "%~dp0scripts\office-check-windows.ps1" (
  echo Please extract the complete OfficeBridge ZIP before opening this launcher.
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0scripts\office-check-windows.ps1"
