@echo off
REM ============================================================
REM  CNextCaseNotificationInstall.bat
REM  Double-click launcher for the Case Notification installer.
REM  Runs the PowerShell installer that sits next to this file,
REM  bypassing the execution policy so no manual step is needed.
REM ============================================================

title CNext Case Notification - Setup

echo Starting CNext Case Notification setup...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0CNextCaseNotificationInstall.ps1"

echo.
echo Setup finished. You can close this window.
pause
