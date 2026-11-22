@ECHO OFF
TITLE Device Cleanup Manager
ECHO Starting Device Cleanup Manager...
POWERSHELL.EXE -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0DeviceCleanupManager.ps1"
IF ERRORLEVEL 1 ECHO The script exited with an error.
PAUSE
