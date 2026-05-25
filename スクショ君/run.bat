```bat
@echo off
cd /d %~dp0

powershell -ExecutionPolicy Bypass -File .\screenshot.ps1
pause
```
