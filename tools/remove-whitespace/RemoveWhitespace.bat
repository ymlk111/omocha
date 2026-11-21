@echo off
rem Drop files onto this .bat to run the .ps1 with the same name.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1" %*
