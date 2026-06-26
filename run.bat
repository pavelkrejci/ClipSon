@echo off
setlocal

echo Current directory: %CD%

powershell.exe -NoProfile -Command "Get-ChildItem -Path . -Filter '*.ps1' -File -Recurse | Unblock-File"
powershell.exe -NoProfile -File ".\\clipson.ps1"

pause
