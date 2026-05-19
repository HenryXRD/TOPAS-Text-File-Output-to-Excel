@echo off
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "Batch-Phases.ps1" -FolderPath .

pause