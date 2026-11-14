@echo off
rem Double-click to install the PdfCraft Word bridge (output saved to install-log.txt).
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %* > "%~dp0install-log.txt" 2>&1
type "%~dp0install-log.txt"
timeout /t 10
