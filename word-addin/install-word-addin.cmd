@echo off
rem Installs PdfCraft for Word: copies the template into Word's STARTUP folder so the PdfCraft tab
rem appears every time Word starts. Close Word first, or restart it afterwards.
rem Uninstall: delete %APPDATA%\Microsoft\Word\STARTUP\PdfCraft for Word.dotm
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%~dp0PdfCraft for Word.dotm" "%DEST%\" > "%~dp0install-word-addin-log.txt" 2>&1
if errorlevel 1 (echo Install failed. >> "%~dp0install-word-addin-log.txt") else (echo Installed to %DEST% >> "%~dp0install-word-addin-log.txt")
type "%~dp0install-word-addin-log.txt"
timeout /t 8
