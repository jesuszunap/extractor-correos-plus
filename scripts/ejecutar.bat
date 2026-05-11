@echo off
title Extractor Correos +

cd /d "%~dp0\..\src"

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "%CD%\menu_principal.py"
    exit /b
)

where pyw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pyw.exe "%CD%\menu_principal.py"
    exit /b
)

start "" python "%CD%\menu_principal.py"
exit /b
