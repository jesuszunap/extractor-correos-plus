```bat
@echo off
title Crear acceso directo - Extractor Correos +

REM ===== Rutas =====
set TARGET=%~dp0\scripts\ejecutar.bat
set ICON=%~dp0\icons\icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor Correos +.lnk

IF NOT EXIST "%TARGET%" (
    echo ERROR:
    echo No se encontró scripts\ejecutar.bat
    echo.
    pause
    exit /b
)

REM ===== Convertir rutas =====
for %%A in ("%TARGET%") do set "PTARGET=%%~fA"
for %%A in ("%ICON%") do set "PICON=%%~fA"
for %%A in ("%SHORTCUT_PATH%") do set "PSHORTCUT=%%~fA"

powershell -NoProfile -Command ^
    "$W = New-Object -ComObject WScript.Shell; " ^
    "$S = $W.CreateShortcut('%PSHORTCUT%'); " ^
    "$S.TargetPath = '%PTARGET%'; " ^
    "$S.WorkingDirectory = (Split-Path '%PTARGET%'); " ^
    "$S.IconLocation = '%PICON%'; " ^
    "$S.Save()"

cls
echo ============================================
echo     ACCESO DIRECTO CREADO CORRECTAMENTE
echo ============================================
echo.
echo Nombre:
echo Extractor Correos +
echo.

pause
```
