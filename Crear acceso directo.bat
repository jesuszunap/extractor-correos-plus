@echo off
title Crear acceso directo - Extractor Correos +

REM ===== Rutas =====
set "TARGET=%~dp0scripts\ejecutar.bat"
set "LAUNCHER=%~dp0scripts\ejecutar_oculto.vbs"
set "ICON=%~dp0icons\icono_mail.ico"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor Correos +.lnk"
set "WSCRIPT=%SystemRoot%\System32\wscript.exe"

IF NOT EXIST "%TARGET%" (
    echo ERROR:
    echo No se encontro scripts\ejecutar.bat
    echo.
    pause
    exit /b
)

IF NOT EXIST "%LAUNCHER%" (
    echo ERROR:
    echo No se encontro scripts\ejecutar_oculto.vbs
    echo.
    pause
    exit /b
)

REM ===== Convertir rutas =====
for %%A in ("%LAUNCHER%") do set "PLAUNCHER=%%~fA"
for %%A in ("%ICON%") do set "PICON=%%~fA"
for %%A in ("%SHORTCUT_PATH%") do set "PSHORTCUT=%%~fA"
for %%A in ("%WSCRIPT%") do set "PWSCRIPT=%%~fA"

powershell -NoProfile -Command ^
    "$W = New-Object -ComObject WScript.Shell; " ^
    "$S = $W.CreateShortcut('%PSHORTCUT%'); " ^
    "$S.TargetPath = '%PWSCRIPT%'; " ^
    "$S.Arguments = '""%PLAUNCHER%""'; " ^
    "$S.WorkingDirectory = (Split-Path '%PLAUNCHER%'); " ^
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
echo El acceso directo abrira solo la ventana del programa.
echo.

pause
