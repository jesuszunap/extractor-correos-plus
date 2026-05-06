@echo off

REM ====== Rutas ======
set TARGET=%~dp0\scripts\ejecutar.bat
set ICON=%~dp0\icons\icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor de Correos.lnk

@REM Convertir rutas a formato PowerShell seguro
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

@echo Acceso directo creado correctamente.
@echo.

pause
