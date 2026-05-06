@echo off

REM --- Ir a la carpeta donde está el .py ---
cd /d "%~dp0\..\src"

REM --- Actualizar desde Github ---
git pull

cls

REM --- Verificar si Python está instalado ---
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python no esta instalado en esta PC.
    echo Por favor instala Python desde la carpeta Utils.
    pause
    exit /b
)

@REM echo.
echo Python detectado correctamente.
echo.

cd /d "%~dp0\..\docs"
REM --- Instalar dependencias ---
echo Instalando dependencias . . .
echo.
@REM IF EXIST requirements.txt (
@REM     python -m pip install --upgrade pip >nul
@REM     python -m pip install -r requirements.txt
@REM )

cls

cd /d "%~dp0\..\src"
REM --- Ejecutar el script principal ---
python extractor_de_correos.py

pause



