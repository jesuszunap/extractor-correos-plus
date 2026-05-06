```bat
@echo off
title Extractor Correos +

REM ===== Ir a carpeta SRC =====
cd /d "%~dp0\..\src"

cls
echo ============================================
echo           EXTRACTOR CORREOS +
echo ============================================
echo.

REM ===== Verificar Python =====
python --version >nul 2>&1

IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python no esta instalado.
    echo.
    echo Instale Python desde la carpeta:
    echo utils
    echo.
    pause
    exit /b
)

echo Python detectado correctamente.
echo.

REM ===== Verificar personas.json =====
IF NOT EXIST personas.json (
    echo ERROR: No se encontro personas.json
    echo.
    echo Copie el archivo dentro de:
    echo src\
    echo.
    pause
    exit /b
)

REM ===== Instalar dependencias =====
cd /d "%~dp0\..\docs"

IF EXIST requirements.txt (
    echo Verificando dependencias...
    echo.
    python -m pip install -r requirements.txt >nul 2>&1
)

REM ===== Ejecutar programa =====
cd /d "%~dp0\..\src"

cls
python extractor_de_correos.py

echo.
pause
```
