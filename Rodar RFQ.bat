@echo off
chcp 65001 > nul
title Automacao RFQ
cd /d "%~dp0"

set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"

echo ==================================================
echo             INICIANDO AUTOMACAO RFQ
echo ==================================================
echo.

"%PY_EXE%" "RFQ.py"
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Ocorreu um erro ao executar a automacao RFQ.
    echo.
    pause
)
