@echo off
chcp 65001 > nul
title Automacao Schulz Pedidos
cd /d "%~dp0..\Schulz"

set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"

echo ==================================================
echo           INICIANDO AUTOMACAO SCHULZ
echo ==================================================
echo.

"%PY_EXE%" "Schulz Pedidos.py"
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Ocorreu um erro ao executar a automacao Schulz.
    echo.
    pause
)
