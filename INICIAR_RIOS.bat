@echo off
title RIOS — Restaurant Intelligence OS
color 0B
echo.
echo  =============================================
echo   RIOS — Restaurant Intelligence OS v1.9.1
echo   Iniciando servidor...
echo  =============================================
echo.

:: Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERRO: Python nao encontrado.
    echo  Instale em: https://www.python.org/downloads/
    echo.
    pause
    exit
)

:: Muda para a pasta do BAT
cd /d "%~dp0"

:: Encerra qualquer instancia anterior do servidor RIOS na porta 8765
echo  Encerrando instancias anteriores do RIOS...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8765" ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak >nul

:: Inicia o servidor
echo  Servidor iniciado em http://localhost:8765
echo  Abrindo navegador...
timeout /t 2 /nobreak >nul
start http://localhost:8765
python rios_server.py

pause
