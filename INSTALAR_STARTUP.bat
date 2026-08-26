@echo off
title RIOS — Instalador de Startup Automático
color 0B
echo.
echo  =====================================================
echo   RIOS — Configurar Inicio Automatico com o Windows
echo  =====================================================
echo.
echo  Este script vai configurar o RIOS para iniciar
echo  automaticamente toda vez que o Windows ligar.
echo.
echo  O servidor abrira em segundo plano (sem janela).
echo  O navegador abrira automaticamente em:
echo    http://localhost:8765
echo.
pause

:: Copia o VBS para a pasta de Startup do Windows
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SOURCE=%~dp0RIOS_Startup.vbs"
set "DEST=%STARTUP%\RIOS_Startup.vbs"

if not exist "%SOURCE%" (
    echo.
    echo  ERRO: RIOS_Startup.vbs nao encontrado na mesma pasta!
    echo  Certifique-se de que os arquivos estao juntos.
    pause
    exit
)

copy /Y "%SOURCE%" "%DEST%" >nul

if exist "%DEST%" (
    echo.
    echo  =====================================================
    echo   RIOS configurado com sucesso!
    echo.
    echo   O RIOS vai iniciar automaticamente com o Windows.
    echo   Para desinstalar: delete o arquivo
    echo   %DEST%
    echo  =====================================================
    echo.
    echo  Deseja iniciar o RIOS agora? (S/N)
    set /p resposta=  Resposta:
    if /i "%resposta%"=="S" (
        cscript //nologo "%SOURCE%"
    )
) else (
    echo.
    echo  ERRO: Nao foi possivel copiar o arquivo.
    echo  Tente executar como Administrador.
)

echo.
pause
