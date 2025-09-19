@echo off
ECHO.
ECHO --- Este script instala as dependencias para o projeto de envio de e-mails. ---
ECHO.

REM --- Verifica se o Python esta instalado ---
python --version >NUL 2>&1
IF %ERRORLEVEL% NEQ 0 (
    ECHO.
    ECHO ERRO: Python nao foi encontrado. Por favor, instale o Python 3.x a partir de python.org e tente novamente.
    ECHO.
    pause
    exit /b 1
)

REM --- Verifica e instala o pip se necessario ---
python -m pip --version >NUL 2>&1
IF %ERRORLEVEL% NEQ 0 (
    ECHO.
    ECHO pip nao foi encontrado. Tentando instalar o pip...
    python -m ensurepip --upgrade
    IF %ERRORLEVEL% NEQ 0 (
        ECHO.
        ECHO ERRO: Erro ao instalar o pip. Por favor, verifique sua instalacao do Python.
        ECHO.
        pause
        exit /b 1
    )
)

ECHO.
ECHO Verificando e instalando as bibliotecas necessarias...
ECHO.

REM --- Lista de bibliotecas a serem instaladas ---
SET "LIBRARIES=pywin32"

REM --- Instala as bibliotecas ---
python -m pip install %LIBRARIES%

IF %ERRORLEVEL% EQU 0 (
    ECHO.
    ECHO ✅ Todas as bibliotecas foram instaladas com sucesso!
    ECHO.
    ECHO Iniciando o programa 'import.exe'...
    ECHO.
    REM --- Executa o programa principal ---
    start import.exe
) ELSE (
    ECHO.
    ECHO ❌ Ocorreu um erro durante a instalacao das bibliotecas.
    ECHO Por favor, verifique o erro acima e tente novamente.
)

ECHO.
ECHO Processo concluido. Pressione qualquer tecla para sair...
pause >NUL
