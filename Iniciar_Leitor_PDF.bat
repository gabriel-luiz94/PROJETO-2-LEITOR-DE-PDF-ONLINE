@echo off
title Leitor PDF Online - Servidor Local

:: Vai para a pasta do projeto
cd /d "%~dp0"

:: Verifica se o Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado! Instale o Python e tente novamente.
    pause
    exit /b
)

:: Instala dependencias se necessario
echo [INFO] Verificando dependencias...
pip install -r requirements.txt --quiet

:: Mata qualquer processo uvicorn anterior na porta 8000
echo [INFO] Liberando porta 8000...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000 "') do (
    taskkill /PID %%a /F >nul 2>&1
)

echo.
echo ============================================
echo   Leitor PDF Online - Iniciando servidor...
echo   Acesse: http://localhost:8000
echo   Pressione CTRL+C para encerrar
echo ============================================
echo.

:: Inicia o servidor (abre o browser automaticamente via app.py)
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload

pause
