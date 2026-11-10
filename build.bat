@echo off
rem Compila o Leitor de Projetos Pro em um unico .exe (dist\Leitor_Projetos.exe).
rem Usado pelo workflow .github\workflows\release-exe.yml (variavel CI) e tambem pode ser rodado a mao: pip install pyinstaller && build.bat
echo Iniciando a compilacao do Leitor de Projetos Pro...
echo.

python -m PyInstaller --noconfirm --clean --name "Leitor_Projetos" --onefile --windowed --add-data "static;static" --add-data "prompt_rede_eletrica.txt;." --add-data "data;data" --add-data "plugin_nanocad\appload\nanoCAD26;plugin_nanocad\appload\nanoCAD26" --hidden-import uvicorn.logging --hidden-import uvicorn.loops --hidden-import uvicorn.loops.auto --hidden-import uvicorn.protocols.http.auto --hidden-import uvicorn.protocols.websockets.auto --hidden-import uvicorn.lifespan.on --hidden-import uvicorn.lifespan.off --hidden-import fastapi --hidden-import webview --hidden-import openpyxl --hidden-import quickjs --hidden-import _quickjs --hidden-import httpx app.py
if errorlevel 1 (
    echo.
    echo [ERRO] A compilacao falhou.
    if not defined CI pause
    exit /b 1
)

echo.
echo Compilacao concluida! O executavel esta em "dist\Leitor_Projetos.exe".
if not defined CI pause
