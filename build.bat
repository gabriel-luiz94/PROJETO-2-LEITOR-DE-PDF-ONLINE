@echo off
echo Iniciando a compilação do Leitor de Projetos Pro...
echo.

C:\Users\gabriel.sales\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\Scripts\pyinstaller.exe --name "Leitor_Projetos" --onefile --windowed --add-data "static;static" --add-data "prompt_rede_eletrica.txt;." --add-data "data;data" --hidden-import uvicorn.logging --hidden-import uvicorn.loops --hidden-import uvicorn.loops.auto --hidden-import uvicorn.protocols.http.auto --hidden-import uvicorn.protocols.websockets.auto --hidden-import uvicorn.lifespan.on --hidden-import uvicorn.lifespan.off --hidden-import fastapi --hidden-import webview app.py

echo.
echo Compilação concluída! O executável está na pasta "dist".
pause
