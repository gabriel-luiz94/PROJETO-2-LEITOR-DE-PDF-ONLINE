@echo off
rem Recompila a DLL do nanoCAD 26 (precisa do .NET SDK 6+ instalado: https://dotnet.microsoft.com/download).
rem A DLL pronta ja vem versionada em appload\nanoCAD26\ -- so use isto se mudar o codigo.
cd /d "%~dp0"
dotnet build src\OrcamentoNanoCAD.csproj -c Release || exit /b 1
echo OK: appload\nanoCAD26\OrcamentoNanoCAD.dll
