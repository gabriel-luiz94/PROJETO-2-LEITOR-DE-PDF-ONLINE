@echo off
rem ==========================================================================
rem  Compila o plugin para o nanoCAD 5 (.NET Framework 4) NESTE computador,
rem  usando o compilador C# que ja vem no Windows e as DLLs do nanoCAD 5
rem  instalado (hostmgd.dll / hostdbmgd.dll). Nao precisa de Visual Studio.
rem
rem  Uso:  build_nanocad5.bat                      (procura o nanoCAD 5)
rem        build_nanocad5.bat "C:\Program Files\Nanosoft\nanoCAD 5.0"
rem  Saida: appload\nanoCAD5\OrcamentoNanoCAD.dll  (carregar com APPLOAD)
rem ==========================================================================
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "NCAD=%~1"
if "%NCAD%"=="" (
  for %%P in ("%ProgramFiles%\Nanosoft" "%ProgramFiles(x86)%\Nanosoft") do (
    if exist "%%~P" (
      for /d %%D in ("%%~P\nanoCAD*5*") do (
        if exist "%%~D\hostmgd.dll" set "NCAD=%%~D"
        if exist "%%~D\bin\hostmgd.dll" set "NCAD=%%~D\bin"
      )
    )
  )
)
if "%NCAD%"=="" (
  echo Nao encontrei o nanoCAD 5. Rode de novo informando a pasta que tem hostmgd.dll:
  echo   build_nanocad5.bat "C:\Program Files\Nanosoft\nanoCAD 5.0"
  exit /b 1
)
if exist "%NCAD%\bin\hostmgd.dll" set "NCAD=%NCAD%\bin"
if not exist "%NCAD%\hostmgd.dll" (
  echo Pasta sem hostmgd.dll: "%NCAD%"
  exit /b 1
)
echo nanoCAD: %NCAD%

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo Nao encontrei o csc.exe do .NET Framework 4.
  exit /b 1
)

set "DATA=..\data"
set "RES=/resource:%DATA%\regras_leitor_processamento_pernambuco_seed.json,regras.081.processamento.json"
set "RES=%RES% /resource:%DATA%\regras_leitor_classificacao_pernambuco_seed.json,regras.081.classificacao.json"
set "RES=%RES% /resource:%DATA%\vinculacao_pernambuco_seed.json,regras.081.vinculacao.json"
set "RES=%RES% /resource:%DATA%\regras_leitor_processamento_seed.json,regras.padrao.processamento.json"
set "RES=%RES% /resource:%DATA%\regras_leitor_classificacao_seed.json,regras.padrao.classificacao.json"
set "REFS=/r:"%NCAD%\hostmgd.dll" /r:"%NCAD%\hostdbmgd.dll" /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll"

if not exist appload\nanoCAD5 mkdir appload\nanoCAD5
set "SAIDA=appload\nanoCAD5\OrcamentoNanoCAD.dll"
set "OPCOES=/nologo /target:library /platform:anycpu /optimize+ /codepage:65001 /out:%SAIDA%"

echo Compilando...
"%CSC%" %OPCOES% %REFS% %RES% src\Nucleo\*.cs src\Plugin\*.cs > appload\nanoCAD5\compilacao.log 2>&1
if errorlevel 1 (
  echo Primeira tentativa falhou; tentando sem graficos transitorios (SEM_TRANSIENTES^)...
  "%CSC%" %OPCOES% /define:SEM_TRANSIENTES %REFS% %RES% src\Nucleo\*.cs src\Plugin\*.cs > appload\nanoCAD5\compilacao_sem_transientes.log 2>&1
  if errorlevel 1 (
    echo FALHOU. Mande os arquivos appload\nanoCAD5\compilacao*.log para ajuste.
    type appload\nanoCAD5\compilacao_sem_transientes.log
    exit /b 1
  )
)
echo.
echo OK: %SAIDA%
echo No nanoCAD 5: APPLOAD (ou NETLOAD) e escolha esse arquivo. Depois: ORCAMENTO
endlocal
