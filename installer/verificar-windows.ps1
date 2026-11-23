# Ensaio de instalação para um perfil descartável no runner do GitHub Actions.
$ErrorActionPreference = 'Stop'
$setup = (Resolve-Path 'dist\OrcaCore-Setup.exe').Path
$pasta = Join-Path $env:LOCALAPPDATA 'Programs\OrcaCore'
$exe = Join-Path $pasta 'OrcaCore.exe'
$chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\OrcaCore_is1'
$chaveMaquina = $chave.Replace('HKCU:', 'HKLM:')
$nome = 'Or' + [char]0xE7 + 'aCore'
$atalho = Join-Path ([Environment]::GetFolderPath('Programs')) "$nome.lnk"
$versao = $env:ORCACORE_VERSION
$argumentos = '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'

$mt = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\mt.exe" |
    Sort-Object FullName -Descending | Select-Object -First 1
if (-not $mt) { throw 'mt.exe do Windows SDK ausente no runner.' }

function Confirmar-SemElevacao($executavel) {
    $manifesto = Join-Path $env:RUNNER_TEMP 'orcacore-manifest.xml'
    & $mt.FullName -nologo "-inputresource:$executavel;#1" "-out:$manifesto"
    if ($LASTEXITCODE -ne 0) { throw "Falha ao extrair manifesto de $executavel." }
    [xml]$xml = Get-Content -LiteralPath $manifesto -Raw
    $nivel = $xml.SelectSingleNode("//*[local-name()='requestedExecutionLevel']")
    if (-not $nivel -or $nivel.level -ne 'asInvoker') {
        throw "Manifesto de $executavel sem asInvoker."
    }
}

function Instalar-OrcaCore {
    $log = Join-Path $env:RUNNER_TEMP 'orcacore-setup.log'
    $resultado = Start-Process -FilePath $setup -ArgumentList ($argumentos + ('/LOG="{0}"' -f $log)) -Wait -PassThru
    Get-Content -LiteralPath $log -Tail 12 | ForEach-Object { Write-Host $_ }
    return $resultado.ExitCode
}

Confirmar-SemElevacao $setup
$arquivo = (Get-Item $setup).VersionInfo
$versaoArquivo = "$($arquivo.FileMajorPart).$($arquivo.FileMinorPart).$($arquivo.FileBuildPart)"
if ($versaoArquivo -ne $versao) {
    throw 'A versão do instalador difere da versão usada no build.'
}
if ((Test-Path $chave) -or (Test-Path $pasta)) {
    throw 'O ensaio precisa de um perfil sem instalação anterior.'
}
if ((Instalar-OrcaCore) -ne 0) { throw 'Falha na instalação.' }
$registro = Get-ItemProperty -LiteralPath $chave
if ($registro.DisplayName -notlike "$nome*" -or $registro.DisplayVersion -ne $versao -or
    $registro.InstallLocation.TrimEnd('\') -ne $pasta -or (Test-Path $chaveMaquina)) {
    throw 'O registro da instalação por usuário está incorreto.'
}
if ((Get-FileHash $exe).Hash -ne (Get-FileHash 'dist\OrcaCore.exe').Hash) {
    throw 'O executável instalado difere do portátil desta versão.'
}
$shell = New-Object -ComObject WScript.Shell
if (-not (Test-Path $atalho) -or $shell.CreateShortcut($atalho).TargetPath -ne $exe) {
    throw 'O atalho do menu Iniciar não aponta para o executável instalado.'
}

# Chama o trecho que roda na abertura do executável, com o registro real deste perfil.
Set-ItemProperty -LiteralPath $chave -Name DisplayVersion -Value '0.0.0'
$env:ORCACORE_INSTALADO_TESTE = $exe
$env:APP_MODE = 'desktop'
@'
import os
from services import auto_updater

auto_updater.IS_FROZEN = True
auto_updater.sys.executable = os.environ["ORCACORE_INSTALADO_TESTE"]
auto_updater.sincronizar_versao_instalada()
'@ | python -
if ($LASTEXITCODE -ne 0 -or (Get-ItemPropertyValue $chave DisplayVersion) -ne $versao) {
    throw 'A abertura após o auto-update não sincronizou o registro.'
}

$dados = Join-Path $pasta 'banco_resumo.db'
Set-Content -LiteralPath $dados -Value 'dados do usuario'
$hashDados = (Get-FileHash $dados).Hash
foreach ($maisNova in '9999.0.0', '65534.0.0') {
    Set-ItemProperty -LiteralPath $chave -Name DisplayVersion -Value $maisNova
    if ((Instalar-OrcaCore) -eq 0 -or (Get-ItemPropertyValue $chave DisplayVersion) -ne $maisNova) {
        throw "O instalador deixou uma versão anterior substituir $maisNova."
    }
}
Set-ItemProperty -LiteralPath $chave -Name DisplayVersion -Value $versao
if ((Instalar-OrcaCore) -ne 0 -or (Get-FileHash $dados).Hash -ne $hashDados) {
    throw 'A reinstalação não preservou o banco local.'
}

$desinstalador = [regex]::Match($registro.UninstallString, '^"([^"]+)"').Groups[1].Value
Confirmar-SemElevacao $desinstalador
if ((Start-Process -FilePath $desinstalador -ArgumentList $argumentos -Wait -PassThru).ExitCode -ne 0 -or
    (Test-Path $chave) -or (Test-Path $exe) -or (Test-Path $atalho) -or (Get-FileHash $dados).Hash -ne $hashDados) {
    throw 'A desinstalação não removeu a instalação corretamente ou alterou dados locais.'
}
Write-Host 'Manifestos, instalação por usuário, versão, atalho e preservação dos dados verificados.'

# Ensaia o Setup privado com dados sintéticos fora da pasta dos artefatos da release.
Remove-Item -LiteralPath $dados
$amostra = Join-Path $env:RUNNER_TEMP 'instalador-privado-teste'
New-Item -ItemType Directory -Path $amostra | Out-Null
$exeAmostra = Join-Path $amostra 'OrcaCore.exe'
$envAmostra = Join-Path $amostra '.env'
$bancoAmostra = Join-Path $amostra 'banco_resumo.db'
$saidaPrivada = Join-Path $amostra 'saida'
Set-Content -LiteralPath $exeAmostra -Value 'Executavel sintetico' -Encoding ascii
Set-Content -LiteralPath $envAmostra -Value @(
    'SUPABASE_URL=https://exemplo.supabase.co', 'SUPABASE_KEY=chave-sintetica'
) -Encoding ascii
$env:ORCACORE_BANCO_AMOSTRA = $bancoAmostra
@'
import os
import sqlite3

with sqlite3.connect(os.environ["ORCACORE_BANCO_AMOSTRA"]) as banco:
    banco.executescript("""
        CREATE TABLE usuarios_locais (email TEXT);
        CREATE TABLE sessoes (token TEXT);
        CREATE TABLE tabela_orcamento (ativo TEXT);
        INSERT INTO tabela_orcamento VALUES ('N3');
    """)
'@ | python -
if ($LASTEXITCODE -ne 0) { throw 'Falha ao criar o banco de teste.' }
python scripts/compilar_instalador_privado.py --executavel $exeAmostra --env $envAmostra --banco $bancoAmostra --versao $versao --saida $saidaPrivada
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o instalador privado.' }

$setup = Join-Path $saidaPrivada 'OrcaCore-Setup-Privado.exe'
Confirmar-SemElevacao $setup
$envInstalado = Join-Path $pasta '.env'
if ((Instalar-OrcaCore) -ne 0 -or (Get-FileHash $envInstalado).Hash -ne (Get-FileHash $envAmostra).Hash -or
    -not (Test-Path $dados) -or (Get-FileHash $exe).Hash -ne (Get-FileHash $exeAmostra).Hash) {
    throw 'O instalador privado não colocou os três arquivos no diretório do programa.'
}
Set-Content -LiteralPath $envInstalado -Value 'configuracao do usuario'
Set-Content -LiteralPath $dados -Value 'banco do usuario'
$hashEnvPrivado = (Get-FileHash $envInstalado).Hash
$hashBancoPrivado = (Get-FileHash $dados).Hash
if ((Instalar-OrcaCore) -ne 0 -or (Get-FileHash $envInstalado).Hash -ne $hashEnvPrivado -or
    (Get-FileHash $dados).Hash -ne $hashBancoPrivado) {
    throw 'O instalador privado sobrescreveu arquivos existentes.'
}
$registroPrivado = Get-ItemProperty -LiteralPath $chave
$desinstaladorPrivado = [regex]::Match($registroPrivado.UninstallString, '^"([^"]+)"').Groups[1].Value
if ((Start-Process -FilePath $desinstaladorPrivado -ArgumentList $argumentos -Wait -PassThru).ExitCode -ne 0 -or
    (Get-FileHash $envInstalado).Hash -ne $hashEnvPrivado -or (Get-FileHash $dados).Hash -ne $hashBancoPrivado) {
    throw 'A desinstalação privada removeu dados do usuário.'
}
Write-Host 'Instalação privada e preservação dos dados verificadas.'
