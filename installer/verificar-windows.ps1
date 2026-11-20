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
Set-ItemProperty -LiteralPath $chave -Name DisplayVersion -Value '9999.0.0'
if ((Instalar-OrcaCore) -eq 0 -or (Get-ItemPropertyValue $chave DisplayVersion) -ne '9999.0.0') {
    throw 'O instalador deixou uma versão anterior substituir a mais recente.'
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
