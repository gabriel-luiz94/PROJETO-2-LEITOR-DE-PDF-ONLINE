#ifndef AppVersion
  #error Passe /DAppVersion=X.Y.Z para compilar o instalador.
#endif

#define AppId "OrcaCore"
#define AppExe "OrcaCore.exe"
#define ChaveDesinstalacao "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + AppId + "_is1"

[Setup]
AppId={#AppId}
AppName=OrçaCore
AppVersion={#AppVersion}
AppPublisher=Gabriel Sales
AppPublisherURL=https://github.com/gabriel-luiz94/PROJETO-2-LEITOR-DE-PDF-ONLINE
DefaultDirName={localappdata}\Programs\OrcaCore
DisableDirPage=yes
DisableProgramGroupPage=yes
UsePreviousAppDir=no
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
UninstallDisplayIcon={app}\{#AppExe}
SetupIconFile=..\static\gridr.ico
OutputDir=..\dist
OutputBaseFilename=OrcaCore-Setup
VersionInfoVersion={#AppVersion}
Compression=lzma2/fast
SolidCompression=yes
CloseApplications=no

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "..\dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\OrçaCore"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir o OrçaCore"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
var
  VersaoInstalada, Pasta: String;
  Instalada, Pacote: Int64;
  MaisNova: Boolean;
begin
  Result := True;
  if not RegQueryStringValue(HKCU, '{#ChaveDesinstalacao}', 'InstallLocation', Pasta) then
    Exit;
  if not FileExists(AddBackslash(Pasta) + '{#AppExe}') then
    Exit;
  if not RegQueryStringValue(HKCU, '{#ChaveDesinstalacao}', 'DisplayVersion', VersaoInstalada) then
    Exit;
  Log('Versão registrada: ' + VersaoInstalada + '; versão do pacote: {#AppVersion}');
  if not StrToVersion(VersaoInstalada, Instalada) then
    Exit;
  if not StrToVersion('{#AppVersion}', Pacote) then
    Exit;
  { Inno Setup 6.7 compara Int64 com sinal, mas os componentes da versão são sem sinal. }
  if (Instalada < 0) <> (Pacote < 0) then
    MaisNova := Instalada < 0
  else
    MaisNova := ComparePackedVersion(Instalada, Pacote) > 0;
  if MaisNova then
  begin
    SuppressibleMsgBox('Uma versão mais recente do OrçaCore já está instalada.', mbInformation, MB_OK, IDOK);
    Result := False;
  end;
end;
