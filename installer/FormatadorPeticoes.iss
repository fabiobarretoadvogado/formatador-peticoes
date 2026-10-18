#ifndef AppVersion
  #error AppVersion deve ser fornecida pelo processo de compilação.
#endif
#ifndef SourceDir
  #error SourceDir deve apontar para o aplicativo compilado.
#endif
#ifndef OutputDir
  #define OutputDir "..\dist"
#endif

[Setup]
AppId={{E8CC746E-6040-4AA3-9649-C42D3BBD62D8}
AppName=Formatador de Petições Iniciais
AppVersion={#AppVersion}
AppPublisher=Fabio Barreto Advocacia
AppPublisherURL=https://github.com/fabiobarretoadvogado/formatador-peticoes
AppUpdatesURL=https://github.com/fabiobarretoadvogado/formatador-peticoes/releases/latest
DefaultDirName={localappdata}\Programs\FormatadorPeticoes
DefaultGroupName=Formatador de Petições Iniciais
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=FormatadorPeticoes-Setup
SetupIconFile=..\assets\aplicativo.ico
UninstallDisplayIcon={app}\FormatadorPeticoes.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
DisableProgramGroupPage=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Formatador de Petições Iniciais"; Filename: "{app}\FormatadorPeticoes.exe"
Name: "{autodesktop}\Formatador de Petições Iniciais"; Filename: "{app}\FormatadorPeticoes.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\FormatadorPeticoes.exe"; Description: "Abrir o Formatador de Petições Iniciais"; Flags: nowait postinstall skipifsilent
