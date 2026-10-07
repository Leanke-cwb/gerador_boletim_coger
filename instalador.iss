#define MyAppName "Gerador de Boletim Interno - COGER"
#define MyAppVersion "6.7"
#define MyAppPublisher "COGER"
#define MyAppExeName "Gerador_Boletim_COGER.exe"

[Setup]
AppId={{A427F6F3-6C84-4B41-9932-BC6C7CB87510}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; Instala somente para o usuário atual, sem exigir administrador
DefaultDirName={localappdata}\Programs\Gerador Boletim COGER
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

OutputDir=instalador
OutputBaseFilename=Instalador_Gerador_Boletim_COGER_v6_7

Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; Não solicita elevação
PrivilegesRequired=lowest

ArchitecturesAllowed=x64compatible

UninstallDisplayIcon={app}\{#MyAppExeName}

[Files]
Source: "dist\Gerador_Boletim_COGER\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\Gerador de Boletim COGER"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Gerador de Boletim COGER"; Flags: nowait postinstall skipifsilent
