#define MyAppName "AutoTransfere"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SirSant"
#define MyAppExeName "AutoTransfere.exe"

[Setup]
AppId={{A45B8182-217A-4AA8-9E13-4DB7D4F46E5D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\\AutoTransfere
DefaultGroupName=AutoTransfere
UninstallDisplayIcon={app}\\{#MyAppExeName}
OutputDir=Output
OutputBaseFilename=AutoTransfere_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
CloseApplications=yes
RestartApplications=no

[Files]
Source: "dist\\AutoTransfere\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\\AutoTransfere"; Filename: "{app}\\{#MyAppExeName}"
Name: "{autodesktop}\\AutoTransfere"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos adicionais:"

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "Abrir AutoTransfere"; Flags: nowait postinstall skipifsilent
