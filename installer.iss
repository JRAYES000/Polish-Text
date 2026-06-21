; ============================================================================
;  Installeur Inno Setup pour TextEnhancer AI
;  - Installation PAR UTILISATEUR (pas d'admin/UAC) dans %LOCALAPPDATA%\Programs
;  - Raccourci Bureau + menu Démarrer
;  - Désinstalleur visible dans « Programmes et fonctionnalités »
;  - L'install par-utilisateur garde l'auto-update fonctionnel
;  Version passée à la compilation : ISCC /DMyAppVersion=1.4.0 installer.iss
; ============================================================================

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#define MyAppName "TextEnhancer AI"
#define MyAppExeName "TextEnhancerAI.exe"
#define MyAppPublisher "Julien Rayes"
#define MyAppURL "https://github.com/JRAYES000/Polish-Text"

[Setup]
; AppId identifie l'application de façon unique (NE PAS changer entre versions).
AppId={{B7E2C3A4-9F1D-4E62-8C7A-3D5B6F210E94}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\TextEnhancerAI
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=installer_output
OutputBaseFilename=TextEnhancerAI-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64
; Ferme l'app si elle tourne avant de remplacer les fichiers (mise à jour).
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; \
  GroupDescription: "Raccourcis :"

[Files]
Source: "dist\TextEnhancerAI.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\TextEnhancer AI"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\TextEnhancer AI"; Filename: "{app}\{#MyAppExeName}"; \
  Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Lancer TextEnhancer AI"; \
  Flags: nowait postinstall skipifsilent
