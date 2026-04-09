; ============================================================
;  Makerspace Card Scanner - Inno Setup Script
;  Produces a professional Windows installer that bundles
;  embedded Python + all dependencies + application source.
;
;  Compile with:  ISCC.exe setup.iss
;  Or use:        python build_installer.py
; ============================================================

#define MyAppName       "Makerspace Card Scanner"
#define MyAppVersion    "1.5"
#define MyAppPublisher  "Clemson Makerspace"
#define MyAppURL        "https://github.com/clemsonMakerspace/Makerspace-CardScanner"
#define MyAppExeName    "MakerspaceScanner.bat"
#define MyAppIcon       "MakerspaceLogoIcon.ico"

[Setup]
AppId={{E5A2F3B1-9C4D-4E8A-B6F7-1A2D3E4F5G6H}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
DefaultDirName=C:\MakerspaceScanner
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Output settings
OutputDir=output
OutputBaseFilename=MakerspaceCardScanner_Setup_{#MyAppVersion}
; Compression
Compression=lzma2/ultra64
SolidCompression=yes
; Appearance
SetupIconFile=build\{#MyAppIcon}
WizardStyle=modern
WizardSizePercent=110
; Privileges -- user-level install (no admin needed for auto-updates)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Uninstall
UninstallDisplayIcon={app}\{#MyAppIcon}
UninstallDisplayName={#MyAppName}
; Misc
DisableProgramGroupPage=yes
; Allow upgrades over existing installation
UsePreviousAppDir=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked
Name: "autoupdate"; Description: "Enable automatic updates from GitHub (runs daily at midnight)"; GroupDescription: "Auto-updater:"; Flags: checked

[Files]
; --- Embedded Python runtime ---
Source: "build\python\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs

; --- Application source files (always overwritten on upgrade) ---
Source: "build\MakerspaceSignInTablet.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "build\CardReaderMakerspace.py";   DestDir: "{app}"; Flags: ignoreversion
Source: "build\database.py";               DestDir: "{app}"; Flags: ignoreversion
Source: "build\database_sync.py";          DestDir: "{app}"; Flags: ignoreversion
Source: "build\excel_db_sync.py";          DestDir: "{app}"; Flags: ignoreversion
Source: "build\excel_utils.py";            DestDir: "{app}"; Flags: ignoreversion
Source: "build\bridge_api.py";             DestDir: "{app}"; Flags: ignoreversion
Source: "build\config_examples.py";        DestDir: "{app}"; Flags: ignoreversion
Source: "build\auto_updater.py";           DestDir: "{app}"; Flags: ignoreversion
Source: "build\fetch_missing_training.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "build\fetch_missing_training.bat";DestDir: "{app}"; Flags: ignoreversion
Source: "build\test_scanner_simulation.py";DestDir: "{app}"; Flags: ignoreversion
Source: "build\MakerspaceScanner.bat";     DestDir: "{app}"; Flags: ignoreversion

; --- Image assets (always overwritten on upgrade) ---
; skipifsourcedoesntexist handles optional images gracefully
Source: "build\BackgroundTablet.png";       DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\BackgroundWatt.png";         DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\BackgroundAdobe.png";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\BackgroundTabletLaptop.png"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\background.png";             DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\backgroundWattScreen.png";   DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\background2.png";            DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\LogoBW.png";                 DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "build\MakerspaceLogoIcon.ico";     DestDir: "{app}"; Flags: ignoreversion

; --- Version tracker ---
Source: "build\.version"; DestDir: "{app}"; Flags: ignoreversion

; --- Config template (only on FIRST install -- never overwrite user config) ---
Source: "build\config_examples.py"; DestDir: "{app}"; DestName: "config.py"; Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
; Start Menu shortcut
Name: "{group}\{#MyAppName}";          Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"; WorkingDir: "{app}"
Name: "{group}\Fetch Missing Training"; Filename: "{app}\fetch_missing_training.bat"; WorkingDir: "{app}"
Name: "{group}\Check for Updates";      Filename: "{app}\python\python.exe"; Parameters: "auto_updater.py --verbose"; WorkingDir: "{app}"
Name: "{group}\Configuration Guide";    Filename: "{app}\config_examples.py"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

; Desktop shortcut (optional)
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"; WorkingDir: "{app}"; Tasks: desktopicon

[Dirs]
; Create backups directory
Name: "{app}\backups"; Flags: uninsneveruninstall

[Run]
; --- Create scheduled task for auto-updates ---
Filename: "schtasks.exe"; \
    Parameters: "/create /tn ""MakerspaceCardScanner_AutoUpdate"" /tr ""\""{app}\python\pythonw.exe\"" \""{app}\auto_updater.py\"""" /sc daily /st 00:00 /f"; \
    StatusMsg: "Setting up automatic updates..."; \
    Flags: runhidden nowait; \
    Tasks: autoupdate

; --- Optionally launch the application ---
Filename: "{app}\{#MyAppExeName}"; \
    Description: "Launch {#MyAppName}"; \
    Flags: nowait postinstall skipifsilent shellexec; \
    WorkingDir: "{app}"

[UninstallRun]
; Remove the scheduled task on uninstall
Filename: "schtasks.exe"; \
    Parameters: "/delete /tn ""MakerspaceCardScanner_AutoUpdate"" /f"; \
    Flags: runhidden

[UninstallDelete]
; Clean up generated files (but NOT user data)
Type: files; Name: "{app}\.version"
Type: files; Name: "{app}\update.log"
Type: files; Name: "{app}\update.log.old"
Type: files; Name: "{app}\*.update_backup"
Type: files; Name: "{app}\__pycache__\*"
Type: dirifempty; Name: "{app}\__pycache__"

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%nThe application will be installed to a user-writable directory so automatic updates from GitHub can be applied without administrator rights.%n%nClick Next to continue.

[Code]
// Pascal Script for custom installer logic

procedure CurStepChanged(CurStep: TSetupStep);
var
    ConfigFile: String;
    Msg: String;
begin
    if CurStep = ssPostInstall then
    begin
        ConfigFile := ExpandConstant('{app}\config.py');

        // Show first-run configuration reminder
        Msg := 'Installation complete!' + #13#10 + #13#10 +
               'IMPORTANT: Before first use, please edit:' + #13#10 +
               '  ' + ConfigFile + #13#10 + #13#10 +
               'Set your location (Watt/Cooper) and Bridge API credentials.' + #13#10 +
               'See config_examples.py for reference configurations.';
        MsgBox(Msg, mbInformation, MB_OK);
    end;
end;
