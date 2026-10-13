; DocPilot - Windows installer
;
; Build with:  tools\build_installer.bat        (rebuilds the app, then this)
; or directly: ISCC.exe installer\DocPilot.iss
;
; PER-USER BY DESIGN. Two reasons, both load-bearing:
;   1. PW machines rarely grant admin, and a UAC prompt the client cannot
;      answer is a dead end.
;   2. The app writes its job scratch under its own folder, so the install
;      directory MUST stay writable by the person running it. Program Files
;      is not. PrivilegesRequired=lowest puts it in %LOCALAPPDATA%\Programs.

#define AppName        "DocPilot"
#define AppVersion     "2.0-beta"
#define AppPublisher   "Physics Wallah - OnlyIAS"
#define AppExe         "DocPilot.exe"
#define SrcDir         "..\dist\DocPilot-2.0-beta"
; must match APP_MUTEX in desktop_launcher.py
#define AppMutex       "DocPilot_PW_OnlyIAS_Running"

[Setup]
; never change AppId - it is how Windows recognises an upgrade of this app
AppId={{A2EB0E30-D9AF-4CC0-BFC5-7C34F06720F0}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DirExistsWarning=no
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=DocPilot-Setup-{#AppVersion}
SetupIconFile=..\assets\icon\docpilot.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName} {#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; The app is a running web server, so its files WILL be locked on an upgrade.
; AppMutex is how that is handled: the launcher publishes this mutex, so Setup
; and the uninstaller detect a running DocPilot up front and ask for it to be
; closed by name.
;
; CloseApplications stays OFF, deliberately. Turning it on hands the job to
; Restart Manager, which (a) only *asks* an app to close itself - a request a
; Python/Tk process never answers, producing the "Setup was unable to
; automatically close all applications" dialog - and (b) with force= would
; terminate whatever else RM believes holds a file, up to and including the
; user's browser. Nothing but our own files live in {app}, and the mutex
; already covers our own app, so there is nothing left for RM to do.
AppMutex={#AppMutex}
CloseApplications=no
RestartApplications=no
VersionInfoVersion=2.0.0.0
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}";                        Filename: "{app}\{#AppExe}"; IconFilename: "{app}\{#AppExe}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}";  Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";                  Filename: "{app}\{#AppExe}"; IconFilename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; job scratch the app creates at runtime - Setup did not put it there, so
; Inno would otherwise leave the folder behind
Type: filesandordirs; Name: "{app}\_internal\work"
Type: dirifempty;     Name: "{app}"

[Code]
{ An upgrade must start from an EMPTY _internal. Inno only overwrites the
  files it ships; it never removes what an older build left behind, and a
  PyInstaller tree from another Python (the 2026-09-21 build was 3.13, the
  2026-10-12 one 3.11) leaves stale .pyd / package folders that the new app
  then imports - uvicorn died with "httptools has no attribute
  HttpRequestParser" and the browser sat on "loading" for ever. The job
  scratch under _internal\work is the only thing Setup did not put there, so
  it is parked beside _internal for the duration and moved back. }
procedure CurStepChanged(CurStep: TSetupStep);
var
  App, Internal, Keep: String;
begin
  App := ExpandConstant('{app}');
  Internal := App + '\_internal';
  Keep := App + '\_work_keep';
  if CurStep = ssInstall then
  begin
    if DirExists(Internal) then
    begin
      if DirExists(Internal + '\work') then
        RenameFile(Internal + '\work', Keep);
      DelTree(Internal, True, True, True);
    end;
  end
  else if CurStep = ssPostInstall then
  begin
    if DirExists(Keep) then
    begin
      if DirExists(Internal + '\work') then
        DelTree(Internal + '\work', True, True, True);
      RenameFile(Keep, Internal + '\work');
    end;
  end;
end;

function WordInstalled(): Boolean;
begin
  Result := RegKeyExists(HKEY_CLASSES_ROOT, 'Word.Application');
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  { DocPilot renders every PDF through Word automation. Warn, but let them
    install anyway - Word is often deployed after the tools are. }
  if not WordInstalled() then
    Result := MsgBox('Microsoft Word was not detected on this PC.'#13#10#13#10 +
                     'DocPilot uses Word to produce the PDF versions of each ' +
                     'document. Without it, DOCX output still works but PDF ' +
                     'export will fail.'#13#10#13#10 +
                     'Continue with the installation?',
                     mbConfirmation, MB_YESNO) = IDYES;
end;
