; AESCConstruct2026_Installer.iss

[Setup]
AppName=AESC Construct 2026
AppVersion=2026.1.1
DefaultDirName={autopf}\AESCConstruct2026_Installer
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no
OutputBaseFilename=AESCConstruct2026_Installer
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
WizardSmallImageFile=icons\AESCLogo.bmp

; Source path for plugin files (build output — matches the csproj OutputPath)
#define SourcePlugin "C:\Program Files\ANSYS Inc\v242\scdm\Addins\AESCConstruct2026"
#define DestData     "C:\ProgramData\AESCConstruct"
; Canonical, version-controlled language file (repo is the source of truth, not ProgramData)
#define RepoLang     SourcePath + "\..\Language\languageConstruct.csv"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
; Give full control to local Users on Data folder
Name: "{#DestData}"; Permissions: users-full; Flags: uninsneveruninstall

[Files]
; Stage plugin files in installer app directory; [Code] deploys to all ANSYS versions
Source: "{#SourcePlugin}\*"; DestDir: "{app}\Plugin"; Flags: recursesubdirs createallsubdirs ignoreversion
; Data files go directly to their destination (dev-only files and the stale
; ProgramData language copy excluded — the language file ships from the repo)
Source: "{#DestData}\*"; DestDir: "{#DestData}"; Excludes: "*.bak,*_Input.xlsm,languageConstruct.csv"; Flags: recursesubdirs createallsubdirs ignoreversion
; Ship the canonical, version-controlled language file from the repo
Source: "{#RepoLang}"; DestDir: "{#DestData}\Language"; DestName: "languageConstruct.csv"; Flags: ignoreversion

[Code]
procedure DeployToAllVersions;
var
  AnsysRoot, PluginSource, ScdmAddins, AddinsDir, DeployedVersions: String;
  FindRec: TFindRec;
  ResultCode, Count: Integer;
begin
  AnsysRoot := ExpandConstant('{pf}\ANSYS Inc');
  PluginSource := ExpandConstant('{app}\Plugin');
  Count := 0;
  DeployedVersions := '';

  if not DirExists(AnsysRoot) then
  begin
    MsgBox('No ANSYS installation found at:' + #13#10 + AnsysRoot, mbError, MB_OK);
    Exit;
  end;

  if FindFirst(AnsysRoot + '\v*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          ScdmAddins := AnsysRoot + '\' + FindRec.Name + '\scdm\Addins';
          if DirExists(ScdmAddins) then
          begin
            AddinsDir := ScdmAddins + '\AESCConstruct2026';
            ForceDirectories(AddinsDir);
            Exec('cmd.exe',
                 '/c xcopy "' + PluginSource + '\*" "' + AddinsDir + '\" /E /Y /Q',
                 '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
            Count := Count + 1;
            if DeployedVersions <> '' then
              DeployedVersions := DeployedVersions + ', ';
            DeployedVersions := DeployedVersions + FindRec.Name;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;

  if Count = 0 then
    MsgBox('No SpaceClaim installations found under:' + #13#10 + AnsysRoot, mbError, MB_OK)
  else
    MsgBox('Plugin deployed to ' + IntToStr(Count) + ' version(s): ' + DeployedVersions,
           mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    DeployToAllVersions;
end;
