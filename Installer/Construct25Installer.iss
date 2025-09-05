; AESCConstruct25_Installer.iss

[Setup]
AppName=AESC Construct 25
AppVersion=1.0.0
DefaultDirName={autopf}\AESCConstruct25_Installer
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no
OutputBaseFilename=AESCConstruct25_Installer
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
SetupIconFile=icons\AESC_Logo.ico
WizardSmallImageFile=icons\AESCLogo.bmp

; Current temporary destinations:
#define DestPlugin "C:\ProgramData\SpaceClaim\AddIns\AESCConstruct25"
#define DestData   "C:\ProgramData\AESCConstruct"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
; Ensure the destination directories exist
Name: "{#DestPlugin}"; Flags: uninsneveruninstall
; Give full control to local Users on Data folder
Name: "{#DestData}";   Permissions: users-full; Flags: uninsneveruninstall

[Files]
Source: "C:\Program Files\ANSYS Inc\v251\scdm\Addins\AESCConstruct25\*"; DestDir: "{#DestPlugin}"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "C:\ProgramData\AESCConstruct\*";   DestDir: "{#DestData}";   Flags: recursesubdirs createallsubdirs ignoreversion

[Run]
Filename: "{cmd}"; Parameters: "/c echo Files have been placed successfully."; Flags: runhidden

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
//  if CurStep = ssDone then
//  begin
//    MsgBox('AESC data and plugin have been installed.'#13#10 +
//           'Plugin: ' + ExpandConstant('{#DestPlugin}') + #13#10 +
//           'Data:   ' + ExpandConstant('{#DestData}'),
//           mbInformation, MB_OK);
//  end;
end;
