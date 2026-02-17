; AESCConstruct2026_Installer.iss

[Setup]
AppName=AESC Construct 2026
AppVersion=2026.1.0
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
SetupIconFile=icons\AESC_Logo.ico
WizardSmallImageFile=icons\AESCLogo.bmp

#define DestPlugin "C:\Program Files\ANSYS Inc\v251\scdm\Addins\AESCConstruct2026"
#define DestData   "C:\ProgramData\AESCConstruct"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
; Ensure the destination directories exist
Name: "{#DestPlugin}"; Flags: uninsneveruninstall
; Give full control to local Users on Data folder
Name: "{#DestData}";   Permissions: users-full; Flags: uninsneveruninstall

[Files]
Source: "C:\Program Files\ANSYS Inc\v251\scdm\Addins\AESCConstruct2026\*"; DestDir: "{#DestPlugin}"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "C:\ProgramData\AESCConstruct\*";   DestDir: "{#DestData}";   Flags: recursesubdirs createallsubdirs ignoreversion

[Run]
Filename: "{cmd}"; Parameters: "/c echo Files have been placed successfully."; Flags: runhidden
