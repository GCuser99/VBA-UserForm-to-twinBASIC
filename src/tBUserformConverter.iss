// This installer script references Bill Stewart's UninsIS DLL.
// see https://github.com/Bill-Stewart/UninsIS
// With this installer, the user can (re)install the DLL in a location
// of their choice (DisableDirPage=no). If the user had already
// installed an older (or even same or newer) version compared to the one
// user is currently installing, and then decides to change the
// install location, this installer (with help from UninsIS) will first
// uninstall the previously installed version from the old location,
// so that there is only one version on the system at any one point in time. 
// 
#define AppName "twinBASIC Userform Converter"
#define AppGUID "{ECF26FA9-EE79-104A-ADD2-5A0D56287D2A}"
#define AppPublisher "GCUser99"
#define AppURL "https://github.com/GCuser99/VBA-UserForm-to-twinBASIC"
#define AppHelpURL "https://github.com/GCuser99/VBA-UserForm-to-twinBASIC#quick-how-to-use"
#define InstallerName "tBUserformConverterSetup"
#define DLL64FilePath "..\Build\tBUserFormConverter_win64.dll"
#define DLL32FilePath "..\Build\tBUserFormConverter_win32.dll"
#define LicenseFilePath "..\..\..\GitHub\VBA-UserForm-to-twinBASIC\LICENSE.txt"
#define TestFolderPath "..\test_documents"
#define LogoFilePath "..\logo\logo_setup.bmp"
#define RequirementsFilePath ".\readme.txt"
#define SetupOutputFolderPath "..\..\..\GitHub\VBA-UserForm-to-twinBASIC\dist" 
#define AppVersion GetVersionNumbersString(DLL64FilePath)
// The following definition points to the path of 
// Bill Stewart's UninsIS.dll
#define UninstallDLLFilePath ".\UninsIS\UninsIS-1.0.1\UninsIS.dll"

[Setup]

AppId={{#AppGUID}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
; Set default install location
DefaultDirName={localappdata}\{#AppName}
DefaultGroupName={#AppName}
; Remove the following line to run in administrative
; install mode (install for all users.)
PrivilegesRequired=lowest
OutputBaseFilename={#InstallerName}
LicenseFile={#LicenseFilePath}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
; Uninstallable determines if Inno Setup's 
; automatic uninstaller is to be included in
; the installation folder - this must be set to
; "yes" for PrepareToInstall code to function
; correctly
Uninstallable=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
WizardImageFile={#LogoFilePath}
DisableWelcomePage=no
DisableProgramGroupPage=yes
InfoBeforeFile={#RequirementsFilePath}
; DisableDirPage must be set to "no" to allow 
; User to select a different install location
; if updating
DisableDirPage=no
OutputDir={#SetupOutputFolderPath}

[Languages]

Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]

Name: "pkg_core"; Description: "Userform Converter ActiveX Dll"; Types: full compact custom; Flags: fixed;
Name: "pkg_docs";  Description: "UserForm Test Files"; Types: full compact custom;
  
[Messages]

FinishedLabel=Setup has finished installing [name] on your computer. A shortcut to the Install folder can be found on your Desktop.

[Files]
Source: {#DLL64FilePath}; DestDir: {app};  Flags: ignoreversion regserver ; Check: InstallX64; Components: pkg_core;
Source: {#DLL32FilePath}; DestDir: {app};  Flags: ignoreversion regserver ; Check: InstallX32; Components: pkg_core;
Source: {#TestFolderPath}\dlgAllMSControls.frm; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs; 
Source: {#TestFolderPath}\dlgAllMSControls.frx; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#TestFolderPath}\readme.txt; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#LicenseFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
Source: {#RequirementsFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
; Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme
; For importing DLL functions at setup
Source: {#UninstallDLLFilePath}; Flags: dontcopy

[Icons]
Name: "{autodesktop}\tBUserFormConverter - Shortcut"; Filename: "{app}"
Name: "{app}\GitHub help documentation"; Filename: "{#AppHelpURL}"

[Code]
const
  AppPathsKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths';
  OFFICE_UNKNOWN_BIT = -1;
  OFFICE_32_BIT = 0;
  OFFICE_64_BIT = 6;

// Office version detection notes
// Only if 2007 then we can rule out 64 bit. 2010, 2013, 2016, 
// and 365 all have both 32 and 64 bit versions
// MS not supporting 2013 after April 2023

// first get the path to executables:
// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe
// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE
// The default key value of above should yield something like:
// C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE
// C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE
// then use GetBinaryType to discover bitness

Function GetBinaryType(ApplicationName: string; var BinaryType: Integer): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

// The function below is based on:
// https://stackoverflow.com/questions/47431008/getting-the-version-and-platform-of-office-application-from-windows-registry#47443674
Function GetOfficeBitness():Integer;
  Var OfficeApps: array[1..5] of string; 
  Var i: Integer;
  Var OfficeAppPath: String;
  Var BinaryType: Integer;
  Var KeyFound: Boolean;
  Begin
    //find the Office app binary path from Registry - 
    //first try Excel, and if that fails then Access, then Word, etc
    //once path is found, then use API to determine bitness of app
    Result := OFFICE_UNKNOWN_BIT;

    OfficeApps[1]:= 'excel.exe';
    OfficeApps[2]:= 'MSACESS.EXE';
    OfficeApps[3]:= 'Winword.exe';
    OfficeApps[4]:= 'powerpnt.exe';
    OfficeApps[5]:= 'OUTLOOK.exe';

    For i:=1 To High(OfficeApps) Do
      Begin
        KeyFound := RegQueryStringValue(HKLM, AppPathsKey + '\' + OfficeApps[i], '', OfficeAppPath);
        If KeyFound Then Break;
      End;
    
    If KeyFound Then
      Begin
        // find the bitness of the application binary 
        if GetBinaryType(OfficeAppPath, BinaryType) Then Result:=BinaryType;       
      End;
  End;

Function InstallX64(): Boolean;
  Begin
    Result := (GetOfficeBitness = OFFICE_64_BIT);
  End;

Function InstallX32: Boolean;
  Begin
    Result := IsWin64 And (GetOfficeBitness = OFFICE_32_BIT);
  End;

// The following code is for implementation of Bill Stewart's
// UninsIS DLL - see https://github.com/Bill-Stewart/UninsIS
// Import IsISPackageInstalled() function from UninsIS.dll at setup time
Function DLLIsISPackageInstalled(AppId: string; Is64BitInstallMode,
  IsAdminInstallMode: DWORD): DWORD;
  external 'IsISPackageInstalled@files:UninsIS.dll stdcall setuponly';

// Import CompareISPackageVersion() function from UninsIS.dll at setup time
Function DLLCompareISPackageVersion(AppId, InstallingVersion: string;
  Is64BitInstallMode, IsAdminInstallMode: DWORD): LongInt;
  external 'CompareISPackageVersion@files:UninsIS.dll stdcall setuponly';

// Import UninstallISPackage() function from UninsIS.dll at setup time
Function DLLUninstallISPackage(AppId: string; Is64BitInstallMode,
  IsAdminInstallMode: DWORD): DWORD;
  external 'UninstallISPackage@files:UninsIS.dll stdcall setuponly';

// Wrapper for UninsIS.dll IsISPackageInstalled() function
// Returns true if package is detected as installed, or false otherwise
Function IsISPackageInstalled(): Boolean;
  Begin
    result := DLLIsISPackageInstalled('{#AppGUID}',  // AppId
      DWORD(Is64BitInstallMode()),                   // Is64BitInstallMode
      DWORD(IsAdminInstallMode())) = 1;              // IsAdminInstallMode
    If result Then
      Log('UninsIS.dll - Package detected as installed')
    Else
      Log('UninsIS.dll - Package not detected as installed');
  End;

// Wrapper for UninsIS.dll CompareISPackageVersion() function
// Returns:
// < 0 if version we are installing is < installed version
// 0   if version we are installing is = installed version
// > 0 if version we are installing is > installed version
Function CompareISPackageVersion(): LongInt;
  Begin
    result := DLLCompareISPackageVersion('{#AppGUID}',  // AppId
      '{#AppVersion}',                                  // InstallingVersion
      DWORD(Is64BitInstallMode()),                      // Is64BitInstallMode
      DWORD(IsAdminInstallMode()));                     // IsAdminInstallMode
    If result < 0 Then
      Log('UninsIS.dll - This version {#AppVersion} older than installed version')
    Else If result = 0 Then
      Log('UninsIS.dll - This version {#AppVersion} same as installed version')
    Else
      Log('UninsIS.dll - This version {#AppVersion} newer than installed version');
  End;

// Wrapper for UninsIS.dll UninstallISPackage() function
// Returns 0 for success, non-zero for failure
Function UninstallISPackage(): DWORD;
  Begin
    result := DLLUninstallISPackage('{#AppGUID}',  // AppId
      DWORD(Is64BitInstallMode()),                 // Is64BitInstallMode
      DWORD(IsAdminInstallMode()));                // IsAdminInstallMode
    If result = 0 Then
      Log('UninsIS.dll - Installed package uninstall completed successfully')
    Else
      Log('UninsIS.dll - installed package uninstall did not complete successfully');
  End;

Function PrepareToInstall(Var NeedsRestart: Boolean): string;
  Var oldInstallDir: String; newInstallDir: String;
  Var isDifferentLocation: Boolean;
  Var isDifferentVersion: Boolean;
  Var isInstalled: Boolean;
  Var alwaysUninstall: Boolean;
  Begin
    // set alwaysUninstall to true if uninstall
    // should run regardless of version and location compare
    alwaysUninstall := false;
    isDifferentLocation := false;
    isDifferentVersion := false;

    oldInstallDir := WizardForm.PrevAppDir;
    newInstallDir := ExpandConstant('{app}');
    
    If oldInstallDir <> newInstallDir Then isDifferentLocation := true;

    If CompareISPackageVersion() <> 0 Then isDifferentVersion := true;

    isInstalled := IsISPackageInstalled()

    If isInstalled Then
      Begin
        If isDifferentLocation Or isDifferentVersion Or alwaysUninstall Then
          UninstallISPackage();
      End;

    result := '';
  End;

// Check if the OS/Office requirements have been met - if not warn user 
Function InitializeSetup(): Boolean;
  Var OfficeBitness: Integer;
  Var answer: Integer;
  Begin
    if Not IsWin64 then
      Begin
        answer := MsgBox('Setup has determined that your OS is not 64-bit Windows, which is a requirement of this installation. Do you still want to proceed?', mbConfirmation, MB_YESNO); 
        If answer = IDYES Then
          Begin
            Result:=True;
          End 
        Else
          Begin
            Result := False;
            Exit;
          End; 
      End;
    OfficeBitness:= GetOfficeBitness
    Case OfficeBitness of  
      OFFICE_UNKNOWN_BIT : Begin 
        answer := MsgBox('MS Office bitness could not be determined. Are you sure that you want to proceed with the installation?', mbConfirmation, MB_YESNO); 
        if answer = IDYES then Result:=True else Result := False End;
      OFFICE_32_BIT : Begin 
        Result:=True; End; 
      OFFICE_64_BIT : Begin 
        Result:=True; End;
    Else
        Begin
          answer := MsgBox('The installed version of MS Office was found but is not compatible with this installation. Are you sure that you want to proceed?', mbConfirmation, MB_YESNO); 
          if answer = IDYES then Result:=True else Result := False;
        End;      
    End;
  End;
