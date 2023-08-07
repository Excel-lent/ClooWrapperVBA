[Setup]
AppName=ClooWrapperVBA
AppVerName=ClooWrapperVBA
DefaultDirName={pf}\ClooWrapperVBA
DefaultGroupName=ClooWrapperVBA
Compression=lzma
SolidCompression=yes
SourceDir=.\
PrivilegesRequired=poweruser
AllowCancelDuringInstall=yes
AllowRootDirectory=no
AllowNoIcons=yes
DisableReadyMemo=no
OutputBaseFilename=ClooWrapperVBA setup

[Dirs]
Name: "{app}"; Permissions: users-full 
Name: "{app}\demo"; Permissions: everyone-full
Name: "{app}\demo\cl"; Permissions: everyone-full

[Files]
Source: bin\ClooWrapperVBA.dll; DestDir: {app}; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;  
Source: bin\ClooWrapperVBA_x64.dll; DestDir: {app}; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: bin\Cloo.dll; DestDir: {app}; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: ..\Excel\OpenCl v0.05.xlsm; DestDir: {app}\demo; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: ..\Excel\cl\Performance.cl; DestDir: {app}\demo\cl; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: ..\Excel\cl\MatrixMultiplication.cl; DestDir: {app}\demo\cl; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: ..\Excel\Configuration.vbs; DestDir: {app}\demo; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: bin\register.bat; DestDir: {app}; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;
Source: bin\unregister.bat; DestDir: {app}; Flags: ignoreversion recursesubdirs overwritereadonly; Permissions: everyone-full;

[Icons]
Name: "{group}\Uninstall"; Filename: "{uninstallexe}";

[UninstallDelete]
Type: files; Name: "{app}\ClooWrapperVBA.tlb"
Type: files; Name: "{app}\ClooWrapperVBA_x64.tlb"

[Run]
Filename: "{dotnet40}\RegAsm.exe"; Parameters: /codebase /tlb ClooWrapperVBA.dll; WorkingDir: {app}; Flags: WaitUntilTerminated;
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: /codebase /tlb ClooWrapperVBA_x64.dll; WorkingDir: {app}; Flags: WaitUntilTerminated;

[UninstallRun] 
Filename: "{dotnet40}\RegAsm.exe"; Parameters: /unregister ClooWrapperVBA.dll; WorkingDir: {app};
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: /unregister ClooWrapperVBA_x64.dll; WorkingDir: {app};