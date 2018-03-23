
;Copyright (C) 2001-2002 Alexander Minza
;alex_minza@hotmail.com
;http://www.ournet.md/~fngraph
;http://www.hi-tech.ournet.md

[Setup]
AdminPrivilegesRequired=yes
AppName=FNGraph
AppVerName=FNGraph 2.61
AppVersion=2.61
AppMutex=FNGraphAppMutex
AppPublisher=Alexander Minza
AppPublisherURL=http://www.ournet.md/~fngraph
AppSupportURL=http://www.ournet.md/~fngraph
AppUpdatesURL=http://www.ournet.md/~fngraph
AppCopyright=Copyright (C) 2001-2002 Alexander Minza
ChangesAssociations=yes
Compression=zip/9
DefaultDirName={pf}\FNGraph
DefaultGroupName=FNGraph
LicenseFile=License.txt
MinVersion=4,4
OutputBaseFilename=fngraph2_61
UninstallDisplayIcon={app}\FNGraph.exe,0

[Registry]
Root: HKCU; Subkey: "Software\VB and VBA Program Settings\FNGraph"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\VB and VBA Program Settings\FNGraph\Options"; ValueType: string; ValueName: "RecentFolder"; ValueData: "{app}\Samples"; Flags: createvalueifdoesntexist

Root: HKCR; Subkey: ".fng"; ValueType: string; ValueName: ""; ValueData: "FNGraph.Document"; Flags: uninsdeletekey
Root: HKCR; Subkey: ".fng\ShellNew"; ValueType: string; ValueName: "NullFile"

Root: HKCR; Subkey: "FNGraph.Document"; ValueType: string; ValueName: ""; ValueData: "FNGraph Document"; Flags: uninsdeletekey
Root: HKCR; Subkey: "FNGraph.Document\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\FNGraph.exe,1"
Root: HKCR; Subkey: "FNGraph.Document\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\FNGraph.exe"" ""%1"""

[Dirs]
Name: "{app}\Samples"

[Files]
Source: "..\FNGraph.exe"; DestDir: "{app}"
Source: "..\FNGraph.exe.manifest"; DestDir: "{app}"
Source: "..\HHelp\FNGraph.chm"; DestDir: "{app}"
Source: "License.txt"; DestDir: "{app}"
;Source: "..\PAD\Alexander Minza\fngraph_pad.xml"; DestDir: "{app}"
Source: "..\Samples\*.*"; DestDir: "{app}\Samples"

Source: "COMDLG32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver

[Icons]
Name: "{group}\FNGraph"; Filename: "{app}\FNGraph.exe"; WorkingDir: "{app}";
Name: "{group}\FNGraph Help"; Filename: "{app}\FNGraph.chm"
Name: "{group}\Samples"; Filename: "{app}\Samples"
Name: "{userdesktop}\FNGraph"; Filename: "{app}\FNGraph.exe"; WorkingDir: "{app}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\FNGraph"; Filename: "{app}\FNGraph.exe"; WorkingDir: "{app}"; Tasks: quicklaunchicon

[Tasks]
Name: desktopicon; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:";
Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\FNGraph.exe"; Description: "Launch FNGraph"; Flags: nowait postinstall skipifsilent

