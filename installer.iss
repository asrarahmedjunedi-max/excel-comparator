; ============================================================
;  Inno Setup script for Excel File Comparator
;  Created by Asrar Ahmed Junedi
;
;  1. Download & install Inno Setup (free): https://jrsoftware.org/isinfo.php
;  2. Run build_exe.bat first so dist\ExcelComparator.exe exists.
;  3. Open this file in Inno Setup and click Compile (or right-click
;     it -> "Compile" if you associated .iss files with Inno Setup).
;  4. This produces a single Setup installer .exe. Running that
;     installer puts a shortcut on the user's Desktop and Start Menu.
; ============================================================

#define MyAppName "Excel File Comparator"
#define MyAppVersion "1.0"
#define MyAppPublisher "Asrar Ahmed Junedi"
#define MyAppExeName "ExcelComparator.exe"

[Setup]
AppId={{7C6F2D2E-6C7E-4B8E-9E1B-EXCEL-COMPARE}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=ExcelComparatorSetup
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=yes
SetupIconFile=app_icon.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; This is what makes the Desktop icon a real (unchecked-by-default) option
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: checkedonce

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
