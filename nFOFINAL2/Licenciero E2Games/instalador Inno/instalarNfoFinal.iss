

;SEGUIRAQUI hacer multiinstalador !!!

#define MyAppName "e2G-Licencia"
#define MyAppVerName "e2G-Licencia v1"
#define MyAppPublisher "tbrSoft Internacional"
#define MyAppURL "http://www.tbrsoft.com"
#define MyAppExeName "e2G-Licencia.exe"

[Setup]
AppName={#MyAppName}
AppVerName={#MyAppVerName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
;DisableDirPage=yes
DefaultGroupName=tbrSoft
AllowNoIcons=no
OutputBaseFilename=Instalar-e2G-Licencias
;SetupIconFile=D:\dev\3PM kundera 70000\3pm.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "eng"; MessagesFile: "compiler:Default.isl"
Name: "bra"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "cat"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "cze"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "dan"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dut"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "fin"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "fre"; MessagesFile: "compiler:Languages\French.isl"
Name: "ger"; MessagesFile: "compiler:Languages\German.isl"
Name: "hun"; MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "ita"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "nor"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "pol"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "por"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "rus"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slo"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spa"; MessagesFile: "compiler:Languages\Spanish.isl"
;Name: "spa"; MessagesFile: "compiler:Languages\SpanishMex.isl"
;Name: "spa"; MessagesFile: "compiler:Languages\SpanishEsp.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver

Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\wbemdisp.tlb"; DestDir: "{sys}\Wbem"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrtimer.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrjuse.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrerr.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
; este lo dejo para que confunda
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrNfo.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
; este es el verdadero sist de licencias
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrNewSys.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrCaescrypto.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrnes.dll";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrun.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver

Source: "..\e2G-Licencia.exe"; DestDir: "{app}"; Flags: ignoreversion
;archivo de datos para dar licencias!!!
Source: "..\e2Games.load"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

