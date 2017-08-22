; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "tbfFotos"
#define MyAppVerName "tbrFotos 1.0"
#define MyAppPublisher "tbrSoft"
#define MyAppURL "www.tbrsoft.com"
#define MyAppExeName "WebCam.exe"

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
OutputBaseFilename=Instalar_tbrFotos
;SetupIconFile=H:\otros\multimedia\iconos\MISC25.ICO
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

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
;NO TOCAR
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
; FIN NO TOCAR
;fileSystem Object
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrnes.dll";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrun.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
;WBMem info de la PC
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\wbemdisp.tlb"; DestDir: "{sys}\Wbem"; Flags: restartreplace uninsneveruninstall sharedfile
;Entrada y salida de puertos
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\inpout32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall

;MSCOMCT2.DLL
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\MSCOMCT2.OCX";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver

;----------------------------------------------------------------------------
;NO REGISTRO
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\MSWINSCK.OCX";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\ijl11.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\inpout32.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile

;SI REGISTRO
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\X_Boton II.ocx";   DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;----------------------------------------------------------------------------

;archivos que yo necesito, llevan FLAGs distintos
;Source: "D:\dev\mpRock2\prec.dll"; DestDir: "{sys}"; Flags: ignoreversion

;dlls de tbrsoft
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrReg.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

;ejecutables
Source: "D:\BUP MANUEL\Trabajo tbrSoft\Sacar Fotos\WebCam\WebCam.exe"; DestDir: "{app}"; Flags: ignoreversion

;iconos y accesos directos creados
[Icons]
Name: "{group}\{#MyAppName}\tbrFotos"; Filename: "{app}\{#MyAppExeName}"
;Name: "{group}\{#MyAppName}\Manual"; Filename: "{app}\guia del usuario.doc"
Name: "{group}\{#MyAppName}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon
;cosas que se pueden ejectura al finalizar el instalador. Lo mas comun es el programa y el manual
[Run]
;Filename: "{app}\testsks.exe"; Description: "Probar la interfase SKS"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent