VERSION 5.00
Begin VB.Form frmSystemInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Information"
   ClientHeight    =   1110
   ClientLeft      =   975
   ClientTop       =   990
   ClientWidth     =   3690
   Icon            =   "SI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1110
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Inventory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ArgArray() As String 'tab the arg command line

' Some variables to obtain the system information

Dim Systeem As New System   ' Systeem is Dutch for System
Dim Drive As String

Dim TimeZone As String

Dim Percent As Byte
Dim Free As Long
Dim Total As Long
Dim PagFil As Long
Dim AvPagFil As Long
Dim VirFil As Long
Dim AvVirFil As Long

Dim Processor As String
Dim Number As Long
Dim Active As Long
Dim OemId As Long
Dim PagSiz As Long
Dim MinAdr As Long
Dim MaxAdr As Long
Dim Alloc As Long
Dim Resv As Long

Dim Maj As Integer
Dim Min As Integer
Dim BuildNb As Long
Dim CSDVer As String * 128
Dim pathh As String
Dim Version As String
Dim TotalDiskSpace As Long
Dim FreeDiskSpace As Long
Dim ENTER As String

Dim WinProductId As String
Dim WinOrganisation As String
Dim WinOwner As String
Dim IExplore As String

Dim Identifier As String
Dim SysBiosDate As String
Dim BiosVersion As String
Dim Proce(3) As String

Dim VideoBiosVersion As String
Dim VideoBiosDate As String
Dim VideoX As String
Dim VideoY As String
Dim VideoHz As String
Dim VideoDevice As String
Dim VideoCard As String
Dim VideoChip As String
Dim VideoDac As String

Dim Removable As Integer
Dim Fixed As Integer
Dim Ram As Integer
Dim NetWork As Integer
Dim CDrom As Integer

Dim lngCount As Long 'duree depuis windows is open
Dim lngHours As Long
Dim lngMinutes As Long

Dim BiosDate As String

Dim SrvDNS As String
Dim SrvDHCP As String
Dim AdrGate As String
Dim AdrMask As String


Dim LogDir As String 'path for the log file, defined in thr form load


Private Declare Function apiGetDrives Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'---------------------------------------------------------------------------
' API declarations to obtain the time during Windows is power on
'
'---------------------------------------------------------------------------
Private Declare Function GetTickCount Lib "kernel32" () As Long

Sub DisplayInfo()
Dim FindRegKeyPath As String
Dim RegKeyPath As String

    'find the display info in the registry;
    
    'find the registry key where locate the display info
    FindRegKeyPath = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\VIDEO", "\Device\Video0")
    RegKeyPath = "HKEY_LOCAL_MACHINE\" & Mid(FindRegKeyPath, 19, Len(FindRegKeyPath) - 18)
    
    VideoDevice = GetStringValue(RegKeyPath, "Device Description")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\" & Mid(FindRegKeyPath, 19, Len(FindRegKeyPath) - 18)
    VideoCard = GetBinaryValue(RegKeyPath, "HardwareInformation.AdapterString")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\" & Mid(FindRegKeyPath, 19, Len(FindRegKeyPath) - 18)
    VideoChip = GetBinaryValue(RegKeyPath, "HardwareInformation.ChipType")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\" & Mid(FindRegKeyPath, 19, Len(FindRegKeyPath) - 18)
    VideoDac = GetBinaryValue(RegKeyPath, "HardwareInformation.DacType")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Hardware Profiles\Current\System\CurrentControlSet\" & Mid(FindRegKeyPath, 40, Len(FindRegKeyPath) - 39)
    VideoHz = GetDWORDValue(RegKeyPath, "DefaultSettings.VRefresh")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Hardware Profiles\Current\System\CurrentControlSet\" & Mid(FindRegKeyPath, 40, Len(FindRegKeyPath) - 39)
    VideoX = GetDWORDValue(RegKeyPath, "DefaultSettings.XResolution")
    
    RegKeyPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Hardware Profiles\Current\System\CurrentControlSet\" & Mid(FindRegKeyPath, 40, Len(FindRegKeyPath) - 39)
    VideoY = GetDWORDValue(RegKeyPath, "DefaultSettings.YResolution")
    
End Sub


Sub Drives()

Dim Retrn As Long
Dim buffer As Long
Dim temp As String
Dim intI As Integer
Dim Read(1 To 100) As String
Dim Counter As Integer
Dim X As String

buffer = 10

' This first part is copied from the function Drives in the Class system
' so for the explanation, see the Drives function.

Again:
temp = Space$(buffer)
Retrn = apiGetDrives(buffer, temp)

If Retrn > buffer Then
    buffer = Retrn
    GoTo Again
End If

' The API returns something like :
' A:\*B:\*C:\*D:\**  , with  * = NULL character
' 1234123412341234
' \ 1 \ 2 \ 3 \ 4 \
'
' So we start reading three characters, we step 4 further (the three we read + the
' NULL-character), and we read again three characters, step 4, ect.

Counter = 0
For intI = 1 To (buffer - 4) Step 4
    Counter = Counter + 1
    Read(Counter) = Mid$(temp, intI, 3)   ' Read all the drives into this array
Next                                      ' Now this array contains all the drives
                                          ' so we can check them all out.



For intI = 1 To Counter
' Change the Drive to all the drives we have stored in the array.
' Now we can show information about all the drives in your computer.
    
    Drive = Read(intI)
    Systeem.DriveInfo Drive, TotalDiskSpace, FreeDiskSpace

    If TotalDiskSpace = 0 Then
    'No disk was loaded! Show nothing
        'X = INIWrite("Disks", Drive, "-------------", LogDir)
        X = INIWrite("Disks", "Disk " + Drive, Systeem.DriveType(Drive) _
        + " - (No Disk loaded)", LogDir)
        'X = INIWrite("Disks", "Total disk space " + Drive, "No disk loaded", LogDir)
        'X = INIWrite("Disks", "Free disk space " + Drive, "-", LogDir)
        'X = INIWrite("Disks", "Volumelabel " + Drive, "-", LogDir)
        'X = INIWrite("Disks", "Serial Number " + Drive, "-", LogDir)
        
    Else
    ' If a disk was loaded, show all the information.
        'X = INIWrite("Disks", Drive, "-------------", LogDir)
        X = INIWrite("Disks", "Disk " + Drive, Systeem.DriveType(Drive) & " - " & Systeem.ShowFileSystemType(Drive) _
        + " - Capacity " + Systeem.RetStringKb(TotalDiskSpace) + " - Not use " _
         + Systeem.RetStringKb(FreeDiskSpace) + " - Label " _
         + Systeem.VolumeLabel(Drive) + " - Serial nb " _
         + Systeem.SerialNumber(Drive), LogDir)
        'X = INIWrite("Disks", "Total disk space " + Drive, Systeem.PutPoints(TotalDiskSpace), LogDir)
        'X = INIWrite("Disks", "Free disk space " + Drive, Systeem.PutPoints(FreeDiskSpace), LogDir)
        'X = INIWrite("Disks", "Volumelabel " + Drive, Systeem.VolumeLabel(Drive), LogDir)
        'X = INIWrite("Disks", "Serial Number " + Drive, Systeem.SerialNumber(Drive), LogDir)
    
    
    End If

Next

End Sub
Private Sub WIS_proc()
Dim X As String

ENTER = Chr$(13) + Chr$(10)


' Call some of the function that return values in variables
Systeem.FreeMemory Percent, Total, Free, PagFil, AvPagFil, VirFil, AvVirFil
Systeem.SystemInfo Processor, Number, Active, OemId, PagSiz, MinAdr, MaxAdr, Alloc, Resv
Systeem.Drives Removable, Fixed, CDrom, Ram, NetWork
Systeem.WinVer Maj, Min, Version, BuildNb, CSDVer
Systeem.InfoNetWorkAdr SrvDNS, SrvDHCP, AdrMask, AdrGate

WinProductId = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId")
WinOrganisation = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
WinOwner = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
TimeZone = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation", "DaylightName")

IExplore = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "Plus! VersionNumber")

Identifier = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System", "Identifier")
SysBiosDate = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System", "SystemBiosDate")
BiosVersion = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System", "SystemBiosVersion")
Proce(0) = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "Identifier") + " "
Proce(1) = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "VendorIdentifier")
Proce(2) = GetDWORDValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "~MHz")

VideoBiosVersion = GetDWORDValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System", "VideoBiosVersion")
VideoBiosDate = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System", "VideoBiosDate")
'Systeem.DisplayParametersInfo VideoX, VideoY

'caculc time windows is steel running
lngCount = GetTickCount
lngHours = ((lngCount / 1000) / 60) / 60
lngMinutes = ((lngCount / 1000) / 60) Mod 60


X = INIWrite("OS_Version", "Operating System", Version, LogDir)
X = INIWrite("OS_Version", "Windows version", CStr(Maj) + "." + CStr(Min) _
        + " .number " + CStr(BuildNb) + ": " + RTrim$(CSDVer), LogDir)
X = INIWrite("OS_Version", "Product ID", WinProductId, LogDir)
X = INIWrite("OS_Version", "Licence Owner", WinOwner, LogDir)
X = INIWrite("OS_Version", "Licence Company", WinOrganisation, LogDir)
X = INIWrite("OS_Version", "TimeZone", TimeZone, LogDir)
WhiteLine LogDir

If ArgArray(2) = "Y" Then 'if the command line argument /SR is enter
    X = INIWrite("SystemReport", "Identifier", Identifier, LogDir)
    X = INIWrite("SystemReport", "SystemBiosDate", SysBiosDate, LogDir)
    X = INIWrite("SystemReport", "BiosVersion", BiosVersion, LogDir)
    X = INIWrite("SystemReport", "Processor", Proce(0), LogDir)
    X = INIWrite("SystemReport", "Speed", " ~" & Proce(2) & " Mhz", LogDir)
    X = INIWrite("SystemReport", "Factory", Proce(1), LogDir)
    WhiteLine LogDir
End If

If ArgArray(3) = "Y" Then 'if the command line argument /VID is enter
    X = INIWrite("VIDEO", "VideoBiosVersion", VideoBiosVersion, LogDir)
    X = INIWrite("VIDEO", "VideoBiosDate", VideoBiosDate, LogDir)
    X = INIWrite("VIDEO", "ScreenSize", VideoX & " x " & VideoY & " x " & VideoHz & " Hz", LogDir)
    X = INIWrite("VIDEO", "VideoDevice", VideoDevice, LogDir)
    X = INIWrite("VIDEO", "VideoCard", VideoCard, LogDir)
    X = INIWrite("VIDEO", "VideoChip", VideoChip, LogDir)
    X = INIWrite("VIDEO", "VideoDac", VideoDac, LogDir)
    WhiteLine LogDir
End If


X = INIWrite("SET_Environnement", "OS", Systeem.SetEnv("OS"), LogDir)
X = INIWrite("SET_Environnement", "Windows Directory", Systeem.WinDir, LogDir)
X = INIWrite("SET_Environnement", "System Directory", Systeem.SystemDir, LogDir)
X = INIWrite("SET_Environnement", "Temp Directory", Systeem.TempDir, LogDir)
X = INIWrite("SET_Environnement", "ComSpec", Systeem.SetEnv("ComSpec"), LogDir)
X = INIWrite("SET_Environnement", "Os2LibPath", Systeem.SetEnv("Os2LibPath"), LogDir)
X = INIWrite("SET_Environnement", "Path", Systeem.SetEnv("Path"), LogDir)
X = INIWrite("SET_Environnement", "PATHEXT", Systeem.SetEnv("PATHEXT"), LogDir)
X = INIWrite("SET_Environnement", "USER TEMP", Systeem.SetEnv("TEMP"), LogDir)
X = INIWrite("SET_Environnement", "USER TMP", Systeem.SetEnv("TMP"), LogDir)

X = INIWrite("SET_Environnement", "NUMBER_OF_PROCESSORS", Systeem.SetEnv("NUMBER_OF_PROCESSORS"), LogDir)
X = INIWrite("SET_Environnement", "PROCESSOR_ARCHITECTURE", Systeem.SetEnv("PROCESSOR_ARCHITECTURE"), LogDir)
X = INIWrite("SET_Environnement", "PROCESSOR_IDENTIFIER", Systeem.SetEnv("PROCESSOR_IDENTIFIER"), LogDir)
X = INIWrite("SET_Environnement", "PROCESSOR_LEVEL", Systeem.SetEnv("PROCESSOR_LEVEL"), LogDir)
X = INIWrite("SET_Environnement", "PROCESSOR_REVISION", Systeem.SetEnv("PROCESSOR_REVISION"), LogDir)

X = INIWrite("SET_Environnement", "USERDOMAIN", Systeem.SetEnv("USERDOMAIN"), LogDir)
X = INIWrite("SET_Environnement", "USERNAME", Systeem.SetEnv("USERNAME"), LogDir)
X = INIWrite("SET_Environnement", "USERPROFILE", Systeem.SetEnv("USERPROFILE"), LogDir)
WhiteLine LogDir

X = INIWrite("NetWork", "Hostname", Systeem.computername, LogDir)
X = INIWrite("NetWork", "User name", Systeem.username, LogDir)
X = INIWrite("NetWork", "Running since", lngHours & " hours and " & lngMinutes & _
            " minutes", LogDir)
Systeem.GetWorkstationInfo LogDir
Systeem.NetWorkInf LogDir
X = INIWrite("NetWork", "DNS", SrvDNS, LogDir)
X = INIWrite("NetWork", "DHCP", SrvDHCP, LogDir)
X = INIWrite("NetWork", "Mask", AdrMask, LogDir)
X = INIWrite("NetWork", "GateWay", AdrGate, LogDir)
WhiteLine LogDir

If ArgArray(4) = "Y" Then 'if the command line argument /KB is enter
    X = INIWrite("KeyBoard", "Keyboard Type", Systeem.KeyboardType, LogDir)
    X = INIWrite("KeyBoard", "Functionkeys", Systeem.FunctionKeys, LogDir)
    WhiteLine LogDir
End If

If ArgArray(5) = "Y" Then 'if the command line argument /SOFT is enter
    X = INIWrite("SoftWares", "IExplore", IExplore, LogDir)
    X = INIWrite("SoftWares", "WIS", "1", LogDir)
    WhiteLine LogDir
End If


X = INIWrite("Memory", "Total RAM", CStr(Total) + " Kb", LogDir)
X = INIWrite("Memory", "Free RAM", CStr(Free) + " Kb", LogDir)
X = INIWrite("Memory", "RAM used", CStr(Percent) + " %", LogDir)
X = INIWrite("Memory", "Page File", CStr(PagFil) + " Kb", LogDir)
X = INIWrite("Memory", "Available Page File", CStr(AvPagFil) + " Kb", LogDir)
X = INIWrite("Memory", "Virtual File", CStr(VirFil) + " Kb", LogDir)
X = INIWrite("Memory", "Available Virtual File", CStr(AvVirFil) + " Kb", LogDir)
WhiteLine LogDir

X = INIWrite("Disks", "Nb Removables drives", CStr(Removable), LogDir)
X = INIWrite("Disks", "Nb Locals drives", CStr(Fixed), LogDir)
X = INIWrite("Disks", "Nb CD-ROMs drives", CStr(CDrom), LogDir)
X = INIWrite("Disks", "Nb RAMs drives", CStr(Ram), LogDir)
X = INIWrite("Disks", "Nb Networks drives", CStr(NetWork), LogDir)
' Call the sub Drives (this is a sub of this form, not of the class)
Drives
WhiteLine LogDir

'get the started service
If ArgArray(6) = "Y" Then 'if the command line argument /SS is enter
    Systeem.GetServiceInfo &H1, LogDir
    WhiteLine LogDir
End If

'get the stoped service
Systeem.GetServiceInfo &H2, LogDir
WhiteLine LogDir

'get the host file
If ArgArray(7) = "Y" Then 'if the command line argument /HO is enter
    Systeem.ReadHost LogDir
    WhiteLine LogDir
End If

'get the host file
If ArgArray(8) = "Y" Then 'if the command line argument /LMHOST is enter
    Systeem.ReadLMHost LogDir
    WhiteLine LogDir
End If


'exit the program
End
        
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    'delete the ini file if exist in the current dir
    LogDir = CurDir
    
    If Right(LogDir, 1) = "\" Then
        LogDir = Mid(LogDir, 1, Len(LogDir) - 1)
    End If
    
    LogDir = LogDir + "\" + Systeem.computername + ".INI"
    On Error Resume Next
    Kill LogDir
    
    DoEvents
    
    GetCommandLine 9
    
    If ArgArray(1) = "Y" Then
        MsgBox "Using Windows Information System:" + vbCrLf + vbCrLf _
            + "WIS [/?]  [/SR]  [/Vid]  [/Kb]  [/Soft]  [/SS]  [/Ho]  [/Lm]  [/PATH=path]" + vbCrLf + vbCrLf _
            + "[/?] : This Help" + vbCrLf _
            + "[/SR] : System Report Info" + vbCrLf _
            + "[/Vid] : Video Info" + vbCrLf _
            + "[/Kb] : Keybord Info" + vbCrLf _
            + "[/Soft] : Softwares Info" + vbCrLf _
            + "[/SS] : Services Started Info" + vbCrLf _
            + "[/Ho] : Hosts File" + vbCrLf _
            + "[/LmHost] : LmHost File" + vbCrLf _
            + "[/PATH] : path to write the ini file, GIVE A VALID PATH" + vbCrLf + vbCrLf _
            + "Without PATH argument, writing the ini file in the current directory", 64
        End 'exit the program
    End If
    
    'delete the ini file if exist in the current dir
    If ArgArray(9) <> "N" Then
        LogDir = Mid(ArgArray(9), 7, Len(ArgArray(9))) 'delete /path= in the commande line
        If Right(LogDir, 1) = "\" Then
            LogDir = Mid(LogDir, 1, Len(LogDir) - 1)
        End If
        
        If Dir(LogDir, vbDirectory) <> "" Then
            LogDir = LogDir + "\" + Systeem.computername + ".INI"
        Else
            'if the directory path not exist then the new path is
            LogDir = "c:\temp" + "\" + Systeem.computername + ".INI"
        End If
        On Error Resume Next
        Kill LogDir
    End If
    
    'LogDir = "d:\personal\vb\msinfo1\" + Systeem.computername + ".INI"
    'Kill LogDir
    DisplayInfo
    
    WIS_proc
    
End Sub

Function GetCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   ReDim tmpArg(MaxArgs) As String
   
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      'Test for space or tab.
      If (C <> " " And C <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         'ArgArray(NumArgs) = ArgArray(NumArgs) & C
         tmpArg(NumArgs) = tmpArg(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next I
   
   'Resize array just enough to hold arguments.
   'ReDim Preserve ArgArray(NumArgs)
   ReDim Preserve tmpArg(NumArgs)
   
   'get the arguments in the good case of the tab
    For I = 1 To 9
        ArgArray(I) = "N"
    Next I
   
   For I = 1 To NumArgs
       Select Case UCase(tmpArg(I))
        Case "/?"
            ArgArray(1) = "Y"
        Case "?"
            ArgArray(1) = "Y"
        Case "HELP"
            ArgArray(1) = "Y"
        Case "/HELP"
            ArgArray(1) = "Y"
        
        Case "/SR"
            ArgArray(2) = "Y"
        Case "/VID"
            ArgArray(3) = "Y"
        Case "/KB"
            ArgArray(4) = "Y"
        Case "/SOFT"
            ArgArray(5) = "Y"
        Case "/SS"
            ArgArray(6) = "Y"
        Case "/HO"
            ArgArray(7) = "Y"
        Case "/LMHOST"
            ArgArray(8) = "Y"
    End Select
    
    If UCase(Left(tmpArg(I), 6)) = "/PATH=" Then
        ArgArray(9) = tmpArg(I)
    End If
    
   Next I
   
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function


