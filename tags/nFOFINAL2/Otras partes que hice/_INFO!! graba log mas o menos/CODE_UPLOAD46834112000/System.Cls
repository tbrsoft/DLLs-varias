VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------
' API declarations in order to obtain names (computer and user names)
'---------------------------------------------------------------------------
Private Declare Function apiUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function apiCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'---------------------------------------------------------------------------
' API declarations for Domain/Workgroup and Domain Session
'---------------------------------------------------------------------------
Private Type WKSTA_INFO_101
            wki101_platform_id As Long
            wki101_computername As Long
            wki101_langroup As Long
            wki101_ver_major As Long
            wki101_ver_minor As Long
            wki101_lanroot As Long
         End Type
 
        Private Type WKSTA_USER_INFO_1
            wkui1_username As Long
            wkui1_logon_domain As Long
            wkui1_logon_server As Long
            wkui1_oth_domains As Long
         End Type
 
         Private Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" _
            (lpName As Any, ByVal lpUserName$, lpnLength&)
         Private Declare Function NetWkstaGetInfo& Lib "Netapi32" _
            (strServer As Any, ByVal lLevel&, pbBuffer As Any)
         Private Declare Function NetWkstaUserGetInfo& Lib "Netapi32" _
            (Reserved As Any, ByVal lLevel&, pbBuffer As Any)
         Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)
         Private Declare Sub lstrcpyA Lib "kernel32" (dest As Any, ByVal src As Any)
         'Private Declare Sub RtlMoveMemoryA Lib "kernel32" _
            (dest As Any, src As Any, ByVal size&)
         Private Declare Function NetApiBufferFree& Lib "Netapi32" (ByVal buffer&)
        


'---------------------------------------------------------------------------
' API declarations in order to adr IP and Hostname
'---------------------------------------------------------------------------
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128
 
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
 
Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
 
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
 
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


'---------------------------------------------------------------------------
' API declarations in order to obtain the Windows / System and Temporary
' directories.
'---------------------------------------------------------------------------
Private Declare Function apiWindDir Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiSysDir Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiTempDir Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiSetEnv Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

'---------------------------------------------------------------------------
' API declarations in order to obtain Memory and System information.
'---------------------------------------------------------------------------
Private Declare Sub apiMemStatus Lib "kernel32" Alias "GlobalMemoryStatus" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS                   ' size of 'Type' = 8 x 4 bytes = 32 (a Long is 4 Bytes)
        dwLength As Long                    ' This need to be set at the size of this 'Type'  = 32
        dwMemoryLoad As Long                ' Gives global indication of used RAM (in %)
        dwTotalPhys As Long                 ' Gives total RAM of the computer
        dwAvailPhys As Long                 ' Gives the amount of free RAM
        dwTotalPageFile As Long             ' I don't use this (don't know what it means)
        dwAvailPageFile As Long             ' I don't use this (don't know what it means)
        dwTotalVirtual As Long              ' I don't use this (don't know what it means)
        dwAvailVirtual As Long              ' I don't use this (don't know what it means)
End Type

Private Declare Sub apiSystemInfo Lib "kernel32" Alias "GetSystemInfo" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO                    ' size of 'Type' = 9 x 4 bytes = 36
        dwOemID As Long
        dwPageSize As Long                  ' Must be set at the size of this 'Type'
        lpMinimumApplicationAddress As Long '
        lpMaximumApplicationAddress As Long '
        dwActiveProcessorMask As Long       ' Gives the active processor number
        dwNumberOfProcessors As Long        ' Gives number of processors
        dwProcessorType As Long             ' Gives the processor type (386,486,586)
        dwAllocationGranularity As Long     '
        dwReserved As Long                  '
End Type

'---------------------------------------------------------------------------
' API declaration to get information about the drives.
'---------------------------------------------------------------------------
Private Declare Function apiDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function apiDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
'Private Declare Function apiFastFreeSpace Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long
Private Declare Function apiGetDrives Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function apiMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Type MONITOR_INFO_1
        pName As String
End Type
Private Type MONITOR_INFO_2
        pName As String
        pEnvironment As String
        pDLLName As String
End Type

'---------------------------------------------------------------------------
' API declarations to obtain the Windows version and the type of
' keyboard.
'---------------------------------------------------------------------------
Private Declare Function apiGetVersion Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        ' size of 'Type' = (5 x 4 bytes) =  20 bytes (the 5 Longs)
        '                                  128 bytes (fixed-length string)
        '                                 ----- +
        '                                  148 bytes
        
        dwOSVersionInfoSize As Long         ' Has to be set to size of 'type'= 148
        dwMajorVersion As Long              ' Gives the Major version
        dwMinorVersion As Long              ' Gives the Minor version
        dwBuildNumber As Long               ' Gives the buildnumber (I don't use it)
        dwPlatformId As Long                ' Gives the operating system.
        szCSDVersion As String * 128        '
End Type

Private Declare Function apiKeyboardType Lib "user32" Alias "GetKeyboardType" (ByVal nTypeFlag As Long) As Long

'---------------------------------------------------------------------------
' API declarations to obtain the Windows Display
'
'---------------------------------------------------------------------------

Private Type RECT
           Left As Long
           Top As Long
           Right As Long
           Bottom As Long
End Type
 
Private Const SPI_GETWORKAREA = 48
 
Private Declare Function apiDisplayParametersInfo Lib "user32" _
           Alias "SystemParametersInfoA" (ByVal uAction As Long, _
           ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) _
           As Long

 '****************************************
 ' API to Find the services
 '*******************************************************
  Private Type SERVICE_STATUS
            dwServiceType As Long
            dwCurrentState As Long
            dwControlsAccepted As Long
            dwWin32ExitCode As Long
            dwServiceSpecificExitCode As Long
            dwCheckPoint As Long
            dwWaitHint As Long
         End Type
 
  Private Type ENUM_SERVICE_STATUS
            lpServiceName As Long
            lpDisplayName As Long
            ServiceStatus As SERVICE_STATUS
         End Type
         '*******************************************************
         'Constants
         '*******************************************************
         Const ERROR_MORE_DATA = 234
         Const SERVICE_ACTIVE = &H1
         Const SERVICE_INACTIVE = &H2
         Const SERVICE_COMPLETE_LIST = &H3
         Const SC_MANAGER_ENUMERATE_SERVICE = &H4
         Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
         Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
         Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS _
                     + SERVICE_WIN32_SHARE_PROCESS
        
         Private Declare Function OpenSCManager Lib "advapi32.dll" _
               Alias "OpenSCManagerA" _
            (ByVal lpMachineName As String, ByVal lpDatabaseName As String, _
                  ByVal dwDesiredAccess As Long) As Long
 
         Private Declare Function EnumServicesStatus Lib "advapi32.dll" _
               Alias "EnumServicesStatusA" (ByVal hSCManager As Long, _
               ByVal dwServiceType As Long, ByVal dwServiceState As Long, _
               lpServices As Any, ByVal cbBufSize As Long, _
               pcbBytesNeeded As Long, lpServicesReturned As Long, _
               lpResumeHandle As Long) As Long
 
         Private Declare Function CloseServiceHandle Lib "advapi32.dll" _
               (ByVal hSCObject As Long) As Long
 
         Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" _
               (szDest As String, szcSource As Long) As Long


Public Sub DriveInfo(ByVal strRoot As String, ByRef lngTotalSpace As Long, ByRef lngFreeSpace As Long)
'---------------------------------------------------------------------------
' SUB: DriveInfo
'
' This Sub returns the amount of total disk space and free disk space. The
' API call return the number of Clusters, Free Clusters, Sectors per cluster,
' and Bytes per cluster. By multiplying these values you can get the
' required information.
'
' lngTotalSpace and lngFreeSpace give the amount of space in BYTES!
'   (see PutPoints)
'
' IN: strRoot       - String containing the root of the drive you want to
'                     check. (e.g. "A:\", or "C:\")
'
' OUT: lngTotalSpace - Long containing the total disk space of the drive.
'      lngFreeSpace  - Long containing the amount of free disk space.
'
' If the API call fails, Zero is returned in both variables.
'---------------------------------------------------------------------------
'
Dim TotalClusters As Long
Dim FreeClusters As Long
Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim Maxlong As Long
Dim KTotCluster As Long
Dim KFreeClust As Long


If apiDiskFreeSpace(strRoot, SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters) Then
'   If the call succeeds return the asked amount of diskspace.

Maxlong = SectorsPerCluster * BytesPerSector
KTotCluster = TotalClusters / 1024
Maxlong = Maxlong * KTotCluster

    If Maxlong < 2147483646 Then
        lngTotalSpace = SectorsPerCluster * BytesPerSector
        lngTotalSpace = lngTotalSpace * KTotCluster
        KFreeClust = FreeClusters / 1024
        lngFreeSpace = SectorsPerCluster * BytesPerSector * KFreeClust
    Else
        lngTotalSpace = 0
        lngFreeSpace = 0
    End If
Else
'   Otherwise return zero.
    lngTotalSpace = 0
    lngFreeSpace = 0
End If

End Sub


Public Sub Drives(ByRef intRemovable As Integer, ByRef intNotRemovable As Integer, ByRef intCD As Integer, ByRef intRAM As Integer, ByRef intNetwork As Integer)
'---------------------------------------------------------------------------
' SUB: Drives
'
' Returns the number of removable, fixed, CD-ROM, RAM, and Network drives
' that are connected to your computer.
'
' THIS FUNCTION USES THE DRIVETYPE FUNCTION, SO IF YOU MODIFY THAT FUNCTION
' YOU MUST ALSO MODIFY THIS FUNCTION.
'
' OUT:  intRemovable        - Integer containing the number of removable drives
'       intNotRemovable     - Integer containing the number of fixed drives
'       intCD               - Integer containing the number of CD drives
'       intRAM              - Integer containing the number of RAM disks
'       int Network         - Integer containing the number of Network drives
'
'---------------------------------------------------------------------------
'
Dim Retrn As Long
Dim buffer As Long
Dim temp As String
Dim intI As Integer
Dim Read As String
Dim Counter As Integer
buffer = 10

Again:
temp = Space$(buffer)
Retrn = apiGetDrives(buffer, temp)
' Call the API function.

If Retrn > buffer Then ' If the API returned a value that is bigger than Buffer,
    buffer = Retrn     ' than the Buffer isn't big enough to hold the information.
    GoTo Again         ' In that case adjust the Buffer to the right size (returned by
End If                 ' the API) and try again.


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
    Read = Mid$(temp, intI, 3)
    
    Select Case DriveType(Read)
        Case "Removable"
            intRemovable = intRemovable + 1
        Case "Local"
            intNotRemovable = intNotRemovable + 1
        Case "Network"
            intNetwork = intNetwork + 1
        Case "CD-ROM"
            intCD = intCD + 1
        Case "RAM-disk"
            intRAM = intRAM + 1
    End Select

Next

End Sub

Public Function DriveType(ByVal strRoot As String) As String
'---------------------------------------------------------------------------
' FUNCTION: DriveType
'
' This function returns information about the drive you asked for. It will
' return whether the drive is a Removable drive, a non-removable (fixed)
' drive, a CD-ROM drive, a RAM drive or a Network drive.
'
' IN:  strRoot      - String containing the root of a drive. (e.g. "C:\")
'
' OUT: DriveType    - String containing type of drive.
'
' If the function fails a empty string is returned.
'
' You can also re-program this Function so that it doens't return a string,
' but it returns the value. That can be easier if you want to work with
' the returned information. I let it return a string, so that I can print
' it.
'
' THE DRIVES FUNCTION USES THIS FUNCTION, SO IF YOU MODIFY THIS FUNCTION,
' YOU ALSO HAVE TO MODIFY THAT FUNCTION!
'
'---------------------------------------------------------------------------
'
Dim lngType As Long
Const DRIVE_CDROM = 5       ' Some API constants required to
Const DRIVE_FIXED = 3       ' get the difference between the
Const DRIVE_RAMDISK = 6     ' drive types.
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2

lngType = apiDriveType(strRoot)
' The API returns a value in lngType. Use the Constants to
' make the strings.

Select Case lngType
    Case DRIVE_REMOVABLE
        DriveType = "Removable"
    Case DRIVE_FIXED
        DriveType = "Local"
    Case DRIVE_REMOTE
        DriveType = "Network"
    Case DRIVE_CDROM
        DriveType = "CD-ROM"
    Case DRIVE_RAMDISK
        DriveType = "RAM-disk"
    Case Else
        DriveType = ""   ' If the API returns an error, we return a empty string
End Select

End Function

Public Function ShowFileSystemType(drvspec As String) As String

'Available return types include FAT, NTFS, and CDFS

  Dim fso
  Dim d
  
  On Error GoTo errFile
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set d = fso.GetDrive(drvspec)
  
  ShowFileSystemType = d.FileSystem
  Exit Function
  
errFile:
  ShowFileSystemType = "Err.SystemFile"
  
End Function

Public Function RetStringKb(ByVal lngNumber As Long) As String

    RetStringKb = CStr(Round(lngNumber)) + " Kb"

End Function

Public Function FastDiskSpace(ByVal strRoot As String) As Long
'---------------------------------------------------------------------------
' FUNCTION: FastDiskSpace
'
' Returns the amount of free disk space. See also the DriveInfo function.
' This function is faster because here we don't have to multiply some
' values in order to get the diskspace.
'
' The amount of free space is given in BYTES! (See PutPoints)
'
' IN:  strRoot       - String containing the root of the drive you want to
'                      check out.
'
' OUT: FastDiskSpace - Long containing the amount of free space (in bytes)
'---------------------------------------------------------------------------
'
Dim strCurrent As String
On Error GoTo Bliep              '(Dutch variation of Beep, means an error)

strCurrent = CurDir              ' Save the current drive
ChDrive strRoot                  ' Change to the requested drive

'FastDiskSpace = apiFastFreeSpace 'Get the free space

ChDrive Left$(strCurrent, 2)     ' Return to the saved drive
ChDir strCurrent                 ' Return to the saved directory

Exit Function

Bliep:
' If the drive wasn't ready or something
FastDiskSpace = 0            ' Return zero as free disk space
ChDrive Left$(strCurrent, 2) ' Retur to the saved drive
ChDir strCurrent             ' and directory

End Function

Public Sub FreeMemory(ByRef btePercentUsed As Byte, ByRef lngTotalRam As Long, _
    ByRef lngFreeRam As Long, ByRef lngTotPagFil As Long, ByRef lngAvPagFil As Long, _
    ByRef lngTotVirFil As Long, ByRef lngAvVirFil As Long)
'---------------------------------------------------------------------------
' SUB: FreeMemory
'
' Returns information about your RAM (-memory).
' lngTotalRam en lngFreeRam return the amount of RAM in Kbytes!
'
' OUT:  btePercentUsed      - Byte that gives an indication of used RAM in %
'       lngTotalRam         - Long containing the amount of total RAM
'       lngFreeRam          - Long containing the aomunt of free RAM
'       dwLength As Long                    ' This need to be set at the size of this 'Type'  = 32
'
'        dwMemoryLoad As Long                ' Gives global indication of used RAM (in %)
'        dwTotalPhys As Long                 ' Gives total RAM of the computer
'        dwAvailPhys As Long                 ' Gives the amount of free RAM
'        dwTotalPageFile As Long             '
'        dwAvailPageFile As Long             '
'        dwTotalVirtual As Long              '
'        dwAvailVirtual As Long

'---------------------------------------------------------------------------
'
Dim Memory As MEMORYSTATUS

Memory.dwLength = 32
' This must be set to the size of the structure before the call

apiMemStatus Memory
' Call the API. This function fills the Memory structure (Type) with
' a lot of information. I only use three parts of it.

' Fill the variables with the desired values.
btePercentUsed = Memory.dwMemoryLoad
lngTotalRam = Memory.dwTotalPhys / 1024
lngFreeRam = Memory.dwAvailPhys / 1024
lngTotPagFil = Memory.dwTotalPageFile / 1024
lngAvPagFil = Memory.dwAvailPageFile / 1024
lngTotVirFil = Memory.dwTotalVirtual / 1024
lngAvVirFil = Memory.dwAvailVirtual / 1024

End Sub
Public Function FunctionKeys() As Byte
'---------------------------------------------------------------------------
' FUNCTION: FunctionKeys
'
' Returns the number of function keys your keyboard has.
' See the KeyboardType function for more information about your keyboard.
'---------------------------------------------------------------------------
'
FunctionKeys = apiKeyboardType(2)

End Function
Public Function KeyboardType() As String
'---------------------------------------------------------------------------
' FUNCTION: KeyboardType
'
' Returns a string containing the type of Keyboard you use.
'
'---------------------------------------------------------------------------
'
Dim intBuffer As Long

intBuffer = apiKeyboardType(0)
' Call the API. The zero specifies that I want to get information
' about the keyboard type. The FunctionKeys function uses the
' same API call, only with a "2" specified.


' These are all constants I've found somewhere, so I can't explain
' it. It's just true.
Select Case intBuffer
Case 1
    KeyboardType = "IBM PC/XT or compatible (83 key)"
Case 2
    KeyboardType = "Olivetti ""ico"" (102 key)"
Case 3
    KeyboardType = "IBM PC/AT or compatible (84 key)"
Case 4
    KeyboardType = "IBM enhanced (101 or 102 key)"
Case 5
    KeyboardType = "Nokia 1050 or compatible"
Case 6
    KeyboardType = "Nokia 9140 or compatible"
Case 7
    KeyboardType = "Japanese"
End Select

End Function
Public Function TempDir() As String
'---------------------------------------------------------------------------
' FUNCTION: TempDir
'
' Get the Temporary directory windows uses.
'
' OUT: TempDir  - String containing the directory.
'
' If the function fails a empty string is returned.
'---------------------------------------------------------------------------
'
Dim Bufstr As String
Bufstr = Space$(50)


'---------------------------------------------------------------------------
' Call the API and remove the spaces using RTrim. Next, remove the terminating
' character using StripTerminator, and add a backslash, when if it wasn't
' already there.
'---------------------------------------------------------------------------
If apiTempDir(50, Bufstr) > 0 Then
    TempDir = Bufstr
    TempDir = RTrim(TempDir)
    TempDir = StripTerminator(TempDir)
    
    If Right$(TempDir, 1) <> "\" Then
        TempDir = TempDir + "\"
    End If
    
Else
    TempDir = ""
End If
        
End Function
Public Function SetEnv(SetVall As String) As String
'---------------------------------------------------------------------------
' FUNCTION: SetEnv
'
' Get the Path directory windows uses.
'
' OUT: SetEnv  - String containing the directory.
'
' If the function fails a empty string is returned.
'---------------------------------------------------------------------------
'
    Dim strEnvVarName As String
    Dim strVarBuffer As String * 4096

    strEnvVarName = SetVall
    apiSetEnv strEnvVarName, strVarBuffer, Len(strVarBuffer)
    SetEnv = Trim$(strVarBuffer)
        
End Function

Public Function SystemDir() As String
'---------------------------------------------------------------------------
' FUNCTION: SystemDir
'
' Gets the WINDOWS\SYSTEM directory.
'
' Returns a string containing the full path, ends with a "\". If the
' call fails a empty string is returned.
'---------------------------------------------------------------------------
'
Dim Bufstr As String
Bufstr = Space$(50)


'---------------------------------------------------------------------------
' Call the API and remove the spaces using RTrim. Remove the terminating
' character and add a backslash when it isn't already there.
'---------------------------------------------------------------------------
If apiSysDir(Bufstr, 50) > 0 Then
    SystemDir = Bufstr
    SystemDir = RTrim(SystemDir)
    SystemDir = StripTerminator(SystemDir)
    
    If Right$(SystemDir, 1) <> "\" Then
        SystemDir = SystemDir + "\"
    End If
    
Else
    SystemDir = ""
End If
        
End Function
Public Function computername() As String
'---------------------------------------------------------------------------
' FUNCTION: ComputerName
'
' This function retrieves the Computer name (is the computer is connected to
' a network) and removes the terminating character that the Windows API
' returns.
'
' OUT:  ComputerName - String containing the name of the computer. If the
'                      API call fails, an empty string is returned.
'
'---------------------------------------------------------------------------
'
Dim Bufstr As String
Bufstr = Space$(50)


'---------------------------------------------------------------------------
' Call the API and remove the empty spaces behind the name using RTrim.
' Afterwards remove the terminating character.
'---------------------------------------------------------------------------
If apiCompName(Bufstr, 50) > 0 Then
    computername = Bufstr
    computername = RTrim(computername)
    computername = StripTerminator(computername)
Else
    computername = ""
End If
        
End Function
Private Function StripTerminator(ByVal strString As String) As String
'---------------------------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'
'
' THIS FUNCTION I GOT FROM THE SETUP PROJECT THAT CAME WITH VISUAL BASIC 4.
'---------------------------------------------------------------------------
'
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If

End Function
Public Function username() As String
'---------------------------------------------------------------------------
' FUNCTION: UserName
'
' Get the name of the user.
'
' OUT: UserName     - String containing the name of the user of the computer.
'
' If the function fails, an empty string is returned.
'---------------------------------------------------------------------------
'
Dim Bufstr As String
Bufstr = Space$(50)


'---------------------------------------------------------------------------
' Call the API, remove the spaces using RTrim, remove the NULL-character
' using StripTerminator.
'---------------------------------------------------------------------------
If apiUserName(Bufstr, 50) > 0 Then
    username = Bufstr
    username = RTrim(username)
    username = StripTerminator(username)
Else
    username = ""
End If
        
End Function
Public Function SerialNumber(ByVal strRoot As String) As String
'---------------------------------------------------------------------------
' FUNCTION: SerialNumber
'
' Returns the serial number of a drive. It returns the number exactly the
' same as DOS does (hexadecimal value e.g. : 1104-224E)
'
' IN:   strRoot      - String containing the root of a drive (e.g. "A:\").
'
' OUT:  SerialNumber - String containing the serial number.
'
' If the function fails (because the drive wasn't ready or something), the
' function returns "0000-0000" as the serial number.
'
'---------------------------------------------------------------------------
'
Dim VolLabel As String
Dim VolSize As Long
Dim SerNum As Long
Dim MaxLen As Long
Dim Flags As Long
Dim Name As String
Dim NameSize As Long
Dim Check As String

If apiSerialNumber(strRoot, VolLabel, VolSize, SerNum, MaxLen, Flags, Name, NameSize) Then
' This function returns a lot more, but I can get that information via another function.

    Check = Format(Hex(SerNum), "00000000")
    ' Make sure that the length = 8. So convert "123456" to "00123456"
    
    SerialNumber = Left$(Check, 4) + "-" + Right$(Check, 4)
    ' Split the number in two parts of four and add a "-" between them.

Else
    ' Return "0000-0000" is the function fails.
    SerialNumber = "0000-0000"

End If

End Function

Public Function VolumeLabel(ByVal strRoot As String) As String
'---------------------------------------------------------------------------
' FUNCTION: VolumeLabel
'
' Returns the VolumeLabel of a drive. This function doesn't need an API
' Call, because the Visual Basic command "Dir" returns this label.
'
' IN: strRoot       - String containing the root of the drive you want
'                     the Volume Label of.
'
' OUT: VolumeLabel  - String containing the Volume Label.
'
' If the function fails an empty string is returned. When the drive hasn't
' got a name, "NoName" is returned as the Volume label.
'---------------------------------------------------------------------------
'

On Error GoTo Further

VolumeLabel = Dir(strRoot, vbVolume)
VolumeLabel = StripTerminator(VolumeLabel)
' Get the volume label and remove the NULL character.

If VolumeLabel = "" Then VolumeLabel = "NoName"
' Set the label to "NoName", when the drive hasn't got a name.


Exit Function


Further:
VolumeLabel = ""

' If the function fails, return an empty string.

End Function

Public Function WinDir() As String
'---------------------------------------------------------------------------
' FUNCTION: WinDir
'
' Returns the Windows directory (Mostly "C:\WINDOWS\")
'
' If the function fails an empty string is returned.
'---------------------------------------------------------------------------
'
Dim Bufstr As String
Bufstr = Space$(50)


'---------------------------------------------------------------------------
' Call the API, remove the extra spaces using RTrim, remove the NULL-character
' using StripTerminator, and add a backslash.
'---------------------------------------------------------------------------
If apiWindDir(Bufstr, 50) > 0 Then
    WinDir = Bufstr
    WinDir = RTrim(WinDir)
    WinDir = StripTerminator(WinDir)
    
    If Right$(WinDir, 1) <> "\" Then
        WinDir = WinDir + "\"
    End If
    
Else
    WinDir = ""
End If
        
End Function
Public Sub WinVer(ByRef intMajor As Integer, ByRef intMinor As Integer, _
ByRef strPlatform As String, ByRef BuildNb As Long, ByRef CSDVer As String)
'---------------------------------------------------------------------------
' SUB: WinVer
'
' This sub returns information about the operating system, and about
' the Windows Version.
'
' e.g. Windows 3.11
'      The sub will return:
'         - intMajor    = 3
'         - intMinor    = 11
'         - strPlatform = Windows 3.11
'
' OUT:  intMajor        - Integer containing the major version of Windows.
'       intMinor        - Integer containing the minor version of windows.
'
' strPlatfrom returns one of the following :
'   Windows 95
'   Windows NT
'   Windows + Version
'
'dwOSVersionInfoSize As Long         ' Has to be set to size of 'type'= 148
'        dwMajorVersion As Long              ' Gives the Major version
'        dwMinorVersion As Long              ' Gives the Minor version
'        dwBuildNumber As Long               ' Gives the buildnumber (I don't use it)
'        dwPlatformId As Long                ' Gives the operating system.
'        szCSDVersion As String * 128

' If the call fails intMajor = 0, intMinor = 0, and strPlatform = ""
'---------------------------------------------------------------------------
'
Dim OSystem As OSVERSIONINFO

OSystem.dwOSVersionInfoSize = 148
' The size of the structure must be set before the call.

If apiGetVersion(OSystem) Then
' Call the API. It fills the OSystem type.

    intMajor = OSystem.dwMajorVersion   ' Store the Major version in intMajor
    intMinor = OSystem.dwMinorVersion   ' Store the Minor version in intMinor
    BuildNb = OSystem.dwBuildNumber
    CSDVer = OSystem.szCSDVersion
    
    
    Select Case OSystem.dwPlatformId    ' Set strPlatform
    Case 0
        strPlatform = "Windows " + CStr(intMajor) + "." + CStr(intMinor)
    Case 1
        strPlatform = "Windows 95"
    Case 2
        strPlatform = "Windows NT"
    End Select

Else
' The call failed, set the values to zero
    intMajor = 0
    intMinor = 0
    strPlatform = ""
    BuildNb = 0
    CSDVer = ""
    
End If

End Sub

Public Sub SystemInfo(ByRef strProcessor As String, ByRef lngNumOfProcessors As Long, ByRef lngActiveProcessor As Long, _
    ByRef lngOemId As Long, ByRef lngPagSiz As Long, ByRef lngMinAdr As Long, _
    ByRef lngMaxAdr As Long, ByRef lngAlloc As Long, ByRef lngResv As Long)
'---------------------------------------------------------------------------
' SUB: SystemInfo
'
' This Sub returns the number of processors, the active processor and
' the type of the processor.
'
' OUT:  strProcessor        - String containing the type of processor.
'       lngNumOfProcessors  - Long containing the number of processors.
'       lngActiveProcessor  - Long containing the number of the active
'                             processor (mostly 1)
'        dwOemID As Long
'        dwPageSize As Long                  ' Must be set at the size of this 'Type'
'        lpMinimumApplicationAddress As Long '
'        lpMaximumApplicationAddress As Long '
'        dwAllocationGranularity As Long     '
'        dwReserved As Long
'---------------------------------------------------------------------------
'
Dim SI As SYSTEM_INFO

apiSystemInfo SI
' The API call fills the SI type with a lot of information
' but I only use three parts of it.
        
lngActiveProcessor = SI.dwActiveProcessorMask
lngNumOfProcessors = SI.dwNumberOfProcessors
lngMinAdr = SI.lpMinimumApplicationAddress
lngMaxAdr = SI.lpMaximumApplicationAddress
lngOemId = SI.dwOemID
lngPagSiz = SI.dwPageSize
lngAlloc = SI.dwAllocationGranularity
lngResv = SI.dwReserved

Select Case SI.dwProcessorType
Case 386
    strProcessor = "80386"           ' Return the processor type.
Case 486                             ' Windows 95 only recognises
    strProcessor = "80486"           ' these three.
Case 586
    strProcessor = "Intel Pentium"
Case 4000
    strProcessor = "MIPS R4000"
Case 21064
    strProcessor = "DEC Alpha 21064"
Case Else
    strProcessor = "(unknown)"

End Select


End Sub

Public Sub DisplayParametersInfo(ByRef VideoX As String, ByRef VideoY As String)

Dim lRet As Long
Dim apiRECT As RECT
 
lRet = apiDisplayParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)
 
If lRet Then
    'VideoX0 = "WorkAreaLeft: " & apiRECT.Left      'not use
    'VideoY0 = "WorkAreaTop: " & apiRECT.Top        'not use
    VideoX = Str(apiRECT.Right - apiRECT.Left)
    VideoY = Str(apiRECT.Bottom - apiRECT.Top)
Else
    VideoX = "Cannot Detected"
    VideoY = "Cannot Detected"
End If


End Sub

Public Sub GetServiceInfo(lngServiceType As Long, LogDir As String)
            Dim hSCM As Long
            Dim lpEnumServiceStatus() As ENUM_SERVICE_STATUS
            Dim lngServiceStatusInfoBuffer As Long
            Dim strServiceName As String * 250
            Dim lngBytesNeeded As Long
            Dim lngServicesReturned As Long
            Dim hNextUnreadEntry As Long
            Dim lngStructsNeeded As Long
            Dim lngResult As Long
            Dim I As Long
            Dim ServName As String
            Dim ServDisp As String
            Dim X As String
            Dim J As Integer
            Dim startmode As String
            Dim ServCount As String
            
            
            
            'Open connection to Service Control Manager.
            hSCM = OpenSCManager(vbNullString, vbNullString, _
                                 SC_MANAGER_ENUMERATE_SERVICE)
 
            If hSCM = 0 Then
               MsgBox "OpenSCManager failed. LastDllError = " _
               & CStr(Err.LastDllError)
               Exit Sub
            End If
 
            'Get buffer size (bytes) without passing a buffer
            'and make sure starts at 0th entry.
            hNextUnreadEntry = 0
            lngResult = EnumServicesStatus(hSCM, _
                                           SERVICE_WIN32, _
                                           lngServiceType, _
                                           ByVal &H0, _
                                           &H0, _
                                           lngBytesNeeded, _
                                           lngServicesReturned, _
                                           hNextUnreadEntry)
 
            'We should receive MORE_DATA error.
            If Not Err.LastDllError = ERROR_MORE_DATA Then
               MsgBox "LastDLLError = " & CStr(Err.LastDllError)
               Exit Sub
            End If
 
            'Calculate the number of structures needed.
            lngStructsNeeded = _
            lngBytesNeeded / Len(lpEnumServiceStatus(0)) + 1
 
            'Redimension the array according to our calculation.
            ReDim lpEnumServiceStatus(lngStructsNeeded - 1)
 
            'Get buffer size in bytes.
            lngServiceStatusInfoBuffer = _
               lngStructsNeeded * Len(lpEnumServiceStatus(0))
 
            'Get services information starting entry 0.
            hNextUnreadEntry = 0
            lngResult = EnumServicesStatus(hSCM, _
                                          SERVICE_WIN32, _
                                          lngServiceType, _
                                          lpEnumServiceStatus(0), _
                                          lngServiceStatusInfoBuffer, _
                                          lngBytesNeeded, _
                                          lngServicesReturned, _
                                          hNextUnreadEntry)
 
            If lngResult = 0 Then
               MsgBox "EnumServicesStatus failed. LastDllError = " _
                  & CStr(Err.LastDllError)
               Exit Sub
            End If
 
            'Get the strings and display them.
            'With Form1
            '   .List1.Clear
            '   .List1.AddItem "Display Names"
            '   .List2.Clear
            '   .List2.AddItem "Service Names"
               For I = 0 To lngServicesReturned - 1
                  lngResult = lstrcpy(ByVal strServiceName, _
                     ByVal lpEnumServiceStatus(I).lpDisplayName)
            '      .List1.AddItem strServiceName
                    ServName = strServiceName
                    'del the null caracter
                    X = ""
                    For J = 1 To Len(ServName)
                    If Asc(Mid(ServName, J, 1)) Then
                        X = X + Mid(ServName, J, 1)
                    Else
                        Exit For
                    End If
                    Next J
                    ServName = X
                    
                    lngResult = lstrcpy(ByVal strServiceName, _
                     ByVal lpEnumServiceStatus(I).lpServiceName)
            '      .List2.AddItem strServiceName
                    ServDisp = strServiceName
                   
                    
                  'verify the start mode of the service in the registry
                  startmode = GetDWORDValue("HKEY_LOCAL_MACHINE\system\CurrentControlSet\Services\" + ServDisp, "Start")
                  Select Case startmode
                    Case "2"
                        startmode = " - Automatique"
                    Case "3"
                        startmode = " - Manual"
                    Case "4"
                        startmode = " - Not Active"
                  End Select
                  
                  ServCount = " - " + GetStringValue("HKEY_LOCAL_MACHINE\system\CurrentControlSet\Services\" + ServDisp, "ObjectName")
                  
                  If lngServiceType = &H1 Then
                    X = INIWrite("SERVICES STARTED", ServDisp, ServName + startmode + ServCount, LogDir)
                  Else
                    X = INIWrite("SERVICES STOPED", ServDisp, ServName + startmode + ServCount, LogDir)
                  End If
                  
               Next
            'End With
 
            'Clean up.
            CloseServiceHandle (hSCM)
 
End Sub

Public Sub NetWorkInf(LogDir As String)
'********************************************
'Found Ip Adr and computername
'********************************************
 
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim I, J As Integer
    Dim ip_address As String
    Dim X As String
    
      SocketsInitialize LogDir
 
    If gethostname(hostname, 256) = SOCKET_ERROR Then
        X = INIWrite("NetWork", "ComputerName", "Windows Sockets error " & Str(WSAGetLastError()), LogDir)

        Exit Sub
    Else
        hostname = Trim$(hostname)
    End If
 
    hostent_addr = gethostbyname(hostname)
 
    If hostent_addr = 0 Then
        X = INIWrite("NetWork", "ComputerName", "Winsock.dll is not responding", LogDir)
        Exit Sub
    End If
 
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
 
    'MsgBox hostname
    X = INIWrite("NetWork", "ComputerName", hostname, LogDir)
    
    'get all of the IP address if machine is  multi-homed
    J = 1
    Do
              ReDim temp_ip_address(1 To host.hLength)
              RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

              For I = 1 To host.hLength
                  ip_address = ip_address & temp_ip_address(I) & "."
              Next
              ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

              X = INIWrite("NetWork", "IP " + Str(J), ip_address, LogDir)
             
              J = J + 1
              ip_address = ""
              host.hAddrList = host.hAddrList + LenB(host.hAddrList)
              RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)

SocketsCleanup

End Sub


Private Sub SocketsInitialize(LogDir As String)
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    Dim X As String
    
 
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
 
    If iReturn <> 0 Then
        X = INIWrite("NetWork", "WINSOCK_Err", "Winsock.dll is not responding", LogDir)
        End
    End If
 
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
 
        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        X = INIWrite("NetWork", "WINSOCK_Err", sMsg, LogDir)
        End
    End If
 
    'iMaxSockets is not used in winsock 2. So the following check is only
    'necessary for winsock 1. If winsock 2 is requested,
    'the following check can be skipped.
 
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        X = INIWrite("NetWork", "WINSOCK_Err", sMsg, LogDir)
        End
    End If
 
End Sub

Private Sub SocketsCleanup()
           Dim lReturn As Long
           lReturn = WSACleanup()
           If lReturn <> 0 Then
             MsgBox "Socket error " & Trim$(Str$(lReturn)) & _
      " occurred in Cleanup"
             End
           End If
End Sub


Private Function lobyte(ByVal wParam As Integer)
 
    lobyte = wParam And &HFF&
 
End Function

Private Function hibyte(ByVal wParam As Integer)
 
    hibyte = wParam \ &H100 And &HFF&
 
End Function


Function GetWorkstationInfo(LogDir As String)
'---------------------------------------------------------------------------
' Found Domain/Workgroup and Logon Domain
'---------------------------------------------------------------------------
            Dim ret As Long, buffer(512) As Byte, I As Integer
            Dim wk101 As WKSTA_INFO_101, pwk101 As Long
            Dim wk1 As WKSTA_USER_INFO_1, pwk1 As Long
            Dim cbusername As Long, username As String
            Dim computername As String, langroup As String, logondomain As _
               String
            Dim X As String
            
 
            ' Clear all of the display values.
            computername = "": langroup = "": username = "": logondomain = ""
 
            ' Windows 95 or NT - call WNetGetUser to get the name of the user.
            username = Space(256)
            cbusername = Len(username)
            ret = WNetGetUser(ByVal 0&, username, cbusername)
            If ret = 0 Then
               ' Success - strip off the null.
               username = Left(username, InStr(username, Chr(0)) - 1)
            Else
               username = ""
            End If
 
         '================================================================
         ' The following section works only under Windows NT
         '================================================================
 
            'NT only - call NetWkstaGetInfo to get computer name and lan group
            ret = NetWkstaGetInfo(ByVal 0&, 101, pwk101)
            RtlMoveMemory wk101, ByVal pwk101, Len(wk101)
            lstrcpyW buffer(0), wk101.wki101_computername
            ' Get every other byte from Unicode string.
            I = 0
            Do While buffer(I) <> 0
               computername = computername & Chr(buffer(I))
               I = I + 2
            Loop
            lstrcpyW buffer(0), wk101.wki101_langroup
            I = 0
            Do While buffer(I) <> 0
               langroup = langroup & Chr(buffer(I))
               I = I + 2
            Loop
            ret = NetApiBufferFree(pwk101)
 
            ' NT only - call NetWkstaUserGetInfo.
            ret = NetWkstaUserGetInfo(ByVal 0&, 1, pwk1)
            RtlMoveMemory wk1, ByVal pwk1, Len(wk1)
            lstrcpyW buffer(0), wk1.wkui1_logon_domain
            I = 0
            Do While buffer(I) <> 0
               logondomain = logondomain & Chr(buffer(I))
               I = I + 2
            Loop
            ret = NetApiBufferFree(pwk1)
 
         '================================================================
         'End NT-specific section
         '================================================================
            X = INIWrite("NetWork", "Domain", langroup, LogDir)
            X = INIWrite("NetWork", "LogonDomain", logondomain, LogDir)
            'MsgBox langroup + " - " + username + " - " + logondomain
         End Function

Public Sub ReadHost(LogDir As String)
'---------------------------------------------------------------------------
' Read the Host files
'---------------------------------------------------------------------------
Dim HostF As String
Dim Ligne As String
Dim X As String
Dim I As Integer
    
    I = 1
    HostF = WinDir + "system32\drivers\etc\hosts."

    On Error GoTo NoFile
    Open HostF For Input As #1
    On Error GoTo 0
    Do While Not EOF(1)
        Line Input #1, Ligne
        If Left(Ligne, 1) <> "#" And Ligne <> "" Then
            X = INIWrite("Hosts", "H" + Format(I, "0000"), Ligne, LogDir)
            I = I + 1
        End If
    Loop
    Close #1
    If I = 1 Then
        X = INIWrite("Hosts", "Hosts", "Empty", LogDir)
    End If
    Exit Sub

NoFile:
X = INIWrite("Hosts", "Error", "No Host File", LogDir)

End Sub


Public Sub ReadLMHost(LogDir As String)
'---------------------------------------------------------------------------
' Read the LMHost files
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
' Read the Host files
'---------------------------------------------------------------------------
Dim HostF As String
Dim Ligne As String
Dim X As String
Dim I As Integer
    
    I = 1
    HostF = WinDir + "system32\drivers\etc\lmhosts."

    On Error GoTo NoFile
    Open HostF For Input As #1
    On Error GoTo 0
    Do While Not EOF(1)
        Line Input #1, Ligne
        If Left(Ligne, 1) <> "#" And Ligne <> "" Then
            X = INIWrite("LmHost", "H" + Format(I, "0000"), Ligne, LogDir)
            I = I + 1
        End If
    Loop
    Close #1
    If I = 1 Then
        X = INIWrite("LmHost", "LmHosts", "Empty", LogDir)
    End If
    
    Exit Sub

NoFile:
X = INIWrite("LmHost", "Error", "No LmHost File", LogDir)

End Sub

Public Sub InfoNetWorkAdr(ByRef SrvDNS As String, ByRef SrvDHCP As String, _
    ByRef AdrMask As String, ByRef AdrGate As String)
    
    Dim tempkey As String
    Dim tempNet As String
    Dim I As Integer
    Dim DHCP_Ok As String
    
    SrvDNS = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "NameServer")
    
    'find the networkcard
    tempNet = ""
    tempkey = GetBinaryValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Linkage", "Bind")
    tempkey = Mid(tempkey, 9, Len(tempkey))
    
    For I = 1 To Len(tempkey)
        If Asc(Mid(tempkey, I, 1)) = 92 Then
            Exit For
        End If
        tempNet = tempNet + Mid(tempkey, I, 1)
    Next I
    
    'find here you are on a dhcp network
    DHCP_Ok = GetDWORDValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "EnableDHCP")
    
    If DHCP_Ok = "1" Then
        SrvDHCP = GetBinaryValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "DhcpServer")
        AdrMask = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "DhcpSubnetMask")
        AdrGate = GetBinaryValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "DhcpDefaultGateway")
    Else
        SrvDHCP = "No DHCP"
        AdrMask = GetBinaryValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "SubnetMask")
        AdrGate = GetBinaryValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & tempNet & "\Parameters\Tcpip", "DefaultGateway")
    
    End If
    
    
End Sub

