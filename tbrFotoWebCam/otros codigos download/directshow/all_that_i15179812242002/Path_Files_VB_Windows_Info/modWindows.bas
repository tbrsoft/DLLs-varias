Attribute VB_Name = "modWindows"
Option Explicit

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(0 To 0) As LUID_AND_ATTRIBUTES
End Type

Private Enum PlatformType   ' dwPlatformId
  VER_PLATFORM_WIN32s = 0        ' Unknown Version
  VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 3.1/95/98/Me
  VER_PLATFORM_WIN32_NT = 2      ' Windows NT/2000/XP/.NET
End Enum

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As PlatformType
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type WindowsVersionInfo
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As PlatformType
        szCSDVersion As Variant
        dwFullVersion As Variant
        dwTextVersion As Variant
        dwFullTextV As Variant
End Type

Public Enum ShutDownType
  EWX_LOGOFF = &H0
  EWX_SHUTDOWN = &H1
  EWX_REBOOT = &H2
  EWX_POWEROFF = &H8     ' SHUTDOWN is better
End Enum

Public Enum ForceType
  EWX_NORMAL = &H0
  EWX_FORCEIFHUNG = &H10
  EWX_FORCE = &H4        ' better not use !
End Enum

Public Enum DIR_ID
  DIR_USER = &H28
  DIR_USER_DESKTOP = &H10
  DIR_USER_MY_DOCUMENTS = &H5
  DIR_USER_START_MENU = &HB
  DIR_USER_START_MENU_PROGRAMS = &H2
  DIR_USER_START_MENU_PROGRAMS_STARTUP = &H7
  DIR_COMMON_DESKTOP = &H19
  DIR_COMMON_DOCUMENTS = &H2E
  DIR_COMMON_START_MENU = &H16
  DIR_COMMON_START_MENU_PROGRAMS = &H17
  DIR_COMMON_START_MENU_PROGRAMS_STARTUP = &H18
  DIR_WINDOWS = &H24
  DIR_SYSTEM = &H25
  DIR_FONTS = &H14
  DIR_PROGRAM_FILES = &H26
  DIR_PROGRAM_FILES_COMMON_FILES = &H2B
End Enum

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Const NoShutDownPrivilege As String = "No ShutDown Privilege !"

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal HWND As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Function isNT2000XP() As Boolean
  Dim lpv As OSVERSIONINFO
  lpv.dwOSVersionInfoSize = Len(lpv)
  GetVersionEx lpv
  If lpv.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    isNT2000XP = True
  Else
    isNT2000XP = False
  End If
End Function

Private Function ShutDownPrivilege() As Boolean
  Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
  Const TOKEN_QUERY As Long = &H8
  Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
  Const SE_PRIVILEGE_ENABLED As Long = &H2
  Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
  Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
  Const Language_Neutral As Long = &H0
  Const User_Default_Language As Long = &H400
  Const System_Default_Language As Long = &H800
  Dim ErrorNumber As Long
  Dim ErrorMessage As String
  Dim hToken As Long
  Dim tkp As TOKEN_PRIVILEGES
  Dim tkpNULL As TOKEN_PRIVILEGES
  If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
    ShutDownPrivilege = False
    Exit Function
  End If
  LookupPrivilegeValue vbNullString, SE_SHUTDOWN_NAME, tkp.Privileges(0).pLuid
  tkp.PrivilegeCount = 1
  tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
  AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkpNULL, Len(tkpNULL)
  ErrorNumber = GetLastError
  If ErrorNumber <> 0 Then
    ErrorMessage = Space$(500)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrorNumber, User_Default_Language, ErrorMessage, Len(ErrorMessage), 0
    MsgBox Trim2(ErrorMessage), vbExclamation, "Super"
    ShutDownPrivilege = False
    Exit Function
  End If
  ShutDownPrivilege = True
End Function

Public Sub SHUTDOWN(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG, Optional ByVal SDT As ShutDownType = EWX_SHUTDOWN)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = SDT Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "Super"
    End If
  Else
    var1 = SDT Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub LOGOFF(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_LOGOFF Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "Super"
    End If
  Else
    var1 = EWX_LOGOFF Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub REBOOT(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_REBOOT Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "Super"
    End If
  Else
    var1 = EWX_REBOOT Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub POWEROFF(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_POWEROFF Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "Super"
    End If
  Else
    var1 = EWX_POWEROFF Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub LockComputer()
  If isNT2000XP Then
    LockWorkStation
  Else
    LOGOFF EWX_FORCEIFHUNG
  End If
End Sub

Public Function GetWindowsVersion() As WindowsVersionInfo
  Dim lpv As OSVERSIONINFO
  Dim wvi As WindowsVersionInfo
  Dim qwe As String
  Dim qaz As String
  Dim t As Byte
  lpv.dwOSVersionInfoSize = Len(lpv)
  GetVersionEx lpv
  qwe = ""
  For t = 1 To 128
    qaz = Mid$(lpv.szCSDVersion, t, 1)
    Select Case qaz
      Case Chr$(0), Chr$(32), Chr$(255):
        qwe = qwe & Chr$(32)
      Case Else:
        qwe = qwe & qaz
    End Select
  Next t
  Select Case lpv.dwPlatformId
    Case VER_PLATFORM_WIN32_NT
      Select Case lpv.dwMajorVersion
        Case 3
          wvi.dwTextVersion = "Windows NT 3.51"
        Case 4
          wvi.dwTextVersion = "Windows NT 4.0"
        Case 5
          Select Case lpv.dwMinorVersion
            Case 0
              wvi.dwTextVersion = "Windows 2000"
            Case 1
              wvi.dwTextVersion = "Windows XP"
            Case 2
              wvi.dwTextVersion = "Windows .NET"
            Case Else
              wvi.dwTextVersion = "Windows 2000/XP/.NET"
          End Select
        Case Else
          wvi.dwTextVersion = "Windows NT/2000/XP/.NET"
      End Select
    Case VER_PLATFORM_WIN32_WINDOWS
      Select Case lpv.dwMajorVersion
        Case 3
          wvi.dwTextVersion = "Windows 3.1"
        Case 4
          Select Case lpv.dwMinorVersion
            Case 0
              Select Case Left$(lpv.szCSDVersion, 1)
                Case "C"
                  wvi.dwTextVersion = "Windows 95 C"
                Case "B"
                  wvi.dwTextVersion = "Windows 95 B"
                Case Else
                  wvi.dwTextVersion = "Windows 95"
              End Select
            Case 10
              Select Case Left$(lpv.szCSDVersion, 1)
                Case "A"
                  wvi.dwTextVersion = "Windows 98 SE"
                Case Else
                  wvi.dwTextVersion = "Windows 98"
              End Select
            Case 90
              wvi.dwTextVersion = "Windows Millennium"
            Case Else
              wvi.dwTextVersion = "Windows 95/98/ME"
          End Select
        Case Else
          wvi.dwTextVersion = "Windows 3.1/95/98/ME"
      End Select
    Case Else
      wvi.dwTextVersion = "Unknown Version"
  End Select
  wvi.dwBuildNumber = lpv.dwBuildNumber
  wvi.dwMajorVersion = lpv.dwMajorVersion
  wvi.dwMinorVersion = lpv.dwMinorVersion
  wvi.dwPlatformId = lpv.dwPlatformId
  wvi.szCSDVersion = Trim2(qwe)
  wvi.dwFullVersion = Right$(Str(lpv.dwMajorVersion), Len(Str(lpv.dwMajorVersion)) - 1) & "." & Right$(Str(lpv.dwMinorVersion), Len(Str(lpv.dwMinorVersion)) - 1)
  wvi.dwFullTextV = wvi.dwTextVersion & "   Version " & wvi.dwFullVersion & "   Build " & wvi.dwBuildNumber & "   " & wvi.szCSDVersion
  GetWindowsVersion = wvi
End Function

Public Function GetUserName() As String
  Dim var1 As String, ns As Long
  ns = 255
  var1 = String(ns, 0)
  GetUserNameA var1, ns
  var1 = Left$(var1, ns - 1)
  GetUserName = var1
End Function

Public Function GetComputerName() As String
  Dim var1 As String, ns As Long
  ns = 32
  var1 = String(ns, 0)
  GetComputerNameA var1, ns
  var1 = Left$(var1, ns)
  GetComputerName = var1
End Function

Public Function GetWindowsDir() As String
  Dim StrLen As Long, zPath As String
  zPath = String$(MAX_PATH, 0)
  StrLen = GetWindowsDirectory(zPath, MAX_PATH)
  GetWindowsDir = Left$(zPath, StrLen)
End Function

Public Function GetSystemDir() As String
  Dim StrLen As Long, zPath As String
  zPath = String$(MAX_PATH, 0)
  StrLen = GetSystemDirectory(zPath, MAX_PATH)
  GetSystemDir = Left$(zPath, StrLen)
End Function

Public Function GetTempDir() As String
  Dim StrLen As Long, zPath As String
  zPath = String$(MAX_PATH, 0)
  StrLen = GetTempPath(MAX_PATH, zPath)
  GetTempDir = Left$(zPath, StrLen)
End Function

Public Function GetTempFile() As String
  Dim StrLen As Long, zPath As String, zPath2 As String
  zPath = String$(MAX_PATH, 0)
  zPath2 = String$(MAX_PATH, 0)
  StrLen = GetTempPath(MAX_PATH, zPath)
  zPath = Left$(zPath, StrLen)
  GetTempFileName zPath, "TMP", 0, zPath2
  GetTempFile = Left$(zPath2, InStr(1, zPath2, Chr$(0)) - 1)
End Function

Public Sub ShowAbout(zApp As App, Optional zForm As Form = Nothing)
  Dim qwe As String, texte As String, IconAbout As Long
  If Len(zApp.CompanyName) > 0 Then texte = zApp.CompanyName & " ® " Else texte = ""
  texte = texte & zApp.Title & "   Version " & zApp.Major & "." & zApp.Minor & " (Build " & zApp.Revision & ")"
  If Len(zApp.LegalCopyright) > 0 Then
    texte = texte & vbCrLf & "Copyright © " & zApp.LegalCopyright
  Else
    If Len(zApp.FileDescription) > 0 Then texte = texte & vbCrLf & zApp.FileDescription
  End If
  If Right$(zApp.Path, 1) = "\" Then
    qwe = zApp.Path & zApp.EXEName & ".exe"
  Else
    qwe = zApp.Path & "\" & zApp.EXEName & ".exe"
  End If
  IconAbout = ExtractIcon(zApp.hInstance, qwe, 0)
  If zForm Is Nothing Then
    ShellAbout ByVal 0&, "About " & zApp.Title & "#Windows", texte, IconAbout
  Else
    ShellAbout zForm.HWND, "About " & zApp.Title & "#Windows", texte, IconAbout
  End If
  DestroyIcon IconAbout
End Sub

Public Function GetSpecialFolder(CSIDL As DIR_ID) As String
  Dim zPath As String, r As Long, IDL As ITEMIDLIST
  r = SHGetSpecialFolderLocation(ByVal 0&, CSIDL, IDL)
  If r = 0 Then
    zPath = String$(MAX_PATH, 0)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal zPath)
    GetSpecialFolder = Left$(zPath, InStr(zPath, Chr$(0)) - 1)
  Else
    GetSpecialFolder = ""
  End If
End Function
