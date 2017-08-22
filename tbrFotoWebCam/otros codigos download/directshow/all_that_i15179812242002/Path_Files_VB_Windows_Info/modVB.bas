Attribute VB_Name = "modVB"
Option Explicit

Public Enum VirtualKey
  VK_LBUTTON = &H1
  VK_RBUTTON = &H2
  VK_CTRLBREAK = &H3
  VK_MBUTTON = &H4
  VK_BACKSPACE = &H8
  VK_TAB = &H9
  VK_ENTER = &HD
  VK_SHIFT = &H10
  VK_CONTROL = &H11
  VK_ALT = &H12
  VK_PAUSE = &H13
  VK_CAPSLOCK = &H14
  VK_ESCAPE = &H1B
  VK_SPACE = &H20
  VK_PAGEUP = &H21
  VK_PAGEDOWN = &H22
  VK_END = &H23
  VK_HOME = &H24
  VK_LEFT = &H25
  VK_UP = &H26
  VK_RIGHT = &H27
  VK_DOWN = &H28
  VK_PRINTSCREEN = &H2C
  VK_INSERT = &H2D
  VK_DELETE = &H2E
  VK_0 = &H30
  VK_1 = &H31
  VK_2 = &H32
  VK_3 = &H33
  VK_4 = &H34
  VK_5 = &H35
  VK_6 = &H36
  VK_7 = &H37
  VK_8 = &H38
  VK_9 = &H39
  VK_A = &H41
  VK_B = &H42
  VK_C = &H43
  VK_D = &H44
  VK_E = &H45
  VK_F = &H46
  VK_G = &H47
  VK_H = &H48
  VK_I = &H49
  VK_J = &H4A
  VK_K = &H4B
  VK_L = &H4C
  VK_M = &H4D
  VK_N = &H4E
  VK_O = &H4F
  VK_P = &H50
  VK_Q = &H51
  VK_R = &H52
  VK_S = &H53
  VK_T = &H54
  VK_U = &H55
  VK_V = &H56
  VK_W = &H57
  VK_X = &H58
  VK_Y = &H59
  VK_Z = &H5A
  VK_LWINDOWS = &H5B
  VK_RWINDOWS = &H5C
  VK_APPSPOPUP = &H5D
  VK_NUMPAD0 = &H60
  VK_NUMPAD1 = &H61
  VK_NUMPAD2 = &H62
  VK_NUMPAD3 = &H63
  VK_NUMPAD4 = &H64
  VK_NUMPAD5 = &H65
  VK_NUMPAD6 = &H66
  VK_NUMPAD7 = &H67
  VK_NUMPAD8 = &H68
  VK_NUMPAD9 = &H69
  VK_MULTIPLY = &H6A
  VK_ADD = &H6B
  VK_SUBTRACT = &H6D
  VK_DECIMAL = &H6E
  VK_DIVIDE = &H6F
  VK_F1 = &H70
  VK_F2 = &H71
  VK_F3 = &H72
  VK_F4 = &H73
  VK_F5 = &H74
  VK_F6 = &H75
  VK_F7 = &H76
  VK_F8 = &H77
  VK_F9 = &H78
  VK_F10 = &H79
  VK_F11 = &H7A
  VK_F12 = &H7B
  VK_NUMLOCK = &H90
  VK_SCROLL = &H91
  VK_LSHIFT = &HA0
  VK_RSHIFT = &HA1
  VK_LCONTROL = &HA2
  VK_RCONTROL = &HA3
  VK_LALT = &HA4
  VK_RALT = &HA5
  VK_POINTVIRGULE = &HBA
  VK_ADD_EQUAL = &HBB
  VK_VIRGULE = &HBC
  VK_MINUS_UNDERLINE = &HBD
  VK_POINT = &HBE
  VK_SLASH = &HBF
  VK_TILDE = &HC0
  VK_LEFTBRACKET = &HDB
  VK_BACKSLASH = &HDC
  VK_RIGHTBRACKET = &HDD
  VK_QUOTE = &HDE
  VK_APOSTROPHE = &HDE
End Enum

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Public Type MemoryStatus
    MemoryLoad As Long
    MemoryLoad2 As Single
    TotalPhys As Currency
    AvailPhys As Currency
    TotalVirtual As Currency
    AvailVirtual As Currency
    TotalPageFile As Currency
    AvailPageFile As Currency
    AvailExtendedVirtual As Currency
End Type

Public Enum PRIORITY_CLASS
  REALTIME_PRIORITY = &H100
  HIGH_PRIORITY = &H80
  NORMAL_PRIORITY = &H20
  IDLE_PRIORITY = &H40
End Enum

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As VbAppWinStyle
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function IsCharAlphaZ Lib "user32.dll" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumericZ Lib "user32.dll" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLowerZ Lib "user32.dll" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpperZ Lib "user32.dll" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare Sub Sleepy Lib "kernel32.dll" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function FlashWindow Lib "user32.dll" (ByVal HWND As Long, ByVal bInvert As Long) As Long
Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal StringToExec As Long, ByVal Any1 As Long, ByVal Any2 As Long, ByVal CheckOnly As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Private Declare Function WinExec Lib "kernel32.dll" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Function isKeyDown(ByVal zKey As VirtualKey) As Boolean
  Dim var1 As Long
  If (zKey = VK_CAPSLOCK) Or (zKey = VK_NUMLOCK) Or (zKey = VK_SCROLL) Then
    var1 = &H1
  Else
    var1 = &H80
  End If
  If (GetKeyState(zKey) And var1) = var1 Then
    isKeyDown = True
  Else
    isKeyDown = False
  End If
End Function

Public Function isAnyKeyDown(Optional ByVal IgnoreLocksKeys As Boolean = False) As Boolean
  Dim t As Integer, KD As Boolean
  Dim keystat(0 To 255) As Byte
  GetKeyboardState keystat(0)
  KD = False
  If IgnoreLocksKeys = False Then
    For t = 0 To 255
      If (keystat(t) And &H80) = &H80 Then
        KD = True
        Exit For
      End If
    Next t
  Else
    For t = 0 To 255
      If ((keystat(t) And &H80) = &H80) And (t <> VK_CAPSLOCK) And (t <> VK_NUMLOCK) And (t <> VK_SCROLL) Then
        KD = True
        Exit For
      End If
    Next t
  End If
  isAnyKeyDown = KD
End Function

Public Function IsCharAlpha(ByVal cChar As Byte) As Boolean
  IsCharAlpha = IsCharAlphaZ(ByVal cChar)
End Function

Public Function IsCharAlphaNumeric(ByVal cChar As Byte) As Boolean
  IsCharAlphaNumeric = IsCharAlphaNumericZ(ByVal cChar)
End Function

Public Function IsCharNumeric(ByVal cChar As Byte) As Boolean
  IsCharNumeric = IsCharAlphaNumericZ(ByVal cChar) And (Not IsCharAlphaZ(ByVal cChar))
End Function

Public Function IsCharLower(ByVal cChar As Byte) As Boolean
  IsCharLower = IsCharLowerZ(ByVal cChar)
End Function

Public Function IsCharUpper(ByVal cChar As Byte) As Boolean
  IsCharUpper = IsCharUpperZ(ByVal cChar)
End Function

Public Function IsStringNumeric(ByVal q As String, Optional ByVal WithNegative As Boolean = True, Optional ByVal WithDecimal As Boolean = True) As Boolean
  Dim t As Long, zPoint As Boolean
  If Len(q) > 0 Then
    For t = 1 To Len(q)
      If Not IsCharNumeric(Asc(Mid$(q, t, 1))) Then
        Select Case Mid$(q, t, 1)
          Case "-"
            If t <> 1 Or WithNegative = False Or Len(q) < 2 Then
              IsStringNumeric = False
              Exit Function
            End If
          Case "."
            If zPoint = True Or WithDecimal = False Or Len(q) < 2 Then
              IsStringNumeric = False
              Exit Function
            Else
              zPoint = True
            End If
          Case Else
            IsStringNumeric = False
            Exit Function
        End Select
      End If
    Next t
    If Left$(q, 1) = "-" And zPoint = True And Len(q) = 2 Then
      IsStringNumeric = False
    Else
      IsStringNumeric = True
    End If
  Else
    IsStringNumeric = False
  End If
End Function

Public Sub Sleep(ByVal dwMilliseconds As Long)
  Dim zz As Single
  zz = Timer
  Do
    DoEvents
    If Timer >= zz Then
      If (Timer - zz) >= (dwMilliseconds / 1000) Then Exit Do
    Else
      If ((86400 - zz) + Timer) >= (dwMilliseconds / 1000) Then Exit Do
    End If
  Loop
End Sub

Public Sub Sleep2(ByVal dwMilliseconds As Long)
  Sleepy ByVal dwMilliseconds
End Sub

Public Sub Flash(zForm As Form)
  FlashWindow zForm.HWND, 1
End Sub

Public Function GetAbout(zApp As App) As String
  Dim qwe As String
  If Len(zApp.CompanyName) > 0 Then qwe = zApp.CompanyName & " ® " Else qwe = ""
  qwe = qwe & zApp.Title & vbCrLf & "Version " & zApp.Major & "." & zApp.Minor & "." & zApp.Revision
  If Len(zApp.LegalCopyright) > 0 Then qwe = qwe & vbCrLf & "Copyright © " & zApp.LegalCopyright
  If Len(zApp.FileDescription) > 0 Then qwe = qwe & vbCrLf & zApp.FileDescription
  GetAbout = qwe
End Function

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
  CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
  LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

Public Function GetMemory() As MemoryStatus
  Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
  Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
  Const Language_Neutral As Long = &H0
  Const User_Default_Language As Long = &H400
  Const System_Default_Language As Long = &H800
  Dim MemStat As MEMORYSTATUSEX
  Dim MemStat2 As MemoryStatus
  Dim ErrorMessage As String
  MemStat.dwLength = Len(MemStat)
  If GlobalMemoryStatusEx(MemStat) <> 0 Then
    MemStat2.AvailExtendedVirtual = LargeIntToCurrency(MemStat.ullAvailExtendedVirtual)
    MemStat2.AvailPageFile = LargeIntToCurrency(MemStat.ullAvailPageFile)
    MemStat2.AvailPhys = LargeIntToCurrency(MemStat.ullAvailPhys)
    MemStat2.AvailVirtual = LargeIntToCurrency(MemStat.ullAvailVirtual)
    MemStat2.TotalPageFile = LargeIntToCurrency(MemStat.ullTotalPageFile)
    MemStat2.TotalPhys = LargeIntToCurrency(MemStat.ullTotalPhys)
    MemStat2.TotalVirtual = LargeIntToCurrency(MemStat.ullTotalVirtual)
    MemStat2.MemoryLoad = MemStat.dwMemoryLoad
    MemStat2.MemoryLoad2 = ((MemStat2.TotalPhys - MemStat2.AvailPhys) / MemStat2.TotalPhys) * 100
  Else
    ErrorMessage = Space$(500)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, GetLastError, User_Default_Language, ErrorMessage, Len(ErrorMessage), 0
    MsgBox Trim2(ErrorMessage), vbExclamation, "Super"
  End If
  GetMemory = MemStat2
End Function

Public Function vbExecute(ByVal var1 As String, Optional ByVal ShowError As Boolean = False) As Long
  On Local Error GoTo ErrHnd
  Dim var2 As Long
  var2 = EbExecuteLine(StrPtr(var1), 0&, 0&, 1)
  If var2 = 0 Then
    EbExecuteLine StrPtr(var1), 0&, 0&, 0&
  Else
    If ShowError Then Error var2
  End If
  vbExecute = var2
  Exit Function
ErrHnd:
  MsgBox "Error # " & Err.Number & " : " & Err.Description, vbExclamation, "Super", Err.HelpFile, Err.HelpContext
  vbExecute = var2
End Function

Public Sub End2(ByVal uExitCode As Long)
  ExitProcess uExitCode
End Sub

Public Function Exec(ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
  Dim var1 As Long
  var1 = WinExec(CmdLine, WindowStyle)
  If var1 > 31 Then
    Exec = True
  Else
    Exec = False
  End If
End Function

Public Function Exec2(ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Boolean
  Dim sinfo As STARTUPINFO, pinfo As PROCESS_INFORMATION
  sinfo.cb = Len(sinfo)
  sinfo.dwFlags = &H1
  sinfo.wShowWindow = WindowStyle
  If CreateProcess(vbNullString, CmdLine, ByVal 0&, ByVal 0&, 1&, pclass, ByVal 0&, vbNullString, sinfo, pinfo) <> 0 Then
    CloseHandle pinfo.hThread
    CloseHandle pinfo.hProcess
    Exec2 = True
  Else
    Exec2 = False
  End If
End Function

Public Function GetExitCode(ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal zWait As Boolean = False, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Variant
  Const Infinite As Long = &HFFFFFFFF
  Const STILL_ACTIVE As Long = &H103
  Dim sinfo As STARTUPINFO, pinfo As PROCESS_INFORMATION, ExitCode As Long
  sinfo.cb = Len(sinfo)
  sinfo.dwFlags = &H1
  sinfo.wShowWindow = WindowStyle
  If CreateProcess(vbNullString, CmdLine, ByVal 0&, ByVal 0&, 1&, pclass, ByVal 0&, vbNullString, sinfo, pinfo) <> 0 Then
    If zWait = True Then WaitForSingleObject pinfo.hProcess, Infinite
    Do
      GetExitCodeProcess pinfo.hProcess, ExitCode
      DoEvents
    Loop While ExitCode = STILL_ACTIVE
    CloseHandle pinfo.hThread
    CloseHandle pinfo.hProcess
    GetExitCode = ExitCode
  Else
    GetExitCode = "ERROR"
  End If
End Function
