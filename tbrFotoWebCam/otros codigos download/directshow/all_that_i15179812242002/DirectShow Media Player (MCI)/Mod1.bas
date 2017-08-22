Attribute VB_Name = "Mod1"
Option Explicit

Public Enum RegKey   ' lPredefinedKey , hMainKey
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

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

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private rtn As Long

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

Private Function GetMainKeyHandle(MainKeyName As Variant) As Long
  Select Case MainKeyName
    Case "HKEY_CLASSES_ROOT"
      GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
      GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
      GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
      GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
      GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
      GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
      GetMainKeyHandle = HKEY_DYN_DATA
    Case Else
      GetMainKeyHandle = 0
  End Select
End Function

Private Sub ParseKey(KeyName As Variant, Keyhandle As Long, Optional ByVal vBox As Boolean = True)
  Keyhandle = 0
  rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname
  If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
    If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
    Exit Sub 'exit the procedure
  ElseIf rtn = 0 Then 'if the Keyname contains no "\"
    Keyhandle = GetMainKeyHandle(KeyName)
    If Keyhandle = 0 Then
      If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
      Exit Sub
    End If
    KeyName = "" 'leave Keyname blank
  Else 'otherwise, Keyname contains "\"
    Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
    If Keyhandle = 0 Then
      If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
      Exit Sub
    End If
    KeyName = Right(KeyName, Len(KeyName) - rtn)
  End If
End Sub

Public Function SetStringValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String) As Boolean
  Dim hKey As Long, MainKeyHandle As Long
  SetStringValue = False
  ParseKey sKey, MainKeyHandle
  If MainKeyHandle Then
     rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey) 'open the key
     If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
        rtn = RegSetValueEx(hKey, sKeyName, 0, REG_SZ, ByVal KeyValue, Len(KeyValue)) 'write the value
        If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
           MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
        Else
           SetStringValue = True
        End If
        rtn = RegCloseKey(hKey) 'close the key
     Else 'if there was an error opening the key
        MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
     End If
  End If
End Function

Private Function GetErrorMsg(ByVal lErrorCode As Long) As String
  'If an error does accurr, and the user wants error messages displayed, then
  'display one of the following error messages
  Select Case lErrorCode
    Case 1, 1009, 1015
      GetErrorMsg = "The Registry Database is corrupt!"
    Case 2, 6, 1010
      GetErrorMsg = "Bad Key Name"
    Case 3, 1011
      GetErrorMsg = "Can't Open Key"
    Case 4, 1012
      GetErrorMsg = "Can't Read Key"
    Case 1013
      GetErrorMsg = "Can't Write Key"
    Case 5
      GetErrorMsg = "Access to this key is denied"
    Case 8, 14
      GetErrorMsg = "Out of memory"
    Case 7, 87
      GetErrorMsg = "Invalid Parameter"
    Case 234
      GetErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case 259
      GetErrorMsg = "No More Items"
    Case Else
      GetErrorMsg = "Undefined Error Code: " & Str$(lErrorCode)
  End Select
End Function
