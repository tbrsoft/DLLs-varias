Attribute VB_Name = "extrastuff"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength _
    As Long, ByVal hwndCallback As Long) As Long


Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal _
    fdwError As Long, ByVal lpszErrorText As String, ByVal cchErrorText As Long) As Long
  
  
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SW_SHOWMAXIMIZED = 3

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Sub Openurl(url As String)
ShellExecute frmMain.hwnd, "open", url, vbNullString, vbNullString, 3
End Sub
