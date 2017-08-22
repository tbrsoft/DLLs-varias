VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsVERSION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32S = 0
Const PLANES = 14
Const BITSPIXEL = 12

Public Enum vWindows
    NoCalculadoAun = -2
    Error = -1
    Win31 = 0
    Win95 = 1
    Win95OSR2 = 2
    Win98 = 3
    Win98SE = 4
    WinME = 5
    WinNT4 = 6
    Win2000 = 7
    WinXp = 8
    WinXP2 = 9 'despeus del XP
End Enum

Private ToNoCalculateAgain As vWindows
Private mVersionString As String

Public Property Get Version() As vWindows
    Version = ToNoCalculateAgain
End Property

Public Function GetVersion() As vWindows
    
    On Error GoTo errh
    'on the event of an erro jump to label errh
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    
    'initialize variable
    'if the function returned succesfully then...
     If GetVersionEx(osvi) <> 0 Then
     
        Select Case osvi.dwPlatformId
        
            Case VER_PLATFORM_WIN32S:
                ToNoCalculateAgain = Win31
                mVersionString = "Windows 3.1"
                
            Case VER_PLATFORM_WIN32_WINDOWS
                If osvi.dwMajorVersion = 4 Then
                '95''''''''''''''''''''''''''''''
                    If osvi.dwMinorVersion = 0 Then
                        ToNoCalculateAgain = Win95
                        mVersionString = "Windows 95"
                        ' es el OSR2
                        If LCase(osvi.szCSDVersion) = "c" Or LCase(osvi.szCSDVersion) = "b" Then
                            ToNoCalculateAgain = Win95OSR2
                            mVersionString = "Windows 95 OSR2"
                        End If
                    End If
                '98''''''''''''''''''''''''''''''
                    If osvi.dwMinorVersion = 10 Then
                        ToNoCalculateAgain = Win98
                        mVersionString = "Windows 98 1º edicion"
                        'o es el SE
                        If LCase(osvi.szCSDVersion) = "a" Then
                            ToNoCalculateAgain = Win98SE
                            mVersionString = "Windows 98 2º edicion"
                        End If
                    End If
                'Me''''''''''''''''''''''''''''''
                    If osvi.dwMinorVersion = 90 Then
                        ToNoCalculateAgain = WinME
                        mVersionString = "Windows Millenium"
                    End If
                    
                End If
            Case VER_PLATFORM_WIN32_NT:
            'NT4''''''''''''''''''''''''''''''
            If osvi.dwMajorVersion = 4 Then
                ToNoCalculateAgain = WinNT4
                mVersionString = "Windows NT4"
            End If
            If osvi.dwMajorVersion = 5 Then
            '2000''''''''''''''''''''''''''''''
                If osvi.dwMinorVersion = 0 Then
                    ToNoCalculateAgain = Win2000
                    mVersionString = "Windows 2000"
                End If
            'XP''''''''''''''''''''''''''''''
                If osvi.dwMinorVersion >= 1 Then
                    ToNoCalculateAgain = WinXp
                    mVersionString = "Windows XP"
                End If
            End If
            
            'XP futuros y otros sistemas'''''
            If osvi.dwMajorVersion > 5 Then
                ToNoCalculateAgain = WinXP2
                mVersionString = "Windows XP2"
            End If
            
        End Select
    End If
    GetVersion = ToNoCalculateAgain
    Exit Function
errh:
    ToNoCalculateAgain = Error
    GetVersion = ToNoCalculateAgain
End Function

Private Sub Class_Initialize()
    ToNoCalculateAgain = NoCalculadoAun
End Sub

Public Function GetVersionString() As String
    GetVersionString = mVersionString
End Function
