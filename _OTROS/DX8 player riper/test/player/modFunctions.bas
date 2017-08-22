Attribute VB_Name = "modFunctions"
Option Explicit

Public Declare Function GetCurrentThread Lib "kernel32" ( _
) As Long

Public Declare Function SetThreadPriority Lib "kernel32" ( _
    ByVal hThread As Long, _
    ByVal nPriority As Long _
) As Long

Public Function FmtTime( _
    ByVal seconds As Long _
) As String

    Dim minutes As Long

    minutes = seconds \ 60
    seconds = seconds Mod 60

    FmtTime = Format(minutes, "00") & ":" & Format(seconds, "00")
End Function

Public Function AddSlash( _
    ByVal strPath As String _
) As String

    AddSlash = IIf(Right$(strPath, 1) = "\", strPath, strPath & "\")
End Function

Public Function GetExtension( _
    ByVal strFile As String _
) As String

    If InStr(strFile, ".") Then
        GetExtension = Mid$(strFile, InStrRev(strFile, ".") + 1)
    Else
        GetExtension = strFile
    End If
End Function

Public Function GetFilename( _
    ByVal strPath As String _
) As String

    GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function

Public Function FileExists( _
    strPath As String _
) As Boolean

    FileExists = (GetAttr(strPath) And (vbDirectory Or vbVolume)) = 0
End Function

Public Function Max( _
    ByVal val1 As Long, _
    ByVal val2 As Long _
) As Long

    Max = IIf(val1 > val2, val1, val2)
End Function

Public Function Min( _
    ByVal val1 As Long, _
    ByVal val2 As Long _
) As Long

    Min = IIf(val1 < val2, val1, val2)
End Function

Public Sub ProcessThreadPrioritySet( _
    ByVal Priority As Long _
)

    SetThreadPriority GetCurrentThread, Priority
End Sub
