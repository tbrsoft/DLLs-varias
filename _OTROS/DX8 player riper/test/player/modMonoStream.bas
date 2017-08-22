Attribute VB_Name = "modMonoStream"
Option Explicit

Public clsStream        As tbrDX8.ISoundStream
Public clsStreams()     As tbrDX8.ISoundStream
Public lngStreamCnt     As Long

' ************************************************
' * stream manager
' ************************************************

Public Sub AddStream(stream As ISoundStream)
    ReDim Preserve clsStreams(lngStreamCnt) As ISoundStream
    Set clsStreams(lngStreamCnt) = stream
    lngStreamCnt = lngStreamCnt + 1
End Sub

Public Function StreamFromExt( _
    ByVal ext As String _
) As ISoundStream

    Dim i       As Long

    ext = Right$(ext, 3)

    For i = 0 To lngStreamCnt - 1
        If InStr(1, Join(clsStreams(i).Extensions, ";"), ext, vbTextCompare) Then
            Set StreamFromExt = clsStreams(i)
            Exit Function
        End If
    Next
End Function

Public Function GetAllExtensions() As String()
    Dim i               As Long
    Dim j               As Long
    Dim strExt          As String
    Dim strExts()       As String
    Dim strCurrExts()   As String

    For i = 0 To lngStreamCnt - 1
        strCurrExts = clsStreams(i).Extensions

        ' both StreamMP3 and StreamWMA support MP3,
        ' so eliminate double entries
        For j = 0 To UBound(strCurrExts)
            If InStr(strExt, strCurrExts(j)) <= 0 Then
                strExt = strExt & strCurrExts(j) & ";"
            End If
        Next
    Next

    If Right$(strExt, 1) = ";" Then
        strExt = Left$(strExt, Len(strExt) - 1)
    End If

    strExts = Split(strExt, ";")

    GetAllExtensions = strExts
End Function
