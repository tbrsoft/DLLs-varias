VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDTOCTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CD-ROM TOC entry

Private udtTrack    As TRACK_DATA

Friend Sub SetData( _
    trk As TRACK_DATA _
)

    udtTrack = trk
End Sub

Public Property Get IsAudio() As Boolean
    IsAudio = TrackFormat(udtTrack.ADRCTL) = 0
End Property

Public Property Get StartLBA() As Long
    StartLBA = SHL(udtTrack.addr(0), 24) Or _
               SHL(udtTrack.addr(1), 16) Or _
               SHL(udtTrack.addr(2), 8) Or _
               udtTrack.addr(3)
End Property

Public Property Get TrackNumber() As Byte
    TrackNumber = udtTrack.TrackNumber
End Property

Public Property Get ADR() As Byte
    ADR = HiNibble(udtTrack.ADRCTL)
End Property

Public Property Get CTL() As Byte
    CTL = LoNibble(udtTrack.ADRCTL)
End Property

Private Function TrackFormat( _
    ByVal CTRL As Byte _
) As Integer

    Dim strBin  As String

    ' Lo Nibble in binary
    strBin = Right$(format(Dec2Bin(CTRL), "0000"), 4)
    ' zero digital copy bit
    Mid$(strBin, 3, 1) = 0

    Select Case strBin
        Case "0000", "0001", "1000", "1001", _
             "1100", "1101", "1111":
            TrackFormat = 0
        Case Else
            TrackFormat = 1
    End Select
End Function

Private Function Dec2Bin( _
    ByVal number As Long _
) As String

    Dim x As Integer

    If number >= 2 ^ 32 Then Exit Function

    Do
        If (number And 2 ^ x) Then
            Dec2Bin = "1" & Dec2Bin
        Else
            Dec2Bin = "0" & Dec2Bin
        End If
        x = x + 1
    Loop Until 2 ^ x > number
End Function
