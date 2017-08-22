VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDTOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CD-ROM TOC (Table Of Contents)
' Track Collection

Private clsTracks() As CDTOCTrack

Public Property Get TrackCount() As Long
    On Error Resume Next
    TrackCount = UBound(clsTracks)
End Property

Public Property Let TrackCount( _
    ByVal c As Long _
)

    Dim i   As Long

    ReDim clsTracks(1 To c) As CDTOCTrack
    For i = 1 To c
        Set clsTracks(i) = New CDTOCTrack
    Next
End Property

Friend Sub SetTrackData( _
    trk As TRACK_DATA _
)

    If trk.TrackNumber = &HAA Then  ' LeadOut
        clsTracks(TrackCount).SetData trk
    Else
        clsTracks(trk.TrackNumber).SetData trk
    End If
End Sub

Public Property Get track( _
    ByVal NumTrack As Long _
) As CDTOCTrack

    Set track = clsTracks(NumTrack)
End Property
