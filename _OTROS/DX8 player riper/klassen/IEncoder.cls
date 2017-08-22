VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let Title( _
    ByVal strTitle As String _
)

End Property

Public Property Get Title() As String
End Property

Public Property Let Album( _
    ByVal strAlbum As String _
)

End Property

Public Property Get Album() As String
End Property

Public Property Let Artist( _
    ByVal strArtist As String _
)

End Property

Public Property Get Artist() As String
End Property

Public Property Get Extension() As String
End Property

Public Function DeInit() As STREAM_ERR
End Function

Public Function Init( _
    ByVal dstFilename As String, _
    ByVal srcSampleRate As Long, _
    ByVal srcChannels As Integer, _
    ByVal srcBitsPerSample As Integer, _
    cbPCMData As Long _
) As STREAM_ERR

End Function

Public Function Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

End Function
