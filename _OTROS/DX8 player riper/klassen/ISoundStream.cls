VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISoundStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum STREAM_ERR
    STREAM_OK
    STREAM_INVALID_SOURCE
    STREAM_READ_ERROR
    STREAM_OUT_OF_RANGE
    STREAM_NOT_SUPPORTED
    STREAM_END_OF_STREAM
    STREAM_TOO_BIG
    STREAM_CANCELED
    STREAM_NEED_MORE
    STREAM_INVALID_DEVICE
End Enum

Public Enum SEEK_MODE
    SEEK_PERCENT
    SEEK_SECONDS
End Enum

Public Property Get Info() As ISoundInfo
End Property

Public Property Get Extensions() As String()
End Property

Public Property Get EndOfStream() As Boolean
End Property

Public Property Get Seekable() As Boolean
End Property

Public Property Get WFXPointer() As Long
End Property

'me indica si el archivo sirve
Public Function OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

End Function

Public Function CloseSource() As STREAM_ERR
End Function

Public Function Read(ByVal pBuffer As Long, ByVal cBytes As Long, cBytesRead As Long) As STREAM_ERR

End Function

Public Function SeekTo(ByVal value As Long, mode As SEEK_MODE) As STREAM_ERR

End Function

Private Sub Class_Initialize()

End Sub
