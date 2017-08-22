VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICDAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function InitMe() As STREAM_ERR
End Function

Public Function DeInitMe() As STREAM_ERR
End Function

Public Function SelectDevice( _
    ByVal index As Long _
) As STREAM_ERR

End Function

Public Function DeviceChar( _
    ByVal index As Long _
) As String

End Function

Public Function DeviceName( _
    ByVal index As Long _
) As String

End Function

Public Function DeviceCount() As Long
End Function

Public Function GetTOC( _
    clsTOC As CDTOC _
) As STREAM_ERR

End Function

Public Function ReadSectorsRAW( _
    ByVal LBA As Long, _
    ByVal sectors As Long, _
    ByVal ptr As Long _
) As STREAM_ERR

End Function

Public Function OpenCD() As STREAM_ERR
End Function

Public Function CloseCD() As STREAM_ERR
End Function

Public Function IsDiscPresent() As Boolean
End Function
