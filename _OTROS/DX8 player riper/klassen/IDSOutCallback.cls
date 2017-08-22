VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSOutCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub OnEndOfStream( _
    sndout As SoundOut _
)

End Sub

Public Sub OnStatusChanged( _
    sndout As SoundOut, _
    ByVal udeStat As DS_PlayState _
)

End Sub

Public Sub Samples( _
    sndout As SoundOut, _
    intSamples() As Integer, _
    ByVal datalength As Long, _
    ByVal channels As Integer _
)

End Sub
