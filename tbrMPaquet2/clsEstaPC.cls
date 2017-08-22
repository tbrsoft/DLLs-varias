VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRockola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'manejador de todas las clases
'la idea es que se maneje todo desde aqui y se use el init


Private OMs As clsOrigDiscoManager

Public Sub INIT_MM(idp As Long)
    idEstaPC = idp
    Set OMs = New clsOrigDiscoManager
End Sub

Public Function ORIG() As clsOrigDiscoManager
    Set ORIG = OMs
End Function
