VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVerEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'elemento de la lista que tiene la orden de ejecutar

'orden a ejecutar
Private mOrden As String

Public Property Get orden() As String
    orden = mOrden
End Property

Public Property Let orden(val As String)
    mOrden = val
End Property


