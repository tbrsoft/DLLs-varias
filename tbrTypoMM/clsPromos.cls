VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPromos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada una de las promociones

'una promocion es un tipo de contenido en una x cantidad por x creditos
Private mCantidad As Long
Private mCreditos As Long

Public Property Get Cantidad() As Long
    Cantidad = mCantidad
End Property

Public Property Let Cantidad(newV As Long)
    mCantidad = newV
End Property

Public Property Get Creditos() As Long
    Creditos = mCreditos
End Property

Public Property Let Creditos(newV As Long)
    mCreditos = newV
End Property


