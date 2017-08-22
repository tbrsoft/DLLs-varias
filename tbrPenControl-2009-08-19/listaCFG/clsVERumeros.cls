VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVERumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'opciones que se elijen con numeros (solo enteros)

Private mMinVal As Long
Private mMaxVal As Long
Private mStep As Long 'paso de cada cambio
Private mPredeterminado As Long

Private mValTMP As Long
Private mVal As Long

Public Property Get MinVal() As Long
    MinVal = mMinVal
End Property

Public Property Get MaxVal() As Long
    MaxVal = mMaxVal
End Property

Public Property Get Step() As Long
    Step = mStep
End Property

Public Property Let MinVal(nVal As Long)
    mMinVal = nVal
End Property

Public Property Let MaxVal(nVal As Long)
    mMaxVal = nVal
End Property

Public Property Let Step(nVal As Long)
    mStep = nVal
End Property

Public Property Get ValTMP() As Long
    ValTMP = mValTMP
End Property

Public Property Let ValTMP(nVal As Long)
    mValTMP = nVal
End Property

Public Property Get Valor() As Long
    Valor = mVal
End Property

Public Property Let Valor(nVal As Long)
    mVal = nVal
    mValTMP = nVal 'temporal tambien es el actual
End Property

Public Property Get Predeterminado() As Long
    Predeterminado = mPredeterminado
End Property

Public Property Let Predeterminado(nVal As Long)
    mPredeterminado = nVal
End Property

Public Sub ConfirmOption()
    mVal = mValTMP
End Sub

Private Sub Class_Initialize()
    mVal = -9999 'es como el "NULL" de texto simple o lista combo
End Sub
