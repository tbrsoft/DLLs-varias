VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrSuperGuiObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public SGOName As String
Public Estirable As Boolean

Public AlignementH As Long '0 izquierda 1 centro 2 derecha
Public AlignementV As Long '0 izquierda 1 centro 2 derecha

Private mX As Long
Private mY As Long
Private mW As Long
Private mH As Long
'traducido
Private mX2 As Long
Private mY2 As Long
Private mW2 As Long
Private mH2 As Long

Private mTipo As String 'tipo de objeto del manu
Private mOB As Object 'instancia del objeto del manu

Public Property Get x() As Long:   x = mX: End Property
Public Property Get Y() As Long:   Y = mY: End Property
Public Property Get W() As Long:   W = mW: End Property
Public Property Get H() As Long:   H = mH: End Property

Public Property Get X2() As Long:   X2 = mX2: End Property
Public Property Get Y2() As Long:   Y2 = mY2: End Property
Public Property Get W2() As Long:   W2 = mW2: End Property
Public Property Get H2() As Long:   H2 = mH2: End Property

Public Property Let x(newVal As Long):    mX = newVal:    End Property
Public Property Let Y(newVal As Long):    mY = newVal:    End Property
Public Property Let W(newVal As Long):    mW = newVal:    End Property
Public Property Let H(newVal As Long):    mH = newVal:    End Property

Public Property Let X2(newVal As Long):    mX2 = newVal:    End Property
Public Property Let Y2(newVal As Long):    mY2 = newVal:    End Property
Public Property Let W2(newVal As Long):    mW2 = newVal:    End Property
Public Property Let H2(newVal As Long):    mH2 = newVal:    End Property

Private Sub Class_Initialize()
    'predeterminados
    AlignementH = 1
    AlignementV = 1
End Sub

Public Sub DOIT(sComando As String)
    'le digo al manu que el objeto haga algun comando
    'SEGUIRAQUI
    'ejemplo "text=asa"
End Sub
