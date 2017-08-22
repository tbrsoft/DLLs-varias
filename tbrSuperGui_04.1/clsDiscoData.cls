VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiscoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================
Dim qTitulo As String
Dim qNumero As Long
Dim qPathImagen As String
Dim qPathDisco As String
Dim qTagMisterioso As String
'=========================================


Public Property Get Titulo() As String
    Titulo = qTitulo
End Property
Public Property Let Titulo(ByVal NV As String)
    qTitulo = NV
End Property
Public Property Get Numero() As Long
    Numero = qNumero
End Property
Public Property Let Numero(ByVal NV As Long)
    qNumero = NV
End Property
Public Property Get PathImagen() As String
    PathImagen = qPathImagen
End Property
Public Property Let PathImagen(ByVal NV As String)
    qPathImagen = NV
    'CargarImagen qPathImagen
End Property
Public Property Get PathDisco() As String
    PathDisco = qPathDisco
End Property
Public Property Let PathDisco(ByVal NV As String)
    qPathDisco = NV
End Property
Public Property Get TagMisterioso() As String
    TagMisterioso = qTagMisterioso
End Property
Public Property Let TagMisterioso(ByVal NV As String)
    qTagMisterioso = NV
End Property

