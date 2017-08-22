VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTemaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
Dim qTitulo As String
Dim qNumero As Long
Dim qDuracion As String
Dim qBloqueado As Boolean
Dim qTagMisterioso As String
'Grafico----------------------------------
'Dim qX As Long
'Dim qY As Long
'Dim qAncho As Long
'Dim qAlto As Long
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
Public Property Get Duracion() As String
    Duracion = qDuracion
End Property
Public Property Let Duracion(ByVal NV As String)
    qDuracion = NV
End Property
Public Property Get TagMisterioso() As String
    TagMisterioso = qTagMisterioso
End Property
Public Property Let TagMisterioso(ByVal NV As String)
    qTagMisterioso = NV
End Property
Public Property Get Bloqueado() As Boolean
    Bloqueado = qBloqueado
End Property
Public Property Let Bloqueado(ByVal NV As Boolean)
    qBloqueado = NV
End Property

'GRAFICO-----------------------------------------------

'Public Property Get x() As Long
'    x = qX
'End Property
'Public Property Let x(ByVal NV As Long)
'    qX = NV
'End Property
'Public Property Get y() As Long
'    y = qY
'End Property
'Public Property Let y(ByVal NV As Long)
'    qY = NV
'End Property
'Public Property Get Ancho() As Long
'    Ancho = qAncho
'End Property
'Public Property Let Ancho(ByVal NV As Long)
'    qAncho = NV
'End Property
'Public Property Get Alto() As Long
'    Alto = qAlto
'End Property
'Public Property Let Alto(ByVal NV As Long)
'    qAlto = NV
'End Property
'
