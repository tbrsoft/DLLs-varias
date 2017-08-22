VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMMType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada tipo de contenido multimedia tiene algunas especificaciones
'como las extenciones de archivos que pueden ser y otros detalles
'la forma de llamarlo, la descripcion, la ayuda para reproducirlo o venderlo, etc

Private mNombre As String
Private mDescripcion As String

Private mEXTs() As String 'extenciones que puede presentar un archivo con esta multimedia
Private mCanPlay As Long 'indica si se puede reproducir (se refiere a hacerlo por dinero)
Private mCanExpen As Long 'indica si se puede vender con el expendedor
Private mCanPreview As Long 'indica si se puede hacer una previsualizacion al pasar por el en una lista (mp3 puede solo si esta configurado, las imagenes, los juegos java si pueden. por ejemplo los videos 3GP no pueden)

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(newV As String)
    mNombre = newV
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(newV As String)
    mDescripcion = newV
End Property

Public Property Let CanPlay(nVal As Long)
    mCanPlay = nVal
End Property

Public Property Let CanExpen(nVal As Long)
    mCanExpen = nVal
End Property

Public Property Let CanPreview(nVal As Long)
    mCanPreview = nVal
End Property

Public Property Get CanPlay() As Long
    CanPlay = mCanPlay
End Property

Public Property Get CanExpen() As Long
    CanExpen = mCanExpen
End Property

Public Property Get CanPreview() As Long
    CanPreview = mCanPreview
End Property

Public Sub Extension_Clean()
    ReDim mEXTs(0)
End Sub

Public Sub Extencion_Add(sVal As String)
    Dim j As Long
    j = UBound(mEXTs) + 1
    ReDim Preserve mEXTs(j)
    mEXTs(j) = sVal
End Sub

Private Sub Class_Initialize()
    Extension_Clean
End Sub
