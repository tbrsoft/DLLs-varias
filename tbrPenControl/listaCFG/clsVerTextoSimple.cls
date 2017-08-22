VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVerTextoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'opciones que se elijen de un texto y algun valor predeterminado

Private mPredeterminado As String 'valor por defecto
Private mMultiline As Long 'indica si el texto puede tener varios renglones
Private mVal As String 'valor actual
Private mValTMP As String 'valor temporal mientras se escribe la caja de texto. Es el que se asigna a val cuando corresponda

'cuando esto sea path necesitara adicionales
Private mShowFile As Long 'mostrar dialogo de ARCHIVOS
Private mShowFolder As Long 'mostrar dialogo de CARPETAS
'si alguno de los dos es verdadero es para path
Private mFilter As String 'filtro para el showfile
Private mInitialPath As String 'path al abrir el dialogo

Public Property Get Multiline() As Long
    Multiline = mMultiline
End Property

Public Property Let Multiline(val As Long)
    mMultiline = val
End Property

Public Property Get Predeterminado() As String
    Predeterminado = mPredeterminado
End Property

Public Property Let Predeterminado(nVal As String)
    mPredeterminado = nVal
End Property

Public Property Get Valor() As String
    Valor = mVal
End Property

Public Property Let Valor(nVal As String)
    mVal = nVal
    mValTMP = nVal 'temporal tambien es el actual
End Property

Public Property Get ValorTMP() As String
    ValorTMP = mValTMP
End Property

Public Property Let ValorTMP(nVal As String)
    mValTMP = nVal
End Property

Public Sub ConfirmOption()
    mVal = mValTMP
End Sub

Private Sub Class_Initialize()
    mVal = "NULL" 'para indicar que nunca se acepto un valor
    mInitialPath = App.path
    mShowFile = 0
    mShowFolder = 0
End Sub

Public Property Get ShowFile() As Long
    ShowFile = mShowFile
End Property

Public Property Let ShowFile(val As Long)
    mShowFile = val
    If mShowFile > 0 Then
        mShowFolder = 0 'si pongo uno sale el otro
        mMultiline = 1 'si o si multirenglon para leer el path ok
    End If
End Property

Public Property Get ShowFolder() As Long
    ShowFolder = mShowFolder
End Property

Public Property Let ShowFolder(val As Long)
    mShowFolder = val
    If mShowFolder > 0 Then
        mShowFile = 0 'si pongo uno sale el otro
        mMultiline = 1 'si o si multirenglon para leer el path ok
    End If
End Property

Public Property Get InitialPath() As String
    InitialPath = mInitialPath
End Property

Public Property Let InitialPath(nVal As String)
    mInitialPath = nVal
End Property

Public Property Get Filter() As String
    Filter = mFilter
End Property

Public Property Let Filter(nVal As String)
    mFilter = nVal
End Property

