VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada uno de los elementos de la lista

Private mID As Long
Private mPadre As Long 'id del padre

Private mCaption As String
Private mHelp As String 'ayuda sobre la opcion

Public Enum ElemType 'cada elemento al ser elegido mostrara un contenido diferente
    ListaCombo = 0 'lista en la que se elije 1 sola opcion
    ListaMulti = 1 'lista con checkboxes
    Numero = 2 'se elije un numero (habra propiedades minimo, maximo y step
    TextoSimple = 3 'el usuario escribe lo que quiere
    EjecutarProceso = 4 'por ejemplo se carga musica
    SoloInfo = 5 'es para los tiutulos de las secciones, al elegir entrar se despliega su menu interno
    SelectPath = 6
End Enum

Private mType As ElemType

'posibles componentes internos de este.
Public Internal_ListaSImple As New clsVERListaSimple
Public Internal_Numeros As New clsVERumeros
Public Internal_VerEXE As New clsVerEXE
Public Internal_TextoSimple As New clsVerTextoSimple

Private mNodeOP As Node 'si se va a mostrar todo en un treeview cada elemento es un nodo y me queda comodo acceder a el desde aqui

Private mPlusInfo As String 'informacion adicional, help es algo mas fijo y generico, esto es informacion adicional / dinamica segun el caso
Private mMarca As String 'marca para hacer colecciones de elementos que son grupos especificos, por ejemplo los origenes que tiene el pendrive
Private mInternalReal As String 'valor feo que en realidad se usa al elegir este tag. Lo uso para guardar path largos y mostrar solo el baseName

Private mValorElegido As String 'sea del tipo que sea hay un valor elegido si considero cada elemento como una configuracion
'sirve para tener a mano un resumen final y pasarlo a cualquier sistema de configuración


'///////////////////////////////////////////////////////////////

Public Property Get ValorElegido() As String
    
    If mType = EjecutarProceso Then ValorElegido = Internal_VerEXE.orden
    If mType = ListaCombo Then ValorElegido = Internal_ListaSImple.GetSelectOpInternal 'valor interno, no el visible
    If mType = ListaMulti Then ValorElegido = "NULL" '"NOT IMPLEMENTED"
    If mType = Numero Then ValorElegido = CStr(Internal_Numeros.Valor)
    If mType = SoloInfo Then ValorElegido = "NULL" '"NOT USED"
    If mType = TextoSimple Then ValorElegido = Internal_TextoSimple.Valor
    If mType = SelectPath Then ValorElegido = Internal_TextoSimple.Valor
    
End Property

Public Property Get id() As Long
    id = mID
End Property

Public Property Let id(newID As Long)
    mID = newID
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(newCaption As String)
    mCaption = newCaption
End Property

Public Property Get Help() As String
    Help = mHelp
End Property

Public Property Let Help(newHelp As String)
    mHelp = newHelp
End Property

Public Property Get eType() As ElemType
    eType = mType
End Property

Public Property Let eType(newType As ElemType)
    mType = newType
End Property

Public Property Let Padre(IdPadre As Long)
    mPadre = IdPadre
End Property

Public Property Get Padre() As Long
    Padre = mPadre
End Property

Public Property Set NodeOp(newNode As Node)
    Set mNodeOP = newNode
End Property

Public Property Get NodeOp() As Node
    Set NodeOp = mNodeOP
End Property

'devolver el caption "=" resultado, no es necesario, es solo comodidad
Public Function GetRes() As String
    If mType = ListaCombo Then GetRes = mCaption + "=" + Me.Internal_ListaSImple.GetSelectOp
    If mType = TextoSimple Then GetRes = mCaption + "=" + Me.Internal_TextoSimple.Valor
    If mType = SelectPath Then GetRes = mCaption + "=" + Me.Internal_TextoSimple.Valor
    If mType = Numero Then GetRes = mCaption + "=" + CStr(Me.Internal_Numeros.Valor)
End Function

Public Property Get PlusInfo() As String
    PlusInfo = mPlusInfo
End Property

Public Property Let PlusInfo(newInfo As String)
    mPlusInfo = newInfo
End Property

Public Property Get Marca() As String
    Marca = mMarca
End Property

Public Property Let Marca(newInfo As String)
    mMarca = newInfo
End Property

Public Property Get InternalReal() As String
    InternalReal = mInternalReal
End Property

Public Property Let InternalReal(newInfo As String)
    mInternalReal = newInfo
End Property
