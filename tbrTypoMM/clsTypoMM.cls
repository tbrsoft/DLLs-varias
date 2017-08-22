VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypoMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private mNombre As String 'por ejemplo musica video karaoke ..
Private mDescripcion As String
Private mCantidadPedida As Long 'para el carrito contador de unidades de este tipo que se quieren comprar
Private mHelpUse As String 'ayuda para usarlo, por ejemplo juego java = revise su tamaño de pantalla y configuracion CLDC

Private mCanPlay As Long 'indica si se puede reproducir (se refiere a hacerlo por dinero)
Private mCanExpen As Long 'indica si se puede vender con el expendedor
Private mCanPreview As Long 'indica si se puede hacer una previsualizacion al pasar por el en una lista (mp3 puede solo si esta configurado, las imagenes, los juegos java si pueden. por ejemplo los videos 3GP no pueden)

'anteriormente existia la configuración "VentasExtras" que si era falsa
'solo se verian musica videos y karaokes (modo fonola), si era true se mostraban TODOS
'ahora necesito que el cliente defina que contenido desea y cual no
Private mActivo As Long
'seguiraqui quizas necesite para cada tipo activar por separado la reproduccion y la descarga

Private mIdTypo As MMTypes
'***************************************************************************
'CONDICIONES PARA EL VALIDATE


Private mEXTs() As Extes 'extenciones que puede presentar un archivo con esta multimedia

'el tamaño maximo me puede servir para dado un archivo con una extencion que usa mas de un tipo (caso ISO o MP3)
'los ISO de mas de 700MB imagino que son de DVD. No estoy seguro de como medir el tamaño real de un ISO expandido
'los MP3 pueden ringtones o canciones normales, tambien pueden ser entrevistas u otro tipo de contenido
'debe haber una forma de que la extencion no sea lo definitivo
Private mMinSize As Single 'minimo en mb
Private mMaxSize As Single 'maximo en mb

Private mPuedeSerOtro As Long 'lo necesito por que los que se detectan como ringtones puede ser MP3s, ver mas info rn la validacion

'seguiraqui
'si hay un jpg y hay otro archivo con el mismo nombre en otro formato significa que le jpg es solo una muestra
'habra que descartar esos JPGs como "wallpapers"

'seguiraqui
'el tamaño en pixeles deberia tambien diferencias wallpapers de PC de los de celular

'seguiraqui
'tapa.jpg es un nombre reservado, no debe poner en las listas ni sumar como wallapaper
'***************************************************************************


'TODO LO DE LOS PRECIOS para no mezclar cosas aqui
Public PR As clsPR

Dim FSO As New Scripting.FileSystemObject

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

'seguiraqui
'esto no deberia estar aqui
Public Property Get CantidadPedida() As Long
    CantidadPedida = mCantidadPedida
End Property

Public Property Let CantidadPedida(newV As Long)
    mCantidadPedida = newV
End Property

Public Property Get PuedeSerOtro() As Long
    PuedeSerOtro = mPuedeSerOtro
End Property

Public Property Let PuedeSerOtro(newV As Long)
    mPuedeSerOtro = newV
End Property

Public Property Get HelpUse() As String
    HelpUse = mHelpUse
End Property

Public Property Let HelpUse(newV As String)
    mHelpUse = newV
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

Public Property Get CanPreview() As Long 'permite una prueba gratuita del archivo, se trata mas de la posibilidad de mostrarlo que del permiso. Por ejemplo los archivos ISO no tiene preview
    CanPreview = mCanPreview
End Property

Public Property Get Activo() As Long 'dice si se usara y por lo tanto se buscara al autodetectar
    Activo = mActivo
End Property

Public Property Let Activo(newVal As Long)   'dice si se usara y por lo tanto se buscara al autodetectar
    mActivo = newVal
End Property

Public Sub Extension_Clean()
    ReDim mEXTs(0)
End Sub

Public Function Extencion_Add(sVal As String) ' As Extes ----jode del lado de cfg3pm ?????????
    Dim J As Long
    J = UBound(mEXTs) + 1
    ReDim Preserve mEXTs(J)
    mEXTs(J).sName = sVal
    mEXTs(J).isDupli = False 'predeterminado
    
    Extencion_Add = mEXTs(J)
End Function

'para asignar desde afuera el dupli
Public Function GetExt(i As Long) As Extes
    If i > UBound(mEXTs) Then
        GetExt = mEXTs(0)
    Else
        GetExt = mEXTs(i)
    End If
End Function

Public Function GetExtsCant() As Long
    GetExtsCant = UBound(mEXTs)
End Function

Private Sub Class_Initialize()
    Extension_Clean
    Set PR = New clsPR 'lo inicio aqui ya que los new en declaraciones me caen mal
End Sub

Public Property Get MinSize() As Single
    MinSize = mMinSize
End Property

Public Property Get MaxSize() As Single
    MaxSize = mMaxSize
End Property

Public Property Let MinSize(mVal As Single)
    mMinSize = mVal
End Property

Public Property Let MaxSize(mVal As Single)
    mMaxSize = mVal
End Property

'funcion que valida si determinado archivo es de este tipo
'como puede haber grises el resultado no es boolean si no que es long
'la cantidad de condiciones que cumpla sera el resultado
'por lo tanto el tipo que devuelva mayor valor sera el mas acertado

Public Function Validate(ByVal sFile As String) As Long

    'la propiedad ACTIVO de los typos me dice si se usaran o NO (NUEVO AGO-09)
    If mActivo = 0 Then 'ya se verifica por otro lado pero por las dudas !!!
        tErr.AppendSinHist "NoTypo:" + CStr(mIdTypo) + ":" + sFile 'no deberia entrar aqui
        Validate = 0
        Exit Function
    End If

    Dim RES As Long: RES = 0

    Dim ext As String, i As Long
    ext = FSO.GetExtensionName(sFile)
    
    For i = 1 To UBound(mEXTs)
        If LCase(ext) = LCase(mEXTs(i).sName) Then
            RES = RES + 1
            'si la extencion tambien es usada por otro tipo hay que urgar un poco mas para comparar de que tipo es realmente
            If mEXTs(i).isDupli Then RES = RES + OderValid(sFile)
            
            Validate = RES
            Exit Function
        End If
    Next i
    
    Validate = -1 'NO EXISTE EN LOS REGISTROS ESA EXTENCION!!!.
    'Puede ser un archivo mio .3H o alguna basurilla por ahi
    
End Function

Private Function OderValid(ByVal sFile As String) As Long
    'solamente si hay extenciones repetidas deberia pasar a otras validaciones (la validacion de tamaño puede ser pesada)
    'si la hace siempre sera lento entrar a un disco
    
    Dim RES2 As Long: RES2 = 0
    
    Dim siz As Single, fl As File
    Set fl = FSO.GetFile(sFile)
    siz = Round(((fl.Size / 1024) / 1024), 2)
    
    If siz < mMaxSize And siz > mMinSize Then
        RES2 = RES2 + 1
    End If
    
    OderValid = RES2
    
End Function

Public Property Get IdTypo() As MMTypes
    IdTypo = mIdTypo
End Property

Public Property Let IdTypo(mVal As MMTypes)
    mIdTypo = mVal
End Property

