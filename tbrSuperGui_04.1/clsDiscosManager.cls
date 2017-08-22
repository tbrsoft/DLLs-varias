VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiscoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'========================================================
'Medidas de Marcos y Etiquetas

Dim q_x_marco As Long
Dim q_y_marco As Long
Dim q_ancho_marco As Long
Dim q_alto_marco As Long

Dim q_x_etiqueta As Long
Dim q_y_etiqueta As Long
Dim q_ancho_etiqueta As Long
Dim q_alto_etiqueta As Long
'========================================================



'Cuando Cambio de pagina no tengo que cambiar el color label seleccionado
'por que hay un error... necesito saber cuando se llamo a cambiar pagina
Dim CambioDePagina As Boolean

'Para comprobar clics en regiones'
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
'regiones para clic
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Dim UltimoDiscoSeleccionado As Long 'es del manu, se refiere al ultimo disco que se pinto
'seguiraqui SE debe unificar con mIndiceAbsoluto si se confirma que se necesitan y son lo mismo

'Dim mTouchDiscos() As Label

Dim ShowPNG As New tbrPNG_3.tbrPNG3

'definiciones del teclado
Public PasaHojaAdelante As Boolean 'indica si los botones der-izq al final de una linea pasan a la pagina que sigue o si bajan
Public PasaHojaArriba As Boolean

Public MostrarModoNumpad As Boolean 'ver si se muestra como las rockolas viejas (modo numerico)
'--------------------------------------------------------------------
Private mListaRitmos() As String 'nombres de los ritmos
Private mIndiceRitmos() As Long 'ubicacion de los inicios de los ritmos
'--------------------------------------------------------------------
Private mDiscosHorizontal As Long
Private mDiscosVertical As Long

'y los temas para el modo numpad (requiere saber todas las canciones de todos los discos al inciar, no los descubre cuando entra)
Private mMtxTemas() As String 'tiene dos valores
'(ubound(mtxdiscos),30) 'hasta 30 temas
'debe haber una matriz solo con el nombre y otra con los paths

'negrada
Private mMTXDiscos() As String
Private mMTXPathDiscos() As String

'ubicacion del indice de dicos general y relativo a la pagina
Private mIndiceAbsoluto As Long 'del elegido
Private mIndiceRelativo As Long 'del elegido
'totales
Private mTotalDiscos As Long
Private mTotalPaginas As Long
Private mSeparacionEntreDiscosH As Long 'para que no queden pegoteados
Private mSeparacionEntreDiscosV As Long
Private mnPagActual As Long
Private mIndiceInicioPag As Long 'no se expone es interno
Private mIndiceFinPag As Long 'no se expone es interno
Private mPathDiscoElegido As String
Private mNombreDiscoElegido As String
Private mExtencionesUsadas() As String 'que tipos de archivo carga en las listas (mucho mejor que 3pm)

'al cambiar la seleccion
Public Event ChangeSelDisco(nDisco As Long, nDiscoRel As Long, porcDisco As Single, path As String)
Public Event ClickDisco(nDisco As Long, nDiscoRel As Long)
'se indica los nuemros de orden y el path
Public Event LlegoAlUltimo()
Public Event LlegoAlPrimero()


'************************************************
'************************************************
'************************************************



Dim qModoNumerico As Boolean

Dim mDisco() As New clsDiscoGrafico
Dim mDiscoData() As New clsDiscoData


Dim myAlpha As New tbrAlphaB
Dim myAlphaB As New tbrAlphaB
Dim qUsarAlpha As Boolean
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim PNG1 As String 'Archivo PNG Marco Disco
'archivo del marco de los discos SEGUIRAQUI deberia tener dos para que pasar de modo comun a numerico ya este contemplado en el skin

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColor As Long
Dim rFntColorContorno As Long
Dim mColorDiscoSel As Long

Dim qHDCsalvado As Long

Dim qRect As RECT

Dim DestObjHdc As Long
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long
Dim zAnchoMarco As Long
Dim zAltoMarco As Long

Dim FondoHDC As Long
Dim FondoHDC_BUP As Long
Dim FondoHDC_DISCOS_BUP As Long
Dim FondoHDC_Disc As Long
Dim FondoHDC_Disc_BUP As Long
Dim RestauraFondoHDC As Long

'------------------------------------------------------------------------------------------
'Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

'NEGRADA CULPA DEL MANU (no sabe que esta visible y que no)
Public Enum DondeEstoy
    Nada = 0
    Portadas = 1
    Temas = 2
End Enum

Public EstoyEn As DondeEstoy  'indica si estoy con las tapas o viendo _
    el contenido o ranking
    'valores: "portadas", "temas"


Public Sub Renderizar()
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
End Sub

Public Sub RestaurarFondoHDC()
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, RestauraFondoHDC, 0, 0, SRCCOPY
End Sub


Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColor As Long, FntColorContorno As Long)
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    rFntColor = FontColor
    rFntColorContorno = FntColorContorno
End Sub

Sub IniciarPNGs(PNG_Marco_file As String)
    Dim PNG_Token As Long
    
    If fso.FileExists(PNG_Marco_file) = False Then
        terr.AppendLog "ErrLoadPNG__MARCO"
        Exit Sub
    End If
    
    PNG1 = PNG_Marco_file
    
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.GetAnchoAlto PNG_Marco_file, zAnchoMarco, zAltoMarco
    ShowPNG.FreeGDIPlus PNG_Token
    'zAnchoMarco = 250
    'zAltoMarco = 165
End Sub

'lo agrege en oct09 para que se pueda elegir sin tener que inicializar nada
Public Sub SetPNGMarcoDisco(sFilePng As String)
    PNG1 = sFilePng
    If fso.FileExists(sFilePng) = False Then
        'seguiraqui marcar el error ??
    End If
End Sub

Public Function GetPNGMarco() As String
    GetPNGMarco = PNG1
End Function


Public Sub IniciarGraficos(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, _
    AnchoFondo As Long, AltoFondo As Long, UsarAlpha As Boolean, ModoNumerico As Boolean, _
    ColorDiscoSel As Long)
    
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    zX = XFondo
    zY = YFondo
    zAncho = AnchoFondo
    zAlto = AltoFondo
    
    qModoNumerico = ModoNumerico
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, AnchoFondo, AltoFondo
    IniHDC FondoHDC_BUP, AnchoFondo, AltoFondo
    IniHDC FondoHDC_DISCOS_BUP, AnchoFondo, AltoFondo
    IniHDC FondoHDC_Disc, AnchoFondo, AltoFondo
    IniHDC FondoHDC_Disc_BUP, AnchoFondo, AltoFondo
    IniHDC RestauraFondoHDC, AnchoFondo, AltoFondo
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
    BitBlt FondoHDC_BUP, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
    
    BitBlt RestauraFondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
    
    'con estas dos lineas es como que creo un marco de semitransparencia
    'negra para esta clase... como un fondo ;)
    qUsarAlpha = UsarAlpha
    If qUsarAlpha = True Then
        myAlpha.CrearHDC_Color vbBlack
        myAlpha.MostrarAlpha FondoHDC_BUP, 0, 0, 50, AnchoFondo, AltoFondo
        myAlpha.CerrarGraficos
    End If

    'myAlphaB.CrearHDC_Color vbWhite
    mColorDiscoSel = ColorDiscoSel
    myAlphaB.CrearHDC_Color mColorDiscoSel
End Sub

Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(0)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    DeleteObject TempBMP
    DeleteObject ObjCreado2
End Sub

Public Sub ResetDiscos()
    Dim I As Long
    For I = 0 To UBound(mDisco)
        mDisco(I).CerrarGraficos
    Next I
    
    ReDim mDisco(0)
    ReDim mDiscoData(0)
End Sub


'----------------------------------------------
Private Sub Class_Initialize()
    terr.Anotar "aane"
    mIndiceAbsoluto = -1
    mIndiceRelativo = 0
    ReDim Preserve mMTXDiscos(0)
    ReDim Preserve mMTXPathDiscos(0)
    ReDim Preserve mExtencionesUsadas(0)
    
    ReDim mDisco(0)
    ReDim mDiscoData(0)
    ResetDiscos
    
    EstoyEn = Nada
End Sub

Private Sub GetAnchoAltoFont(aForm As Form, CadenaString As String, qFontName As String, qFontSize As Long, qFontBold As Boolean, ByRef ElAncho As Long, ByRef ElAlto As Long)
    ElAncho = 0
    ElAlto = 0

    aForm.FontName = qFontName
    aForm.FontSize = qFontSize
    aForm.FontBold = qFontBold
    aForm.FontItalic = False
    aForm.FontStrikethru = False
    aForm.FontUnderLine = False
    
    
    ElAncho = (CLng(aForm.TextWidth(CadenaString)) / 15)
    ElAlto = (CLng(aForm.TextHeight(CadenaString)) / 15)
End Sub

'--------------------------------------------------
Public Sub IniciarMedidasDiscoGrafico(x_marco As Long, y_marco As Long, ancho_marco As Long, alto_marco As Long, x_etiqueta As Long, y_etiqueta As Long, ancho_etiqueta As Long, alto_etiqueta As Long)
    q_x_marco = x_marco
    q_y_marco = y_marco
    q_ancho_marco = ancho_marco
    q_alto_marco = alto_marco

    '-----------------------------------
    q_x_etiqueta = x_etiqueta
    q_y_etiqueta = y_etiqueta
    q_ancho_etiqueta = ancho_etiqueta
    q_alto_etiqueta = alto_etiqueta
End Sub

Public Function AgregarDiscoGrafico() As clsDiscoGrafico   'Devuelve el elemento agregado
    Dim Ix As Long
    Ix = UBound(mDisco) + 1: ReDim Preserve mDisco(Ix)
    'Iniciar dispositivo...
    mDisco(Ix).CerrarGraficos
    mDisco(Ix).IniciarPNG PNG1
    mDisco(Ix).IniciarMedidas q_x_marco, q_y_marco, q_ancho_marco, q_alto_marco, q_x_etiqueta, q_y_etiqueta, q_ancho_etiqueta, q_alto_etiqueta
    mDisco(Ix).IniciarGraficos FondoHDC, qModoNumerico
    mDisco(Ix).IniciarFuente rFntName, rFntSize, rFntBold, False, False, False, rFntColor, rFntColorContorno
    
    
    Set AgregarDiscoGrafico = mDisco(Ix)
End Function

Public Function GetDiscoGrafico(Index As Long) As clsDiscoGrafico
    Set GetDiscoGrafico = mDisco(Index)
End Function
Public Function AgregarDiscoData() As clsDiscoData   'Devuelve el elemento agregado
    Dim Ix As Long
    Ix = UBound(mDiscoData) + 1: ReDim Preserve mDiscoData(Ix)
    'Iniciar dispositivo...
    Set AgregarDiscoData = mDiscoData(Ix)
End Function

Public Function GetDiscoData(Index As Long) As clsDiscoData
    Set GetDiscoData = mDiscoData(Index)
End Function
'--------------------------------------------------

Public Function GetDiscosGraficosCount() As Long
    GetDiscosGraficosCount = UBound(mDisco)
End Function
Public Function GetDiscosDataCount() As Long
    GetDiscosDataCount = UBound(mDiscoData)
End Function

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(FondoHDC)
    ret = DeleteDC(FondoHDC_BUP)
    ret = DeleteDC(FondoHDC_DISCOS_BUP)
    
    For I = 1 To UBound(mDisco)
        ret = DeleteObject(mDisco(I).IdRegion)
    Next I
    
    myAlpha.CerrarGraficos
    myAlphaB.CerrarGraficos
End Sub

Public Function GetElementoDisco(Index As Long) As clsDiscoGrafico
    Set GetElementoDisco = mDisco(Index)
End Function
'------------------------------------------------------------------------------------------

Public Sub IniciarDiscosGraficos(Optional Cantidad As Long = -1)
    Dim I As Long
    
    'YActual = 0
    
    '1. Subo el Back Up del Fondo
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, FondoHDC_BUP, 0, 0, SRCCOPY
    
    If Cantidad > 0 Then
        For I = 1 To Cantidad
            'en 800x 600 da error (por que hay pocos discos ?)
            mDisco(I).DibujarDiscoGrafico
        Next I
    Else
        For I = 1 To UBound(mDisco)
            mDisco(I).DibujarDiscoGrafico
        Next I
    End If
    
    BitBlt FondoHDC_DISCOS_BUP, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
End Sub

Public Sub SalvarFondoDiscos(HDCsalvado As Long)
    'Hago un Back Up del Fondo
    qHDCsalvado = HDCsalvado
    BitBlt FondoHDC_Disc, 0, 0, zAncho, zAlto, HDCsalvado, zX, zY, SRCCOPY
End Sub
Public Sub RenderFondoDiscos()
    'Hago un Back Up del Fondo
    BitBlt qHDCsalvado, zX, zY, zAncho, zAlto, FondoHDC_Disc_BUP, 0, 0, SRCCOPY
End Sub

Public Sub RestaurarFondo()
    'Subo el Back Up del Fondo
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, FondoHDC_DISCOS_BUP, 0, 0, SRCCOPY
End Sub

Public Sub SeleccionarDisco2(Index As Long)
    On Local Error GoTo errManu
    
    terr.Anotar "aans900"
    'A diferencia de SeleccionarDisco, que es una funcion
    'de datos, SeleccionarDisco2 selecciona el disco graficamente
    Dim MarcoTapa As RECT
    
    Dim mX As Long
    Dim mY As Long
    Dim fAncho As Long
    Dim fAlto As Long
    
    mX = mDisco(Index).X
    mY = mDisco(Index).Y
    fAncho = mDisco(Index).Ancho
    fAlto = mDisco(Indexs).Alto
    
    terr.Anotar "aans901", mX, mY, fAncho, fAlto
    
    If CambioDePagina = False Then
        If UltimoDiscoSeleccionado > 0 Then
            terr.Anotar "aans902", UltimoDiscoSeleccionado
            CambiarColorTituloDisco UltimoDiscoSeleccionado, vbBlack
            terr.Anotar "aans902b"
        End If
        terr.Anotar "aans902c"
        CambiarColorTituloDisco Index, RGB(0, 0, 50)
    End If
    
    terr.Anotar "aans903", zAncho, zAlto, FondoHDC_Disc_BUP, FondoHDC_Disc
    
    CambioDePagina = False
    
    'Cambio el color del titulo del disco---------
    '---------------------------------------------
    
    BitBlt FondoHDC_Disc_BUP, 0, 0, zAncho, zAlto, FondoHDC_Disc, 0, 0, SRCCOPY
    
    terr.Anotar "aans904", mX, mY
    myAlphaB.MostrarAlpha FondoHDC_Disc_BUP, mX + mDisco(Index).GetMarcoTapaX, mY + mDisco(Index).GetMarcoTapaY, _
        100, mDisco(Index).GetMarcoTapaAncho, mDisco(Index).GetMarcoTapaAlto
    
    terr.Anotar "aans905", mDisco(Index).GetMarcoTapaAncho, mDisco(Index).GetMarcoTapaAlto
    UltimoDiscoSeleccionado = Index
    
    terr.Anotar "aans906", Index
    
    Exit Sub
errManu:
    terr.AppendLog "huytatat2", terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub CambiarColorTituloDisco(Index As Long, NuevoColor As Long)
    terr.Anotar "cct-901", Index, NuevoColor
    
    mDisco(Index).CambiarColorTitulo NuevoColor
    
    terr.Anotar "cct-902"
    Renderizar
    'F1.PM.Refresh
    'SalvarFondoDiscos F1.PM.hdc
    
    terr.Anotar "cct-903"
    SalvarFondoDiscos DestObjHdc
    
    terr.Anotar "cct-904"
End Sub

Public Function GetX() As Long
    GetX = zX
End Function

Public Function GetY() As Long
    GetY = zY
End Function

Public Function GetAncho() As Long
    GetAncho = zAncho
End Function

Public Function GetAlto() As Long
    GetAlto = zAlto
End Function

'**********************************************
'**********************************************
'**********************************************

Public Function GetTemas(snDisco As Long, snTema As Long) As String
    If snTema > 30 Then
        GetTemas = "OUTRANGE"
    Else
        GetTemas = mMtxTemas(snDisco, snTema)
    End If
End Function

Public Sub Clear()
    On Error GoTo ErrF1
    terr.Anotar "aamr"
    'limpiar todos los discos que hay
    Dim TotToClean As Long
    
    TotToClean = (mDiscosHorizontal * mDiscosVertical) - 1
    ' y si no hay discos!
    If UBound(mMTXDiscos) = 0 Then GoTo FIN
    terr.Anotar "aamr2", UBound(mMTXDiscos)
    For A = 1 To TotToClean
        'terr.Anotar "aamr4", A, mMTXDiscos(CLng(A))
        
        'MANU17
        'yo aqui hago un unload de mis objetos
        'es una limpieza para recargar todo de nuevo
        'se deberian matar odos los discos_graficos
        
    Next
FIN:
    terr.Anotar "aamr3"
    
    ReDim Preserve mMTXDiscos(0)
    ReDim Preserve mMTXPathDiscos(0)
    
    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), "DispManager.zaas"
    Resume Next
End Sub

'******* EXTENCIONES USADAS

Public Sub AgregarExtencion(newEXT As String)
    terr.Anotar "aams"
    Dim I As Long
    I = UBound(mExtencionesUsadas) + 1
    ReDim Preserve mExtencionesUsadas(I)
    mExtencionesUsadas(I) = newEXT
End Sub

Public Sub CleanExtenciones()
    'si no se limpian nunca se multiplican en cada reiniciartbrdc
    terr.Anotar "bang"
    Erase mExtencionesUsadas
    ReDim Preserve mExtencionesUsadas(0)
End Sub
Public Function Extencion(IND As Long) As String
    terr.Anotar "aamt"
    Extencion = mExtencionesUsadas(IND)
End Function

Public Function GetCantEXTs()
    terr.Anotar "banh"
    GetCantEXTs = UBound(mExtencionesUsadas)
End Function

'****************************
'segun las extenciones cargadas entrega una matriz con los archivo en la carpeta de disco elegido o segun el numero que se pasa como parametro
Public Function ArchivosEnDisco(Optional nDisco As Long = -1) As String()
    On Error GoTo ErrF1
    
    'en general se usa para el elegido pero para el caso ModoNumpad
    'no hay elegido solo esta a la vista
    
    Dim PathBuscando As String
    If nDisco > -1 Then
        'si pide un disco en particular se lo doy
        PathBuscando = mMTXPathDiscos(nDisco)
    Else
        PathBuscando = mPathDiscoElegido
    End If
    
    terr.Anotar "aamu", nDisco, PathBuscando
    'devuelve todos los temas ORDENADOS
    Dim tmpArch As String
    Dim CantTemas As Long: CantTemas = 0
    
    Dim ArchSort() As String 'matriz desordenada a ordenar
    ReDim ArchSort(0) 'para que no de error
    
    For A = 0 To UBound(mExtencionesUsadas)
        'sacar todos los temas de una extencion
        tmpArch = Dir(PathBuscando + "*." + mExtencionesUsadas(A))
        Do While Len(tmpArch)
            'AGREGADO 28/10/05
            CantTemas = CantTemas + 1
            'es solo el nombre el path lo tengo en la matriz
            
            'LOS CARGO PARA QUE SE ORDENEN DESPUES
            ReDim Preserve ArchSort(CantTemas - 1)
            ArchSort(CantTemas - 1) = tmpArch
            
            tmpArch = Dir
        Loop
    Next
    
    'MANU17 ordenarlas ahora el ArchSort alfabeticamente
    terr.Anotar "aamw"
    ArchivosEnDisco = ArchSort
    
    Exit Function
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), "DispManager.zaat"
    Resume Next
End Function

'solo pàra saber que tiene algo de MM
Public Function DiscoTieneMM(sPath As String) As Boolean
    On Error GoTo ErrF1
    
    If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    PathBuscando = sPath

    terr.Anotar "aamu6", PathBuscando
    'devuelve todos los temas ORDENADOS
    Dim tmpArch As String
    
    For A = 1 To UBound(mExtencionesUsadas)
        'sacar todos los temas de una extencion
        tmpArch = Dir(PathBuscando + "*." + mExtencionesUsadas(A))
        Do While Len(tmpArch)
        
            terr.Anotar "aamw7"
            DiscoTieneMM = True
            Exit Function
            
        Loop
    Next
    
    'MANU17 ordenarlas ahora el ArchSort alfabeticamente
    terr.Anotar "aamw6"
    
    DiscoTieneMM = False
    Exit Function
    
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), "DispManager.zaat6"
    Resume Next
End Function

Public Sub MostrarTodo(Optional SeleccionarAlTerminar As Boolean = True)
    
    On Error GoTo ErrMostrTodoTbrDC
        
    terr.Anotar "aanf2"
    
    'AnchoTodos = GetDiscoGrafico(0).Alto
    'AltoTodos = GetDiscoGrafico(0).Ancho
    
    mSeparacionEntreDiscosH = (GetAncho - (mDiscosHorizontal * GetAnchoMarco)) / (mDiscosHorizontal + 1)
    mSeparacionEntreDiscosV = (GetAlto - (mDiscosVertical * GetAltoMarco)) / (mDiscosVertical + 1)
    
    terr.Anotar "aanf4", mSeparacionEntreDiscosH, mSeparacionEntreDiscosV
    
    'si es modoNumPad mostrar lo que no esta visible en diseño!
    'solo en el modo 2 por que el 3 ya indica el numero a la derecha junto con la lista
    If MostrarModoNumpad Then
        'manu17
        'yo aca defino como visible la lista de disco y el numero de los discos
        'vos hace lo que te guste
    End If
    
    'cargar uno por uno
    Dim Hor As Long, Ver As Long
    Hor = 1: Ver = 1 'posicion del disco que estoy ubicando en cada momento
    mTotalPaginas = 1
    terr.Anotar "aang"
    
    'mtotal discos y mtotal paginas se deben saber
    mTotalDiscos = UBound(mMTXDiscos)
    
    Dim A As Long
    'la parte de la carga de imagenes serà opcional y configurable
    'podràn cargarse solo 6 picTapas
    'es un ahorro grande de memoria
    terr.Anotar "aang2"
    'aqui debo obtener el total de paginas
    'definir la ubicacion de los pictures
    'los tamaños ya estan definidos antes
    
    'no se muestra aqui ni se carga nada ya que estro se hace con el mostrar
    'pagina que en el otro caso solo hace visible lo ya cargado
    
    'agregar ranking al modo manu
    Dim DSK09 As New clsDiscoGrafico
    
    Set DSK09 = AgregarDiscoGrafico()
    DSK09.X = mSeparacionEntreDiscosH
    DSK09.Y = mSeparacionEntreDiscosV
    
    'solo la ubicacion de una página
    Dim t As String
    For A = 1 To (mDiscosHorizontal * mDiscosVertical) - 1
        terr.Anotar "aang3[" + CStr(A) + "]"
        
        Set DSK09 = AgregarDiscoGrafico()
        
        If Hor = mDiscosHorizontal Then
            Hor = 0
            terr.Anotar "aang4", mDiscosHorizontal
            
            DSK09.X = mSeparacionEntreDiscosH
            
            If Ver = mDiscosVertical Then
                'termino la hoja
                Exit For
            Else
                terr.Anotar "aang5", A
                GetDiscoGrafico(A + 1).Y = GetDiscoGrafico(A).Y + GetAltoMarco + mSeparacionEntreDiscosV
            End If
            Ver = Ver + 1
        Else
            GetDiscoGrafico(A + 1).X = GetDiscoGrafico(A).X + GetAnchoMarco + mSeparacionEntreDiscosH
            GetDiscoGrafico(A + 1).Y = GetDiscoGrafico(A).Y

        End If
        
        
        'mostrar disco
        GetDiscoGrafico(A).Visible = True
        
        Hor = Hor + 1
        
    Next A

    '///////////////////////////////////////////////
    '///////////////////////////////////////////////
    'manu
    IniciarDiscosGraficos
    
    'debo cargar mtotalpaginas
    'de alguna forma debo calcularlo!
    mTotalPaginas = ((mTotalDiscos + 1) \ (mDiscosVertical * mDiscosHorizontal)) + 1
       
    terr.Anotar "aank"
    'al mostrar todo se carga la primera hoja
    If SeleccionarAlTerminar Then
        MostrarPaginaN 1, -1
    Else
        MostrarPaginaN 1, -3
    End If
    
    'ConectarTouchSDiscos
    
    'REFRESCAR TODO
    
    Exit Sub
ErrMostrTodoTbrDC:
    terr.AppendLog terr.ErrToTXT(Err), ".zaau"
    Resume Next
End Sub

'obtener solo la imagen a usar mejor posicionada
'como es un lio buscar la imagene tapa jpg y si se pasa del tamaño poner la predet
'directamente cada vez que haga falta esta funcion resuleve el problema

Public Function GetPathPortada(ByVal PathDisco As String)

    If Right(PathDisco, 1) <> "\" Then PathDisco = PathDisco + "\"
    Dim PortadaElegida As String
    
    Dim MemoriaEstaImg As Single
    Dim TapaDef As String, TapaLoc As String 'tapa por defecto y tapa local
    Dim TapaDefGrande As String 'tapa para los que tiene portadas muy grandes en KB
    terr.Anotar "aani[" + PathDisco + "]"
    TapaLoc = PathDisco + "tapa.jpg"
    
    TapaDef = TapaJPGDef 'AP + "tapa.jpg"
    
    If fso.FileExists(AP + "tapa2.jpg") Then
        TapaDefGrande = Tapa2JPGDef 'AP + "tapa2.jpg"
    Else
        TapaDefGrande = TapaJPGDef 'AP + "tapa.jpg"
    End If
    
    'ver si tiene tapa
    If Dir(TapaLoc) = "" Then
        terr.Anotar "aadr", TapaDef
        GetPathPortada = TapaDef
    Else
        terr.Anotar "aads"
        MemoriaEstaImg = FileLen(TapaLoc)
        Dim M2 As Single
        M2 = FileLen(TapaDefGrande)
        Dim MaxTam As Long
        MaxTam = CFG.LeerValor("TamMaxTapa", "50")
        MaxTam = MaxTam * 1024
        'si tiene mas del tamaño permitido...
        If MemoriaEstaImg > MaxTam Then
            'hay ahorro de memoria!!!
            terr.Anotar "aalb*", TapaDefGrande
            GetPathPortada = TapaDefGrande
        Else
            terr.Anotar "aalb", TapaLoc
            GetPathPortada = TapaLoc
        End If
    End If

End Function

'si se hace click en alguna letra necesito que se elija el primer disco con esa letra
Public Function GetNextDiskByLetter(sLetra As String) As Long
    'devuelve el numero de disco que corresponde
    
    Dim M As Long, Letra As String, Letra2 As String
    Letra2 = LCase(sLetra)
    
    Dim MM As Long 'retorno
    MM = -1
    
    For M = mIndiceAbsoluto To UBound(mMTXDiscos)
        Letra = LCase(Left(mMTXDiscos(M), 1))
        
        'tener en cuentas caracteres extraños
        If InStr("abcdefghijklmnñopqrstuvwxyz", Letra) = 0 Then Letra = "#"
        
        If Letra = Letra2 Then
            MM = M
            Exit For
        End If
    Next M
    
    'si no encontro devolvera -1
    GetNextDiskByLetter = MM
        
End Function

'carga la pagina que me piden
'se usa para pasar una a la que sigue o para mostrar cualquiera solicitada

Public Sub MostrarPaginaN(n As Long, SeleccionarPrimero As Long, _
    Optional SeleccionarLaPaginaDelDisco As Long = -1)
        
    CambioDePagina = True
    'puede abrir una pàgina "n" conocida por que va en orden
    'o puede abrir la pagina que tiene el disco "SeleccionarLaPaginaDelDisco"
    
    'si seleccionar primero es -1 es true
    'el -2 es el false (selecciona el ultimo)
    'si es un numero es el indice ABSOLUTO en la pagina!
    
    On Error GoTo ErrF1
    Dim A As Long
    Dim DiscosXpag As Long
    DiscosXpag = mDiscosHorizontal * mDiscosVertical
    
    terr.Anotar "aanm[" + CStr(n) + ":" + CStr(SeleccionarPrimero) + ":" + _
        CStr(SeleccionarLaPaginaDelDisco) + "]"
    
    If SeleccionarLaPaginaDelDisco > -1 Then
        'calcular la página N conociendo el numero de disco
        'todas las páginas tiene la misma cantidad "DiscosXpag"
        'divido por la cantidad por pagina y le sumo uno (NOTESE QUE ES DIVISION ENTERA)
        'ej: disco 1 = (1+1\6) + 1=pagina 1
        'ej: disco 10 = (10\6) + 1 = 1+1 (segunda pagina)
        'ej: disco 7 = (7\6) + 1 = 1+1 (segunda pagina)
        'el mas uno es por el ranking que corre todo
        '***********************
        'decia seleccionarla... + 1 y le saque el mas 1
        '***********************
        n = (SeleccionarLaPaginaDelDisco \ DiscosXpag) + 1
        terr.Anotar "aadi[" + CStr(n) + "]"
    End If
        
    If n > TotalPaginas Then Exit Sub
    If n < 1 Then Exit Sub
    'ver cual es el primer discos de una pagina
    
    'mostrar desde y mostrar hasta dan el numero de disco
    'si no carga on demand tamboien es el numero de pictapa
    Dim MostrarDesde As Long, MostrarHasta As Long
    MostrarDesde = (n * DiscosXpag) - DiscosXpag
    MostrarHasta = (n * DiscosXpag) - 1
    
'    'rrrr para quitar el rank!
'    MostrarDesde = MostrarDesde + 1
'    MostrarHasta = MostrarHasta + 1
'
    terr.Anotar "aadj[" + CStr(MostrarDesde) + ":" + CStr(MostrarHasta) + "]"
    'ver si mostrar hasta se pasa en el total de discos.
    'si es la ultima hoja puede dar error
    '-----------------
    Dim Sobran As Long
    Sobran = 0
    If MostrarHasta > UBound(mtxDiscos) Then
        terr.Anotar "aann"
        'este es el caso de la ultima página
        Sobran = MostrarHasta - UBound(mtxDiscos)
        MostrarHasta = UBound(mtxDiscos)
    End If
    
    '///////////////////////////////////
    '///////////////////////////////////
    RestaurarFondo
    Dim DG As New clsDiscoGrafico
    Dim DD As New clsDiscoData
    'si estoy en modo numpad debo cargar las canciones de cada disco iaaa
    Dim LL As New clsTemasManager
    
    '///////////////////////////////////
    '///////////////////////////////////
    
    terr.Anotar "aadk"
    'aca no hay indices mayores que valgan es solo desde 0 hasta H x V
    'solo paso una vez cargando encima de lo que habia las imagenes _
        que corresponden
    'uso una segunda variable que siempre va desde cero
    
    IniciarDiscosGraficos (MostrarHasta - MostrarDesde) + 1 '- Sobran
    
    For A = MostrarDesde To MostrarHasta + Sobran
        If A <= MostrarHasta Then
            'a-mostradesde serias 0,1,2,...H x V
            terr.Anotar "aadl[" + CStr(A) + ":" + CStr(A - MostrarDesde) + "]"
            'rrrr para quitar el rank empiezo en el disco 1'
            
            'cargarportada era un loadpicture
            'CargarPortada A, (A - MostrarDesde)
            
            If MostrarModoNumpad Then
                terr.Anotar "aalc"
                'manumeric pulenta aqui define que numero lleva el disco
                'nDisco7SEG(A - MostrarDesde).Caption = SuperString(CStr(A), "0", 3)
            End If
                
            If A > 0 Then
                
                terr.Anotar "aadm", A + 1, DD.Titulo
                
                If MostrarModoNumpad Then
                    GetDiscoData(A - MostrarDesde + 1).Titulo = mMTXDiscos(CLng(A))
                    CargarTemas LL, A, A - MostrarDesde
                End If
                
                '////////MANU
                Set DG = GetDiscoGrafico(A - MostrarDesde + 1)
                Set DD = GetDiscoData(A + 1)
                DG.DibujarDiscoData DD, False, LL
                
                
            Else
                Set DG = GetDiscoGrafico(1)
                Set DD = GetDiscoData(1)
                DG.DibujarDiscoData DD, False, LL
            End If
        Else
            terr.Anotar "aadn"
            'para los que sobran por que es ala ultima pagina
            GetDiscoGrafico(A - MostrarDesde).Visible = False
        End If
    
    Next A
    
    '-----------------------------------------
    mIndiceInicioPag = MostrarDesde
    mIndiceFinPag = MostrarHasta
    terr.Anotar "aano"
    '-----------------
    mnPagActual = n
    terr.Anotar "aanp"
    'el indice absoluto y el relativo se cargan en "seleccionardisco"
    ' o sea que los llamo igual y alla se ubica
    
    Renderizar
    SalvarFondoDiscos DestObjHdc
    
    If SeleccionarPrimero = -1 Then SeleccionarDisco MostrarDesde, MostrarModoNumpad
    'selecciona el ultimo
    If SeleccionarPrimero = -2 Then SeleccionarDisco MostrarHasta, MostrarModoNumpad
    'si no seleccionar el que corresponde
    If SeleccionarPrimero > -1 Then SeleccionarDisco SeleccionarPrimero, MostrarModoNumpad
    ' SI NO QUIERO SELECCIONAR NINGUNO
    If SeleccionarPrimero = -3 Then SeleccionarDisco MostrarDesde, True
    
    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaau"
    Resume Next
End Sub

Public Property Get nPaginaActual() As Long
    terr.Anotar "aanq"
    nPaginaActual = mnPagActual
End Property

Public Sub SeleccionarPaginaSig(Adelante As Boolean)
    On Error GoTo ErrF1
    'permite avanzar o retroceder de a una pàgina
    Dim NewInd As Long
    If Adelante Then
        'va al primero de la pàgina que sigue
        NewInd = mIndiceInicioPag + (mDiscosHorizontal * mDiscosVertical)
        'NewInd = mIndiceAbsoluto + (mDiscosHorizontal * mDiscosVertical)
    Else
        NewInd = mIndiceInicioPag - (mDiscosHorizontal * mDiscosVertical)
    End If
    
    'UNA DE TRES:
        '1- LLEGO AL INICIO Y VA A A ULTIMA PAGINA
        '2- LLEGO AL FIN Y VA A LA PRIMERA PÁGINA
        '3- SALE DE UNA PAGINA Y VA A OTRA
    
    '1
    If NewInd < 0 Then
        RaiseEvent LlegoAlPrimero
        'pasar al ultimo disco. Hacer ciclico
        MostrarPaginaN mTotalPaginas, -1
        Exit Sub
    End If
    
    '2
    If NewInd > mTotalDiscos Then
        RaiseEvent LlegoAlUltimo
        MostrarPaginaN 1, -1
        Exit Sub
    End If
    
    '3-
    If Adelante Then
        MostrarPaginaN nPaginaActual + 1, -1
    Else
        MostrarPaginaN nPaginaActual - 1, -1
    End If
    '
    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaav"
    Resume Next
End Sub

Public Sub SeleccionarDisco(indice As Long, Optional sOLOpOSICIONARSEsINpINTAR As Boolean = False)
    
    On Error GoTo ErrF1
    
    'siempre se selecciona alguno salvo si estoy en modo numpad que nada se seleccion directamente
        
    'siempre se elige como absuluto por mas que sea onDemand
    'si es ondemand lo traduizco aqui
    
    'ver si el indice se fue a la bosta!!!
    'ya paso...
    If indice > UBound(mMTXPathDiscos) Then
        'mostrarlo y grabarlo
        terr.Anotar "[" + CStr(indice) + "/" + CStr(UBound(mMTXPathDiscos)) + "/" + _
            CStr(mIndiceAbsoluto) + "/" + CStr(mIndiceRelativo) + _
            CStr(mDiscosHorizontal) + "/" + CStr(mDiscosVertical) + "]"
        'grabarlo
        terr.AppendLog terr.ErrToTXT(Err), ".zaax"
        'y salir....
        Exit Sub
    End If
    
    'el parametro es el indice absoluto
    'primero desseleccionar el activo actual
    
    'si carga on demand es otra cosa ya que activa o desactive segun _
        el indice relativo
    Dim IndToSel As Long
    Dim IndToUnSel As Long
        
    'aqui se traduce el numero absoluto al relativo y se guarda asi
    IndToSel = TraslateAbsInRelativ(indice)
    IndToUnSel = TraslateAbsInRelativ(mIndiceAbsoluto)
    mIndiceRelativo = IndToSel
    
    terr.Anotar "aanr", indice, mIndiceAbsoluto, mIndiceRelativo, IndToUnSel
    
    '---------------------------------
    'dejar claro el indice absoluto elegido!
    mIndiceAbsoluto = indice
    '---------------------------------

    terr.Anotar "aans4[" + CStr(IndToSel) + "]"
    'aqui si va en referencia a lo absoluto!
    mPathDiscoElegido = mMTXPathDiscos(indice)
    mNombreDiscoElegido = mMTXDiscos(indice)
    'mostrar tambien el que se desSelecciona!!!!
    terr.Anotar "aans5[" + CStr(IndToUnSel) + "]"
    
    If sOLOpOSICIONARSEsINpINTAR = False Then
        terr.Anotar "aans5b", IndToUnSel + 1
        SeleccionarDisco2 IndToSel + 1
        
        terr.Anotar "aans5c"
        RenderFondoDiscos
        'F1.PM.Refresh
    End If
        
    'lanzar el evento
    'y si es el total de discos es cero!!!
    'cuando lo usa por primera vez!!!!!!!!
    terr.Anotar "aans6"
    If mTotalDiscos > 1 Then
        terr.Anotar "aans7"
        RaiseEvent ChangeSelDisco(mIndiceAbsoluto, IndToSel, mIndiceAbsoluto / mTotalDiscos * 100, mPathDiscoElegido)
    End If
    
    Exit Sub
    
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), "dispMNG.zaao"
    Resume Next

End Sub

Public Sub SeleccionarDiscoSig(Adelante As Boolean)  'el parametro es el indice absoluto
    On Error GoTo ErrF1

    'permite avanzar o retroceder de a uno
    'ver si el que corresponde esta en la misma pagina
    terr.Anotar "aant"

    Dim NewInd As Long

    If Adelante Then
        'puede bajar al final de cada linea o pasar a la hoja que sigue
        'RRRR mIndicerelativo estaba sumado en uno las 2 veces por que empezaba en cero
        If PasaHojaAdelante And (mIndiceRelativo) / mDiscosHorizontal = _
            (mIndiceRelativo) \ mDiscosHorizontal Then
                'esta al borde de una linea horiz
                'debe sumar la pagina competa menos la linea en la que esta
                NewInd = mIndiceAbsoluto _
                    + (mDiscosHorizontal * mDiscosVertical) _
                    - mDiscosHorizontal + 1
                '¿y si la ultima pagina tiene pocos discos?
                'si vengo de la ultima fila esta cuenta lo haria pasar
                'mTotalDiscos y por lo tanto saltearse la ultima página
                'XXXX
        Else
            'o el avance sencillo
            NewInd = mIndiceAbsoluto + 1
        End If
    Else 'EN EL CASO DE QUE SEA PARA ATRAS
        'puede bajar al final de cada linea o pasar a la hoja que sigue
        If PasaHojaAdelante And mIndiceRelativo / mDiscosHorizontal = _
            mIndiceRelativo \ mDiscosHorizontal Then
                'esta al borde de una linea horiz
                'debe restar la pagina competa menos la linea en la que esta
                NewInd = mIndiceAbsoluto _
                    - (mDiscosHorizontal * mDiscosVertical) _
                    + mDiscosHorizontal - 1
        Else
            'o el movimiento sencillo
            NewInd = mIndiceAbsoluto - 1
        End If
    End If
    If NewInd < 0 Then
        RaiseEvent LlegoAlPrimero
        'Ir al ultimo!!!!!
        MostrarPaginaN mTotalPaginas, -2
        Exit Sub
    End If
    If NewInd > mTotalDiscos Then
        RaiseEvent LlegoAlUltimo
        'Ir al primero!!!!!
        'XXXX
        MostrarPaginaN 1, -1
        Exit Sub
    End If
    'primero desseleccionar el activo actual
    terr.Anotar "aanu"

    If NewInd >= mIndiceInicioPag And NewInd <= mIndiceFinPag Then
    '    'esta en la misma pagina
        SeleccionarDisco NewInd, MostrarModoNumpad 'ultimo parametro agregado
    '    mIndiceAbsoluto = NewInd
    Else
        terr.Anotar "aanv"
        'aqui hago que al pasar de hoja (adelante) vaya al ultimo, de esta forma
        'el boton sig puede ser pasar páginas!!!
        'pero si la ultima página no esta llena!!!. Da error. Corregir
        'If NewInd < mIndiceInicioPag Then MostrarPaginaN nPaginaActual - 1, -1
        'If NewInd > mIndiceFinPag Then MostrarPaginaN nPaginaActual + 1, -2

        'asi funcionaba hasta el 0.9.644
        'va al primero si o si si cambia de pagina
        'If NewInd < mIndiceInicioPag Then MostrarPaginaN nPaginaActual - 1, -2
        'If NewInd > mIndiceFinPag Then MostrarPaginaN nPaginaActual + 1, -1

        'ahora directamente al que le toca
        If NewInd < mIndiceInicioPag Then
            MostrarPaginaN nPaginaActual - 1, NewInd ', NewInd
        End If
        If NewInd > mIndiceFinPag Then
            MostrarPaginaN nPaginaActual + 1, NewInd ', NewInd
        End If
    End If
    'terr.anotar "aanw"
    'mPathDiscoElegido = mMTXPathDiscos(mIndiceAbsoluto)
    'mNombreDiscoElegido = mMTXDiscos(mIndiceAbsoluto)

    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaay"
    Resume Next
End Sub

Public Sub SeleccionarDiscoArriba(Arriba As Boolean)
    On Error GoTo ErrF1

    'permite avanzar o retroceder de a uno
    'ver si el que corresponde esta en la misma pagina
    terr.Anotar "aant2"
    Dim NewInd As Long

    If PasaHojaArriba Then
        If Arriba Then
            'ver si es la primera columna
            If mIndiceRelativo < mDiscosHorizontal Then
                NewInd = mIndiceAbsoluto - mDiscosVertical - 1
            Else 'NO ES LA PRIMERA COLUMNA
                NewInd = mIndiceAbsoluto - mDiscosHorizontal
            End If
        Else 'ABAJO
            'ver si es la ultima columna
            If (mIndiceRelativo + 1) > _
            (mDiscosHorizontal * mDiscosVertical) - mDiscosHorizontal Then
                NewInd = mIndiceAbsoluto + mDiscosHorizontal
            Else 'NO ES LA PRIMERA COLUMNA
                NewInd = mIndiceAbsoluto + mDiscosHorizontal
            End If
        End If

    Else 'PasaHojaArriba=FALSE
        If Arriba Then
            'ver si es la primera columna
            If mIndiceRelativo < mDiscosHorizontal Then
                'SI ES ADEMAS LA PRIMERA FILA SOLO RESTA UNO
                If mIndiceRelativo = 0 Then
                    NewInd = mIndiceAbsoluto - 1
                Else 'NO ES LA PRIMERA FILA
                    NewInd = mIndiceAbsoluto - 1 + _
                        (mDiscosVertical - 1) * mDiscosHorizontal
                End If
            Else 'NO ES LA PRIMERA COLUMNA
                NewInd = mIndiceAbsoluto - mDiscosHorizontal
            End If
        Else 'ABAJO
            'ver si es la ultima columna
            If (mIndiceRelativo + 1) > _
                (mDiscosHorizontal * mDiscosVertical) - mDiscosHorizontal Then

                'SI ES LA ULTIMA FILA TAMBIEN SOILO SUMA UNO (A LA PAG QUE SIGUE)
                If (mIndiceRelativo + 1) = mDiscosHorizontal * mDiscosVertical Then
                    NewInd = mIndiceAbsoluto + 1
                Else 'NO ES LA ULTIMA FILA
                    'en este caso sumo uno para adelantar columna
                    'y resto una pagina entera (con una fila menos) para ir a la primera de la columna siguiente
                    NewInd = mIndiceAbsoluto + 1 - _
                        (mDiscosVertical - 1) * mDiscosHorizontal
                End If
            Else
                NewInd = mIndiceAbsoluto + mDiscosHorizontal
            End If
        End If
    End If

    If NewInd < 0 Then
        RaiseEvent LlegoAlPrimero
        'Ir al ultimo!!!!!
        MostrarPaginaN mTotalPaginas, -2
        Exit Sub
    End If
    If NewInd > mTotalDiscos Then
        RaiseEvent LlegoAlUltimo
        'Ir al primero!!!!!
        'XXXX
        MostrarPaginaN 1, -1
        Exit Sub
    End If
    'primero desseleccionar el activo actual
    terr.Anotar "aanu2"

    If NewInd >= mIndiceInicioPag And NewInd <= mIndiceFinPag Then
        'esta en la misma pagina
        SeleccionarDisco NewInd, MostrarModoNumpad 'ultimo parametro agregado
    Else
        terr.Anotar "aanv"
        'aqui hago que al pasar de hoja (adelante) vaya al ultimo, de esta forma
        'el boton sig puede ser pasar páginas!!!
        'pero si la ultima página no esta llena!!!. Da error. Corregir
        'If NewInd < mIndiceInicioPag Then MostrarPaginaN nPaginaActual - 1, -1
        'If NewInd > mIndiceFinPag Then MostrarPaginaN nPaginaActual + 1, -2

        If NewInd < mIndiceInicioPag Then MostrarPaginaN nPaginaActual - 1, NewInd ' -2
        If NewInd > mIndiceFinPag Then MostrarPaginaN nPaginaActual + 1, NewInd ' -1
    End If

    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaay"
    Resume Next
End Sub

Public Property Get PathDiscoElegido() As String
    terr.Anotar "aanx", mPathDiscoElegido
    PathDiscoElegido = mPathDiscoElegido
End Property

Public Property Get NombreDiscoElegido() As String
    terr.Anotar "aany"
    NombreDiscoElegido = mNombreDiscoElegido
End Property

Public Property Get TotalDiscos() As Long
    terr.Anotar "aanz"
    TotalDiscos = mTotalDiscos
End Property

Public Property Get TotalPaginas() As Long
    terr.Anotar "aaoa"
    TotalPaginas = mTotalPaginas
End Property

Public Property Get DiscosHorizontal() As Long
    terr.Anotar "aaob"
    DiscosHorizontal = mDiscosHorizontal
End Property

Public Property Let DiscosHorizontal(NewValue As Long)
    terr.Anotar "aaoc"
    mDiscosHorizontal = NewValue
End Property

Public Property Get DiscosVertical() As Long
    terr.Anotar "aaod"
    DiscosVertical = mDiscosVertical
End Property

Public Property Let DiscosVertical(NewValue As Long)
    terr.Anotar "aaoe"
    mDiscosVertical = NewValue
End Property
'leer o cargar la matriz de una
Public Property Get mtxDiscos() As String()
    terr.Anotar "aaof"
    'matriz de los nombre de los discos
    mtxDiscos = mMTXDiscos
End Property

Public Function GetNombreDisco(Index As Long) As String
    If Index > UBound(mMTXDiscos) Then
        GetNombreDisco = "OUT"
    Else
        GetNombreDisco = mMTXDiscos(Index)
    End If
End Function

Public Function GetPathDisco(Index As Long) As String
    If Index > UBound(mMTXPathDiscos) Then
        GetPathDisco = "OUT"
    Else
        GetPathDisco = mMTXPathDiscos(Index)
    End If
End Function

Public Property Let mtxDiscos(NewMTX() As String)
    terr.Anotar "aaog"
    mMTXDiscos = NewMTX
End Property

'leer o cargar la matriz de una
Public Property Get MTXPathDiscos() As String()
    terr.Anotar "aaoh"
    'matriz de los nombre de los discos
    MTXPathDiscos = mMTXPathDiscos
End Property

Public Property Let MTXPathDiscos(NewMTX() As String)
    terr.Anotar "aaoi"
    mMTXPathDiscos = NewMTX
End Property

'agregar o quitar discos
Public Sub AgregarDisco(PathFull As String) 'recibe todo y lo separa
    On Error GoTo ErrF1

    'me asegurop que tenga la barra para poder saber bien cual es la ultima carpeta
    terr.Anotar "aaoj"
    If Right(PathFull, 1) <> "\" Then PathFull = PathFull + "\"
    Dim LastDisk As Long
    LastDisk = UBound(mMTXDiscos)
    LastDisk = LastDisk + 1
    terr.Anotar "aaok"
    ReDim Preserve mMTXDiscos(LastDisk)
    ReDim Preserve mMTXPathDiscos(LastDisk)
    'primero el nombre y despues todo
    mMTXDiscos(LastDisk) = txtLista(PathFull, 99998, "\")
    mMTXPathDiscos(LastDisk) = PathFull

    'mtotal discos y mtotal paginas se deben saber
    mTotalDiscos = UBound(mMTXDiscos)

    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaaz"
    Resume Next
End Sub

Public Sub OrdenarDiscos(Optional modoSort As Long = 2)
    On Error GoTo ErrF1
    'modoSort puede valer...

    'debe haber varios modos de ordenar:
    '1: Ignorar los origenes y ordenar por nombre del disco nada mas
    '2: Colocar un origen detras de otro y cada origen ordenado
    '3: desde el disco mas escuchado hasta el menos _
        (empieza bien pero despues se va perdiendo el orden) _
        ademas que hago con los que estan iguales?

    'por defecto es el 2

    terr.Anotar "aaol"

    Dim ORIGs() As String
    ORIGs = mpRock.ORG_GetOrigenes(, True)

    Dim parteEnLST As String
    Dim ContRealOrden As Long
    ContRealOrden = 0
    'estos son contadores de origenes validos
    ReDim Preserve mListaRitmos(0)
    ReDim Preserve mIndiceRitmos(0)
    
    Dim PartOrig() As String, resumenINDS As String
    resumenINDS = ""
    
    For B = 0 To UBound(ORIGs)
        PartOrig = Split(ORIGs(B), ";")
        'si el origen no esta activa saltear e ir al que sigue
        If ORIGs(B) = "" Then GoTo sig666
        If PartOrig(1) = "0" Then GoTo sig666
        '--------------------------------------------
        'entra este origen
        Dim NewI As Long
        NewI = UBound(mListaRitmos) + 1
        'el nombre esta en PartOrig.....
        ReDim Preserve mListaRitmos(NewI)
        mListaRitmos(NewI) = PartOrig(0)
        'y el numero son los discos ordenados ya
        ReDim Preserve mIndiceRitmos(NewI)
        mIndiceRitmos(NewI) = ContRealOrden + 1
        resumenINDS = resumenINDS + "ritmo[" + CStr(NewI) + "]=" + mListaRitmos(NewI) + " IND=" + CStr(mIndiceRitmos(NewI)) + vbCrLf
        
        '---------------------------------------------
        'SEGUIRAQUI - QUE SE ORDENE !!!
        ''nuevas matrices temporales con todo ordenado
        'Dim m2MTXDiscos() As String
        'Dim m2MTXPathDiscos() As String
        ''---------------------------------------------
        For A = 1 To UBound(mMTXPathDiscos)
        '    'si corresponde al origen elegido entonces lo carga!
        '    'si no que espere
            parteEnLST = LCase(Left(mMTXPathDiscos(A), Len(PartOrig(0))))
            If LCase(PartOrig(0)) = parteEnLST Then
                ContRealOrden = ContRealOrden + 1
        '        ReDim Preserve m2MTXDiscos(ContRealOrden)
        '        m2MTXDiscos(ContRealOrden) = txtLista(mMTXPathDiscos(A), 99998, "\")
        '        ReDim Preserve m2MTXPathDiscos(ContRealOrden)
        '        m2MTXPathDiscos(ContRealOrden) = mMTXPathDiscos(A)
            End If
        Next A
sig666:
    Next B
    
    terr.AppendSinHist "SortOK: " + vbCrLf + resumenINDS
    
    terr.Anotar "aaon"

    Exit Sub
ErrF1:
    terr.AppendLog terr.ErrToTXT(Err), ".zaaz"
    Resume Next
End Sub

Public Function GetIndicePrimerDiscoRitmoNext(Posicionarse As Boolean) As Long
    Dim M As Long, Encontro As Boolean
    Encontro = False
    For M = 1 To UBound(mIndiceRitmos)
        If mIndiceRitmos(M) > mIndiceAbsoluto Then
            GetIndicePrimerDiscoRitmoNext = mIndiceRitmos(M)
            If Posicionarse Then MostrarPaginaN 1, mIndiceRitmos(M), mIndiceRitmos(M)
            'si no va entra al que sigue!!!
            Encontro = True
            GoTo FIN44
        End If
    Next M
    
FIN44:
    If Encontro Then Exit Function
    'entonces esta en el ultimo y debe ir al primero
    GetIndicePrimerDiscoRitmoNext = mIndiceRitmos(1)
    If Posicionarse Then MostrarPaginaN 1, mIndiceRitmos(1), mIndiceRitmos(1)
    
End Function
    
Public Function GetIndicePrimerDiscoRitmoLast(Posicionarse As Boolean) As Long
    Dim M As Long, Encontro As Boolean
    Encontro = False
    For M = UBound(mIndiceRitmos) To 1 Step -1
        If mIndiceRitmos(M) < mIndiceAbsoluto Then
            GetIndicePrimerDiscoRitmoLast = mIndiceRitmos(M)
            If Posicionarse Then MostrarPaginaN 1, mIndiceRitmos(M), mIndiceRitmos(M)
            'si no va entra al que sigue!!!
            Encontro = True
            GoTo FIN45
        End If
    Next M
    
FIN45:
    If Encontro Then Exit Function
    'entonces esta en el primero y debe ir al ultimo
    GetIndicePrimerDiscoRitmoLast = mIndiceRitmos(UBound(mIndiceRitmos))
    If Posicionarse Then MostrarPaginaN 1, mIndiceRitmos(UBound(mIndiceRitmos)), mIndiceRitmos(UBound(mIndiceRitmos))
End Function

'obtener el indice del primer disco del ritmo n
Public Function GetIndiceRitmo(IndexRitmo As Long) As String
    If IndexRitmo > UBound(mIndiceRitmos) Then
        GetIndiceRitmo = -1
    Else
        GetIndiceRitmo = mIndiceRitmos(IndexRitmo)
    End If
End Function
 
Public Function GetNombreRitmo(IndexRitmo As Long) As String
    If IndexRitmo > UBound(mListaRitmos) Then
        GetNombreRitmo = "OUT"
    Else
        GetNombreRitmo = mListaRitmos(IndexRitmo)
    End If
End Function

Public Function txtLista(lista As String, Orden As Long, Separador As String) As String
    'lo copio para que se pueda exportar el control y sea independiente
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    'si pongo 99999 en orden saco el ultimo
    Dim lAct As String, lOrden As Long
    Dim palabra(40) As String
    Dim c As Long
    c = 1: lOrden = 0
    Do While c <= Len(lista)
        lAct = mID(lista, c, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
            If lOrden > Orden Then Exit Do
        End If
        c = c + 1
    Loop
    'si oreden solicitado>ultimo oreden de la lista...
    If Orden > lOrden Then
        If Orden = 99999 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtLista = palabra(lOrden): Exit Function
        End If
        If Orden = 99998 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtLista = palabra(lOrden - 1): Exit Function
        End If
        If Orden <> 99999 And Orden <> 99998 Then
            txtLista = "OUT LISTA": Exit Function
        End If
    End If
    txtLista = palabra(Orden)
End Function

Public Function GetIndiceAbsoluto() As Long
    GetIndiceAbsoluto = mIndiceAbsoluto
End Function

Private Function TraslateAbsInRelativ(IndABS As Long) As Long

    Dim TMP As Long
    'El operador de módulo, o resto, divide número1 por número2 _
        (redondeando a enteros los números de signo flotante) y _
        devuelve sólo el resto como resultado. Por ejemplo, en la _
        siguienteexpresión, A (que es el resultado) es igual a 5.
    'A = 19 Mod 6.7

    TMP = IndABS Mod (mDiscosHorizontal * mDiscosVertical)

    TraslateAbsInRelativ = TMP

End Function

'cargar las listas de canciones
'APARENTEMENTE NO SE USA EN NINGUN LADO !!!!!
'lo dejo por que me suena que si hace falta
Private Sub CargarTemas(LS As clsTemasManager, nDisco As Long, IndexPicTapa As Long)

    terr.Anotar "aaqa[" + CStr(IndexPicTapa) + "]"
    'tbrLST(IndexPicTapa).Clear
    LS.ResetElementos
    
    Dim Partes() As String
    
    Dim A As Long, temaShow As String
    A = 0
    
    Do
        'agregado 28-10-05 NO PUEDE MAS DE 30!! por que las matrices bidimensionales son jodidas!
        If A = 31 Then Exit Do
        
        terr.Anotar "aaqb[" + CStr(nDisco) + "-" + CStr(A) + "]"
        If mMtxTemas(nDisco, A) = "" Then Exit Do
        Partes = Split(mMtxTemas(nDisco, A), "*")
        
        
        
        temaShow = Partes(0)
        'manumeric aqui se agrega cada elemnto de una lista elegida
        'tbrLista es como un listbox pero mio
        'los parametros son
        'el primero es lo que quiero mostrar en la columna de la izq
        'el segundo la duracion de la cancion en la de la derecha
        'el tercero es el path completo, o sea el unico valor pulenta que le importa al sistema
        'ademas esta configurado para que a la izquierda haya una primera columna
        'con los numeros de cancion a DOS digitos (01, 02, etc) para que los usuarios escriban igual
        
        A = A + 1
        
        'MANU17 aqui cargo un elemnto a la lista
        'tbrLST(IndexPicTapa).AddItem temaShow, Partes(1), Partes(2)
        Dim Rng As New clsTemaList
        Set Rng = LS.AgregarElemento
        Rng.Titulo = temaShow
        Rng.Duracion = Partes(1)
        Rng.TagMisterioso = Partes(2)
        Rng.Numero = A - 1
        
    Loop
End Sub

'manu17
'SI SE LLAMA DESDE AFUERA PERO NO HACE NADA !!!??
Public Sub CargarTodosTemasModoNumPad()
    Dim A As Long
    Dim Archs() As String
    Dim MP3tmp As New MP3Info
    Dim DuracionTema As String

    ReDim mMtxTemas(mTotalDiscos, 30)
    
    
    For A = 0 To UBound(mMTXPathDiscos) 'mTotalDiscos 'EL CERO ES RANK!
        Archs = ArchivosEnDisco(A)
        'ahora leer los temas
        '-------------------------
        Dim HastaArch As Long
        If UBound(Archs) >= 30 Then
            HastaArch = 29
        Else
            HastaArch = UBound(Archs)
        End If
        '-------------------------
        'NO SE PUEDE DIVIDIR POR CERO!!!
        If mTotalDiscos = 0 Then mTotalDiscos = 1
        'manu18
        
        
        mMtxTemas(A, 0) = "TODO EL DISCO*ALL*ALL"
        'la matriz arch empieza en cero
        For B = 0 To HastaArch 'DECIA CERO
            terr.Anotar "aaed2"
            If LCase(Right(Archs(B), 3)) = "mp3" Or LCase(Right(Archs(B), 3)) = "wma" Then
                'If crgdur Then
                    MP3tmp.FileName = mMTXPathDiscos(A) + Archs(B)
                    DuracionTema = MP3tmp.DurationSTR
                'Else
                    'tERR.Anotar "aaee2"
                    'DuracionTema = "MP3"
                'End If
            Else
                terr.Anotar "aaef2"
                DuracionTema = "VIDEO"
            End If
            terr.Anotar "aaeg2"
                
            'se agrega MOSTRAR,DURACION,VERDADERO VALOR
            'poner el sufijo VIDEO si corresponde
            If LCase(Right(Archs(B), 3)) = "mp3" Or LCase(Right(Archs(B), 3)) = "wma" Then
                'manu17
                'no se por que lo desactive !!(22/05/2009)
                mMtxTemas(A, B + 1) = _
                    UCase(mpRock.QuitarNumeroInicial(Archs(B))) + _
                    "*" + DuracionTema + "*" + Archs(B)
            Else
                mMtxTemas(A, B + 1) = _
                    UCase(mpRock.QuitarNumeroInicial(Archs(B))) + " (V)" + _
                    "*" + DuracionTema + "*" + Archs(B)
            End If
            
        Next B
    Next A

    terr.Anotar "aaej2"
    Set MP3tmp = Nothing

End Sub


'para saber si existe el disco x con la cancion y
'si no devuelve ""
Public Function TemaDeSelecion(Disco As Long, Tema As Long, _
    Optional DevolverSoloNombre As Boolean = False) As String

    'tema es hasta 30!!!!
    If Tema > 30 Then
        'esta pidiendo giladas!
        TemaDeSelecion = ""
        Exit Function
    End If

    'disco es hasta mtotaldiscos!!!!
    If Disco > mTotalDiscos Then
        'esta pidiendo giladas!
        TemaDeSelecion = ""
        Exit Function
    End If

    If mMtxTemas(Disco, Tema) = "" Then
        TemaDeSelecion = ""
        Exit Function
    End If

    'si hay algo ver si es un tema!
    Dim Partes() As String
    Partes = Split(mMtxTemas(Disco, Tema), "*")

    Dim NombreTema As String
    NombreTema = Partes(2)

    If DevolverSoloNombre Then
        TemaDeSelecion = NombreTema
        Exit Function
    End If
    
    Dim sTema As String
    sTema = mMTXPathDiscos(Disco) + NombreTema

    If fso.FileExists(sTema) Then
        TemaDeSelecion = sTema
    Else
       TemaDeSelecion = ""
    End If

End Function

Public Function TemaNumpad(lDisco As Long, lCancion As Long) As String
    On Local Error GoTo erTNP
    TemaNumpad = mMtxTemas(lDisco, lCancion)
    Exit Function
erTNP:
    terr.AppendLog "TNP:" + CStr(lDisco) + "," + CStr(lCancion)
End Function

Public Function GetNumberDisk(PathDisc As String) As Long
    Dim A As Long
    '-1 si no lo encuentra
    GetNumberDisk = -1
    For A = 1 To UBound(mMTXDiscos)
        If LCase(mMTXPathDiscos(A)) = LCase(PathDisc) Then GetNumberDisk = A
    Next A
End Function

Public Function GetAnchoMarco() As Long
    GetAnchoMarco = zAnchoMarco
End Function
Public Function GetAltoMarco() As Long
    GetAltoMarco = zAltoMarco
End Function

'------------------------------------------------------
'TOUCHSCREEN
'------------------------------------------------------

Public Sub IniciarRegiones()
    Dim I As Long
    
    Dim rgnX1 As Long
    Dim rgnY1 As Long
    Dim rgnX2 As Long
    Dim rgnY2 As Long
    
    For I = 1 To UBound(mDisco)
        rgnX1 = zX + mDisco(I).X
        rgnY1 = zY + mDisco(I).Y
        rgnX2 = zX + mDisco(I).X + mDisco(I).GetMarcoTapaAncho
        rgnY2 = zY + mDisco(I).Y + mDisco(I).GetMarcoTapaAlto
        
        mDisco(I).IdRegion = CreateRectRgn(rgnX1, rgnY1, rgnX2, rgnY2)
    Next I
End Sub

Public Sub DoClicRegion(rgnX As Long, rgnY As Long)
    'solo si estoy viendo los discos   (agregar prop visible o algo asi SEGUIAQUI)
    
    If EstoyEn <> Portadas Then Exit Sub
    
    Dim I As Long
    
    For I = 1 To UBound(mDisco)
        If PtInRegion(mDisco(I).IdRegion, rgnX, rgnY) <> 0 Then
            Dim IND As Long
            IND = (mnPagActual - 1) * (mDiscosHorizontal * mDiscosVertical) + I - 1
            SeleccionarDisco IND 'aqui se carga mIndAbsoluto
            RaiseEvent ClickDisco(mIndiceAbsoluto, I)
            Exit Sub
        End If
    Next I
    
End Sub

Public Sub CargarPropiedades(obSGO As tbrSuperGuiObject)
        
    obSGO.X = zX
    obSGO.Y = zY
    obSGO.W = zAncho
    obSGO.H = zAlto
    'estirable ??? SEGUIRAQUI
    'alineaciones ???
    obSGO.SetProp "FontName", rFntName
    obSGO.SetProp "FontSize", CStr(rFntSize)
    obSGO.SetProp "FontBold", CStr(CLng(rFntBold))
    obSGO.SetProp "FontColor", CStr(rFntColor)
    obSGO.SetProp "FntColorContorno", CStr(rFntColorContorno)
    
    obSGO.SetProp "FontItalic", "0" 'no configurables por ahora
    obSGO.SetProp "FontStrike", "0"
    obSGO.SetProp "FontUnderLine", "0"
    
    obSGO.SetProp "ColorDiscoSel", CStr(mColorDiscoSel) 'agregue yo esa variable por que no estaba hecha
    obSGO.SetProp "alphaHabilitado", CStr(CLng(qUsarAlpha))
    
End Sub

'todas las clases del manu que tengan archivos externos
'deben devolver un empaquetado de las imagenes (predeterminado es solo la lista para que empaquete despues todo junto
Public Function GetFIMGs(Optional pathEmpaquetar As String = "") As clsFullImgs
    
    Dim F As New clsFullImgs, F2 As clsFullImgs_FILE
    
    Set F2 = F.AddFileByPath(PNG1, , "PNGMarcoDisco")
    F2.sID = "PNGMarcoDisco"
    F2.TypeFile = mImagePNG
    
    If pathEmpaquetar <> "" Then
        If fso.FileExists(pathEmpaquetar) Then fso.DeleteFile pathEmpaquetar, True
        F.SaveFile pathEmpaquetar
    End If
    
    Set GetFIMGs = F
    
End Function
