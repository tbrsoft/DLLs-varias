VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================================
'Aclaracion sobre el codigo:
'cCda elemento de la lista tiene como tag el index del nodo que debe cargar,
'en el caso de volver tiene como tag el index del padre, si no hay mas padres
'(el nivel mas alto) el tag es "x"
'====================================================================================
'Dim qLabelTouch As Label

Private RectaBotones As New clsRectangulo
Private RectaLista As clsRectangulo
Private RectaTitulo As clsRectangulo
Private RectaTitulo2 As clsRectangulo


Dim AnchoSprite As Long
Dim AltoSprite As Long

Dim qKeyRaiz As String

Dim EstoyEnBotonera As Boolean
Dim IndexBoton As Long
Dim IndexItemLista As Long

Dim ShowPNG As New tbrPNG_3.tbrPNG3
Dim PrimerSeleccion As Boolean
Dim LastBotonSelect As clsPNGBoton
Dim qPNG_Select As String
Dim PNGSelectHDC As Long

Dim LabelTitulo As New clsLabel
Dim LabelExpl As New clsLabel

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColor As Long
Dim rFntColorContorno As Long

'Dim lblTouchBt As Label
Dim ListaConfigElem As New clsTemaList
Dim ListaConfig As New clsTemasManager

Dim qPathImagenes As String

Dim LosBotones() As clsPNGBoton
Dim BotonX As Long
Dim BotonY As Long

Dim qAlgunFormulario As Form

Dim FondoHDC As Long
Dim FondoHDC_BUP As Long
Dim FondoHDC_TMP As Long


Public Nodos As Nodes

'------------------------------------------------
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Private Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Private Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Private Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Private Const LR_LOADFROMFILE = 16  ' used with LoadImage


Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim qRect As RECT

Dim DestObjHdc As Long
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long

Public Sub IniciarNodos(TV As TreeView, KeyRaiz As String)
    Set Nodos = TV.Nodes
    qKeyRaiz = KeyRaiz
End Sub

Public Sub IniciarRectangulos(lRectaBotones As clsRectangulo, lRectaLista As clsRectangulo, lRectaTitulo As clsRectangulo, lRectaTitulo2 As clsRectangulo)
    Set RectaBotones = lRectaBotones
    Set RectaLista = lRectaLista
    Set RectaTitulo = lRectaTitulo
    Set RectaTitulo2 = lRectaTitulo2

End Sub

Public Sub IniciarGrafios(DondeObjetoHdc As Long, _
    XFondo As Long, YFondo As Long, WFondo As Long, HFondo As Long, _
    ArchivoFondo As String, PNG_Select As String, PathImagenes As String) ', LabelTouch As Label)
    
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    qPathImagenes = PathImagenes
    qPNG_Select = PNG_Select
    'Set qLabelTouch = LabelTouch
    
    zX = XFondo   'X e Y se pueden modificar pero ....
    zY = YFondo
    zAncho = WFondo '640  ' ¿ el ancho y el alto no ?
    zAlto = HFondo '480   ' maldito-manu me cago en tu idea de "modificable"
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, zAncho, zAlto
    IniHDC FondoHDC_TMP, zAncho, zAlto
    
    'Cargo la imagen de fondo
    CargarImagenEnHDC ArchivoFondo, FondoHDC_TMP
    
    'pongo la imagen de fondo en FondoHDC
    tbrPintarInHDC FondoHDC_TMP, FondoHDC, 0, 0, zAncho, zAlto, AnchoSprite, AltoSprite
    
    '(backup de FondoHDC para rederizar)
    IniHDC FondoHDC_BUP, zAncho, zAlto
    BitBlt FondoHDC_BUP, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
    IniHDC PNGSelectHDC, 82, 82
    
    
    
    'Andres. Podes poner los altos y anchos proporcionales?
    '=====================================================
    'Medidas Objetos dentro del clsConfig
    '=====================================================
    Dim tmpX As Long
    Dim tmpY As Long
    Dim tmpAncho As Long
    Dim tmpAlto As Long
    
    'Medidas Botonera
    BotonX = RectaBotones.X
    BotonY = RectaBotones.Y
    'BotonX = 55: BotonY = 30
    
    
    'Medidas Lista de Configuraciones
    tmpX = RectaLista.X
    tmpY = RectaLista.Y
    tmpAncho = RectaLista.Ancho
    tmpAlto = RectaLista.Alto
    ListaConfig.IniciarGrafios FondoHDC, tmpX, tmpY, tmpAncho, tmpAlto, False, vbWhite
    'ListaConfig.IniciarGrafios FondoHDC, 56, 145, 320, 235, False, vbWhite
    
    'Medidas del Titulo
    tmpX = RectaTitulo.X
    tmpY = RectaTitulo.Y
    tmpAncho = RectaTitulo.Ancho
    tmpAlto = RectaTitulo.Alto
    LabelTitulo.IniciarGrafios FondoHDC, tmpX, tmpY, tmpAncho, tmpAlto, False, vbBlack, 150
    'LabelTitulo.IniciarGrafios FondoHDC, 55, 112, 320, 28, False, vbBlack, 150
    
    'Medidas del Label de explicacion
    tmpX = RectaTitulo2.X
    tmpY = RectaTitulo2.Y
    tmpAncho = RectaTitulo2.Ancho
    tmpAlto = RectaTitulo2.Alto
    LabelExpl.IniciarGrafios FondoHDC, tmpX, tmpY, tmpAncho, tmpAlto, False, vbBlack, 150
    
    '=====================================================
    
    
    ListaConfig.IniciarTouchScreen
    SetExplicacion "Observaciones..."
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

'----------------------------------------------
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

Public Sub IniciarFuente(AlgunFormulario As Form, FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColor As Long, FntColorContorno As Long)
    Set qAlgunFormulario = AlgunFormulario
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    rFntColor = FontColor
    rFntColorContorno = FntColorContorno
    
    ListaConfig.IniciarFuente qAlgunFormulario, rFntName, rFntSize, rFntBold, False, False, False, rFntColor, RGB(30, 30, 40), rFntColorContorno
    LabelTitulo.IniciarFuente qAlgunFormulario, rFntName, rFntSize, rFntBold, False, False, False, rFntColor, rFntColorContorno
    LabelExpl.IniciarFuente qAlgunFormulario, rFntName, 8, rFntBold, False, False, False, rFntColor, rFntColorContorno
End Sub

Public Sub CargarNodos()
    Dim btAux As clsPNGBoton
    Dim aux As String
    Dim txts() As String
    Dim i As Long
    For i = 1 To Nodos.Count
        aux = Nodos(i).Text
        
        Dim FilPNG As String
        FilPNG = qPathImagenes + Nodos(i).Key + ".png"
        
        If Dir(FilPNG) <> "" Then
            Set btAux = AgregarBoton(CStr(i))
            btAux.IniciarPNGs FilPNG
            
            btAux.IniciarGrafios FondoHDC, BotonX, BotonY ', qLabelTouch
            
            BotonX = BotonX + btAux.GetAncho + 10
            btAux.DibujarBoton
            btAux.Renderizar
        End If
    Next i
End Sub

Public Sub CerrarGraficos()
    Dim i As Long
    Dim ret As Long
    ret = DeleteDC(FondoHDC)
    ret = DeleteDC(FondoHDC_BUP)

    For i = 1 To UBound(LosBotones)
        LosBotones(i).CerrarGraficos
    Next i
    
    LabelTitulo.CerrarGraficos
    LabelExpl.CerrarGraficos
    ListaConfig.CerrarGraficos
End Sub

Private Sub CargarImagenEnHDC(Archivo As String, CualHDC As Long)
    
    Dim ret As Long
    
    Dim stdObjet As Long
    
    Dim lresult As Long
    
    Dim tmpName As String
    
    If Dir(Archivo) = "" Then Exit Sub
    PuedoMostrar = True
    
    Set stdPicAux_CI = LoadPicture(Archivo)
    lresult = SelectObject(CualHDC, stdPicAux_CI.Handle)
    
    AnchoSprite = qAlgunFormulario.ScaleX(stdPicAux_CI.Width, vbHimetric, vbPixels)
    AltoSprite = qAlgunFormulario.ScaleX(stdPicAux_CI.Height, vbHimetric, vbPixels)
    
    'ElDestinoHDC = DestinoHDC
    
    If AnchoSprite < 1 Then Ancho = 1
    If AnchoSprite < 1 Then Alto = 1
    
'DESCARGAR MEMORIA:--------------------------
    ret = DeleteObject(lresult)
    Set stdPicAux_CI = Nothing
'--------------------------------------------
End Sub

Public Sub Renderizar()
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
End Sub

'---------------------
Private Function AgregarBoton(cTag As String) As clsPNGBoton  'Devuelvo index agregado
    Dim Ix As Long
    Ix = UBound(LosBotones) + 1
    ReDim Preserve LosBotones(Ix)
    Set LosBotones(Ix) = New clsPNGBoton
    Set AgregarBoton = LosBotones(Ix)
    'El Tag es necesario para asociar el boton al nodo arre :$
    LosBotones(Ix).TagBoton = cTag
End Function

Private Sub QuitarBoton(index As Long)
    If UBound(LosBotones) = 0 Then
        ReDim LosBotones(0)
        Exit Sub
    End If
    Dim i As Long
    For i = index To UBound(LosBotones) - 1
        LosBotones(i) = LosBotones(i + 1)
    Next i
    ReDim Preserve LosBotones(UBound(LosBotones) - 1)
End Sub

'---------------------
Private Sub Class_Initialize()
    ReDim LosBotones(0)
    'LastBotonSelect = -1
    EstoyEnBotonera = True
End Sub

Private Sub CargarHijosDeNodo(qNodo As Node, Optional OmitirPadre As Boolean = False)
    Dim LaKey As String
    Dim aux As String
    
    Dim i As Long
    Dim i2 As Long
    
    i2 = 0
    'Me fijo el nombre del Nodo al que tengo que buscarle los hijos
    LaKey = qNodo.Key

    'Dibujo el titulo ;)
    aux = qNodo.Text
    LabelTitulo.SetTexto aux
    
    
    
    'Agregar el elmento volver===================================
    Set ListaConfigElem = ListaConfig.AgregarElemento()
    i2 = i2 + 1
    ListaConfigElem.Titulo = "Volver"
    'ListaConfigElem.Numero = i2
    ListaConfigElem.Numero = -1
    ListaConfigElem.Duracion = " "
    'En el caso del tag de "Volver" cargo el index del nodo padre
    If OmitirPadre = False Then
        ListaConfigElem.TagMisterioso = CStr(qNodo.Parent.index)
    Else
        ListaConfigElem.TagMisterioso = "x"
    End If
    '============================================================
    
    'Hago una vuelta por TODOS los nodos
    For i = 1 To Nodos.Count
        
        'Es este nodo hijo de la raiz? por que de ser así
        'me va a dar un error cuando pida .Parent
        'solo voy a tener en cuenta nodos secundarios
        If Nodos(i).Key <> Nodos(i).Root.Key Then
            
            'Es este nodo hijo del que busco?
            If Nodos(i).Parent.Key = LaKey Then
                aux = Nodos(i).Text
                
                Set ListaConfigElem = ListaConfig.AgregarElemento()
                ListaConfigElem.Titulo = aux
                i2 = i2 + 1
                'ListaConfigElem.Numero = i2
                ListaConfigElem.Numero = -1
                ListaConfigElem.Duracion = " "
                'Guardo el index del nodo correspondiente...
                ListaConfigElem.TagMisterioso = CStr(i)
            End If
        End If
    Next i
End Sub

Private Sub MostrarNodo(qNodo As Node, Optional OmitirPadre As Boolean = False)
    ListaConfig.ResetElementos
    CargarHijosDeNodo qNodo, OmitirPadre
    ListaConfig.DibujarTexto 0
    LabelTitulo.Renderizar
End Sub

Private Sub SetExplicacion(Explic As String)
    LabelExpl.SetTexto Explic
    LabelExpl.Renderizar
End Sub

Private Sub SelectBotonPNG(index)
    Dim bX As Long
    Dim bY As Long
    
    bX = LosBotones(index).GetX - 10
    bY = LosBotones(index).GetY - 10
    
    'Borro la seleccion del boton anteriormente seleccionadio
    UselBotonPNG
    
    BitBlt PNGSelectHDC, 0, 0, 82, 82, FondoHDC, bX, bY, SRCCOPY
    mostrarPNG qPNG_Select, bX, bY
    Set LastBotonSelect = LosBotones(index)
    PrimerSeleccion = True
End Sub

Private Sub UselBotonPNG()
    If PrimerSeleccion = True Then
        BitBlt FondoHDC, LastBotonSelect.GetX - 10, LastBotonSelect.GetY - 10, 82, 82, PNGSelectHDC, 0, 0, SRCCOPY
    End If
End Sub

Private Sub mostrarPNG(path As String, mX As Long, mY As Long, Optional sel As Boolean)
    Dim PNG_Token As Long
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.LoadPictureGDIPlus path, mX, mY, FondoHDC
    ShowPNG.FreeGDIPlus PNG_Token
End Sub

'--------------------------------------------------------------------------
Public Sub ComandoAdelante()
    RealComando 1
End Sub

Public Sub ComandoAtras()
    RealComando 0
End Sub

Public Sub ComandoEntrar()
    If EstoyEnBotonera = True Then
        UselBotonPNG
        EstoyEnBotonera = False
        IndexItemLista = 1
        ListaConfig.DibujarTexto IndexItemLista
    Else
        Set ListaConfigElem = ListaConfig.GetElem(IndexItemLista)
        
        'Si hizo click en volver...
        If EstoyEnVolver = True Then
            'Y el padre es el nivel mas alto (los botones de arriba)...
            If Nodos(val(ListaConfigElem.TagMisterioso)).Key = qKeyRaiz Then
                '...vuelvo a la botonera.
                IndexItemLista = 1
                ListaConfig.DibujarTexto 0
                EstoyEnBotonera = True
                SelectBotonPNG IndexBoton
                Exit Sub
            End If
        End If
        
        'Hizo click un item (puede ser volver u otro)
        'cargo los items solo si confime que tiene hijos (segun el TagMisterioso)
        If Nodos(val(ListaConfigElem.TagMisterioso)).Children > 0 Then
            MostrarNodo Nodos(val(ListaConfigElem.TagMisterioso))
            IndexItemLista = 1
            ListaConfig.DibujarTexto IndexItemLista
        End If
    End If
End Sub

Private Sub RealComando(Direccion As Long)
    'Estoy en la Botoner?
    If EstoyEnBotonera = True Then
        If Direccion = 1 Then
            IndexBoton = IndexBoton + 1
            If IndexBoton > UBound(LosBotones) Then
                IndexBoton = 1
            End If
        Else
            IndexBoton = IndexBoton - 1
            If IndexBoton < 1 Then
                IndexBoton = UBound(LosBotones)
            End If
        End If
        SelectBotonPNG IndexBoton
        MostrarNodo Nodos(val(LosBotones(IndexBoton).TagBoton))
    
    Else
    'No estoy en la botonera
        If Direccion = 1 Then
            IndexItemLista = IndexItemLista + 1
            If IndexItemLista > ListaConfig.GetElementoCount Then
                IndexItemLista = 1
            End If
        Else
            IndexItemLista = IndexItemLista - 1
            If IndexItemLista < 1 Then
                IndexItemLista = ListaConfig.GetElementoCount
            End If
        End If
        ListaConfig.DibujarTexto IndexItemLista
    End If
End Sub

Public Function SelectedItem() As Node
    
    If EstoyEnBotonera = True Then
        If IndexBoton > 0 Then
            Set SelectedItem = Nodos(val(LosBotones(IndexBoton).TagBoton))
            terr.AppendSinHist "SelIT2:" + Nodos(val(LosBotones(IndexBoton).TagBoton)).Text
        End If
    Else
        If ListaConfigElem.TagMisterioso <> "" Then
            Set SelectedItem = Nodos(val(ListaConfigElem.TagMisterioso))
            terr.AppendSinHist "SelIT:" + Nodos(val(ListaConfigElem.TagMisterioso)).Text
        End If
    End If
End Function

Public Function EstoyEnVolver() As Boolean
        'en que item de ListaConfig estoy?
        Set ListaConfigElem = ListaConfig.GetElem(IndexItemLista)
        
        'si estoy en la botonera no puedo estar en volver XD
        If EstoyEnBotonera = True Then
            EstoyEnVolver = False
        Else
            If ListaConfigElem.Titulo = "Volver" Then
                EstoyEnVolver = True
            Else
                EstoyEnVolver = False
            End If
        End If
End Function

Public Sub DoClick(clX As Long, clY As Long)
    Dim i As Long
    'Repaso todas las regiones para ver si se hizo click en algun elemento

'BOTONES---------------------------------------
    For i = 1 To UBound(LosBotones)
        If PtInRegion(LosBotones(i).IdRegion, clX, clY) <> 0 Then
            EstoyEnBotonera = True
            IndexBoton = i
            SelectBotonPNG i
            MostrarNodo Nodos(val(LosBotones(i).TagBoton))
            
            'IndexItemLista = 1
            'ListaConfig.DibujarTexto IndexItemLista
            Exit Sub
        End If
    Next i
'ListaConfig--------------------------------------
    i = ListaConfig.DoClick_GetElementoIndex(clX, clY)
    If i > -1 Then
        EstoyEnBotonera = False
        If i <= ListaConfig.GetElementoCount Then
            IndexItemLista = i
            Set ListaConfigElem = ListaConfig.GetElem(IndexItemLista)
            'MsgBox ListaConfigElem.Titulo
        End If
        Exit Sub
    End If
End Sub
