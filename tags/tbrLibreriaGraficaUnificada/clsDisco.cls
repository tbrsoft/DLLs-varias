VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiscoGrafico"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim stdPicAux_CI As StdPicture 'Variable Auxiliar para Cargar imagenes desde un archivo

Dim FuenteIniciada As Boolean

Dim qTitulo As String
Dim ListadeTemasManager As clsTemasManager

Dim MarcoTapa As RECT

Const OBJ_BITMAP = 7

'=========================================
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Private Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Private Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Private Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Private Const LR_LOADFROMFILE = 16  ' used with LoadImage

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

'=========================================
Dim hBitmapSprite As Long
Dim hBitmap As Long

Dim hbmImageOld As Long
'=========================================


Dim ShowPNG As New tbrPNG_3.tbrPNG3

Public AnchoSprite As Long
Public AltoSprite As Long
'========================
Public qX As Long
Public qY As Long
Public qAncho As Long
Public qAlto As Long
'========================

Dim qPNGPath As String

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim qAlgunFormulario As Form

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColor As Long
Dim rFntColorContorno As Long

'Dim DTxt As New pTextEx_Demo.clsMain
'Dim DTxtContorno As New pTextEx_Demo.clsMain

'------------------------------------------------------------------------------------------
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'------------------------------------------------------------------------------------------

Dim DestObjHdc As Long


Dim SpriteHDC As Long

Dim mAncho As Long 'Ancho y Alto del Marco
Dim mAlto As Long

'CHG99
Dim mVisible As Boolean

Private mNombre As String

'--------Titulo del Disco--------
Private TituloDiscoLabel As New clsLabel
Private qColorAlphaTit As Long
'Private qBoolAlphaTit As Boolean


Private maNumerik As New clsNumeriks
Private mUseMaNumerik As Boolean
Private mLastIndexDiscoDataCargado As Long

Public Sub DibujarDiscoGrafico()
    mostrarPNG qPNGPath, X, Y 'manu18
End Sub
'-------------------------------------
Public Sub DibujarDiscoData(ByVal DiscoData As clsDiscoData, Optional Seleccionado As Boolean = False, Optional ListaTemas As clsTemasManager)
    Dim pthimg As String
    pthimg = DiscoData.PathImagen
    mLastIndexDiscoDataCargado = DiscoData.Numero 'al hacerle un click necesito saber que disco es, por ahoa esto es lo mejor
    
    'mostrarPNG qPNGPath, X, Y

    
    'Cargo la nueva imagen
    CargarImagen pthimg
    
    
    tbrPintaNoPix.tbrPintarInHDC SpriteHDC, DestObjHdc, qX + MarcoTapa.qLeft, qY + MarcoTapa.qTop, MarcoTapa.qRight, MarcoTapa.qBottom, AnchoSprite, AltoSprite
    
    
    'Iniciar el texto del titulo
    IniciarTitulo vbBlack
    
    'CambiarColorTitulo vbRed
    FuenteIniciada = True
    
    'qTitulo = DiscoData.Titulo
    Dibujar_Titulo DiscoData.Titulo, Seleccionado
    
    'Dibujo los numeros y la lista
    If mUseMaNumerik = True Then
        Ini_maNumeriks
        maNumerik.SetNumero DiscoData.Numero, False, 3
        maNumerik.Renderizar
        
        Set ListadeTemasManager = ListaTemas
        Ini_TemasManager
        ListadeTemasManager.DibujarTexto -1
    End If
    

End Sub

Private Sub IniciarTitulo(qColor As Long)
    Dim fAncho As Long
    Dim fAlto As Long
    Dim mX As Long
    Dim mY As Long
    
    TituloDiscoLabel.CerrarGraficos
    GetMedidasTitulo fAncho, fAlto, mX, mY
    
    TituloDiscoLabel.IniciarGrafios DestObjHdc, mX, mY, fAncho, fAlto, True, qColor, 200
End Sub
'-------------------------------------
Public Function LastIndexDiscoDataCargado() As Long
    LastIndexDiscoDataCargado = mLastIndexDiscoDataCargado
End Function

'Sub DibujarTextContorno(Texto As String, HDC_Destino As Long, ttX As Long, ttY As Long, ttAncho As Long, ttAlto As Long)
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY - 1, ttAncho, ttAlto
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY + 1, ttAncho, ttAlto
'
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY, ttAncho, ttAlto
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY, ttAncho, ttAlto
'
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY - 1, ttAncho, ttAlto
'    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY + 1, ttAncho, ttAlto
'
'
'    DTxt.Dibujar Texto, HDC_Destino, ttX, ttY, ttAncho, ttAlto
'
'End Sub

Private Sub SetMedidas()
    Dim PNG_Token As Long
    
    If mUseMaNumerik = True Then
'Sin maNumeriks-------------------------------------
        If mpSYS.Name = martino_MP6 Then
            MarcoTapa.qLeft = 12
            MarcoTapa.qTop = 50
            MarcoTapa.qRight = 176
            MarcoTapa.qBottom = 139
        End If
        
        If mpSYS.Name = eCHUgames Then
            MarcoTapa.qLeft = 17
            MarcoTapa.qTop = 3
            MarcoTapa.qRight = 192
            MarcoTapa.qBottom = 159
        End If
        
        If mpSYS.Name = MPR Then
            MarcoTapa.qLeft = 12
            MarcoTapa.qTop = 50
            MarcoTapa.qRight = 176
            MarcoTapa.qBottom = 139
        End If
        
        'mAncho = 200
        'mAlto = 200
        'qAncho = mAncho
        'qAlto = mAlto
    Else
'Con maNumeriks-------------------------------------
        If mpSYS.Name = martino_MP6 Then
            MarcoTapa.qLeft = 12
            MarcoTapa.qTop = 50
            MarcoTapa.qRight = 176
            MarcoTapa.qBottom = 139
        End If
        
        If mpSYS.Name = eCHUgames Then
            MarcoTapa.qLeft = 15
            MarcoTapa.qTop = 4
            MarcoTapa.qRight = 230
            MarcoTapa.qBottom = 158
        End If
        
        If mpSYS.Name = MPR Then
            MarcoTapa.qLeft = 12
            MarcoTapa.qTop = 50
            MarcoTapa.qRight = 176
            MarcoTapa.qBottom = 139
        End If
    End If
End Sub

Public Sub IniciarPNG(PNGPath As String)
    Dim PNG_Token As Long
    
    qPNGPath = PNGPath

    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.GetAnchoAlto qPNGPath, mAncho, mAlto
    ShowPNG.FreeGDIPlus PNG_Token
End Sub
Public Sub IniciarGrafios(DondeObjetoHdc As Long, UseMaNumerik As Boolean, Optional ColorAlphaTit As Long = -1) ', XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    mUseMaNumerik = UseMaNumerik
    
    DestObjHdc = DondeObjetoHdc
    
    qColorAlphaTit = ColorAlphaTit
    
    SetMedidas
    
    
    'Crear el HDC para cargar las tapas de disco
    SpriteHDC = CreateCompatibleDC(0)
    
    Ini_maNumeriks
    
End Sub

Private Sub IniHDC(CualHDC As Long, qAncho2 As Long, qAlto2 As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho2, qAlto2)
    CualHDC = CreateCompatibleDC(0)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    DeleteObject TempBMP
    DeleteObject ObjCreado2
End Sub

Public Sub IniciarFuente(AlgunFormulario As Object, FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColor As Long, FntColorContorno As Long)
    Set qAlgunFormulario = AlgunFormulario
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    rFntColor = FontColor
    rFntColorContorno = FntColorContorno
    
    '        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColor
    'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
    
    TituloDiscoLabel.IniciarFuente qAlgunFormulario, rFntName, rFntSize, rFntBold, False, False, False, rFntColor, rFntColorContorno
    
    If mUseMaNumerik = True Then
        maNumerik.IniciarFuente qAlgunFormulario, "Transponder AOE", 20, False, False, False, False, vbWhite, RGB(20, 20, 20)
    End If
    
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

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
    maNumerik.CerrarGraficos
End Sub

'=================================================================================
Sub mostrarPNG(path As String, mX As Long, mY As Long, Optional sel As Boolean)
    Dim PNG_Token As Long
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.LoadPictureGDIPlus path, mX, mY, DestObjHdc
    ShowPNG.FreeGDIPlus PNG_Token
End Sub
'=================================================================================
Public Sub CargarImagen(Archivo As String)
    
    Dim ret As Long
    Dim AnchoDestino As Long
    Dim AltoDestino As Long
    
    Dim stdObjet As Long
    
    Dim lresult As Long
    
    Dim tmpName As String
    
    If Dir(Archivo) = "" Then Exit Sub
    PuedoMostrar = True
    
    Set stdPicAux_CI = LoadPicture(Archivo)
    lresult = SelectObject(SpriteHDC, stdPicAux_CI.Handle)
    
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

'=================================================================================
'=================================================================================
Public Property Get X() As Long
    X = qX
End Property
Public Property Let X(ByVal NV As Long)
    qX = NV
    'Ini_maNumeriks
End Property
Public Property Get Y() As Long
    Y = qY
End Property
Public Property Let Y(ByVal NV As Long)
    qY = NV
End Property
Public Property Get Ancho() As Long
    Ancho = qAncho
End Property
Public Property Let Ancho(ByVal NV As Long)
    qAncho = NV
End Property
Public Property Get Alto() As Long
    Alto = qAlto
End Property
Public Property Let Alto(ByVal NV As Long)
    qAlto = NV
End Property
'Public Property Get AnchoMarcoNumeriks() As Long
'    AnchoNumeriks = qAnchoMaNumerik
'End Property
'Public Property Let AnchoMarcoNumeriks(ByVal NV As Long)
'    qAnchoMaNumerik = NV
'End Property
'Public Property Get AltoMarcoNumeriks() As Long
'    AltoNumeriks = qAltoMaNumerik
'End Property
'Public Property Let AltoMarcoNumeriks(ByVal NV As Long)
'    qAltoMaNumerik = NV
'End Property


'=================================================================================
'=================================================================================

'CHG99
Public Property Let Visible(ByVal V As Boolean)
    mVisible = V
End Property

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Private Function ObtenerTextoConPuntitos(Texto As String, Ancho As Long) As String
    'Si se pasa del ancho dibuja ...
    Dim Ancho3p As Long
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim i As Long
    Dim TextoFor As String
    
    '1: Obtener el ancho de 3 puntitos
    GetAnchoAltoFont qAlgunFormulario, "...", rFntName, rFntSize, rFntBold, Ancho3p, AltoTexto
    
    '2 busco hasta pasarme del (ancho - Ancho de 3 puntitos)
    For i = 1 To Len(Texto)
        TextoFor = Mid(Texto, 1, i)
        GetAnchoAltoFont qAlgunFormulario, TextoFor, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        If AnchoTexto > (Ancho - Ancho3p) Then
            TextoFor = Mid(Texto, 1, i - 1) + "..."
            ObtenerTextoConPuntitos = TextoFor
            Exit Function
        End If
        
    Next i
    
    'Si llegue aca, por que no hacen falta los "..."
    ObtenerTextoConPuntitos = Texto
End Function

Public Function GetMarcoTapaX() As Long
    GetMarcoTapaX = MarcoTapa.qLeft
End Function

Public Function GetMarcoTapaY() As Long
    GetMarcoTapaY = MarcoTapa.qTop
End Function

Public Function GetMarcoTapaAncho() As Long
    GetMarcoTapaAncho = MarcoTapa.qRight
End Function

Public Function GetMarcoTapaAlto() As Long
    GetMarcoTapaAlto = MarcoTapa.qBottom
End Function
'-----------------------------------------
Public Function GetMarcoAncho() As Long
    GetMarcoAncho = mAncho
End Function
Public Function GetMarcoAlto() As Long
    GetMarcoAlto = mAlto
End Function



'Public Function GetUseMaNumerik() As Boolean
'    GetUseMaNumerik = mUseMaNumerik
'End Function
'
'Public Sub setUseMaNumerik(manuMer As Boolean)
'    mUseMaNumerik = manuMer
'End Sub

Private Sub Ini_maNumeriks()
    Dim XMaNumerik As Long
    Dim YMaNumerik As Long
    Dim AnchoMaNumerik As Long
    Dim AltoMaNumerik As Long
    
    If mUseMaNumerik = True Then
        AnchoMaNumerik = (mAncho * 0.18)
        AltoMaNumerik = (mAlto * 0.17)
        XMaNumerik = qX + 22
        YMaNumerik = qY + 6 '+ (mAlto - AltoMaNumerik) - 5
        maNumerik.CerrarGraficos
        
        maNumerik.IniciarGrafios DestObjHdc, XMaNumerik, YMaNumerik, AnchoMaNumerik, AltoMaNumerik, True, vbBlack, 200
    End If
End Sub

Private Sub Ini_TemasManager()
    Dim XLista As Long
    Dim YLista As Long
    Dim AnchoLista As Long
    Dim AltoLista As Long
    
    If mUseMaNumerik = True Then
        AnchoLista = (mAncho * 0.45) - 2
        AltoLista = (mAlto * 0.95) + 4
        XLista = qX + (mAncho * 0.55) - 1
        YLista = qY + 2
        ListadeTemasManager.CerrarGraficos
        ListadeTemasManager.IniciarGrafios DestObjHdc, XLista, YLista, AnchoLista, AltoLista, True, vbWhite
        ListadeTemasManager.IniciarFuente qAlgunFormulario, "Trebuchet MS", 7, False, False, False, False, RGB(220, 220, 220), vbWhite, RGB(10, 10, 10)
    End If
End Sub

Private Sub Dibujar_Titulo(tit As String, Seleccionado As Boolean)
    TituloDiscoLabel.SetTexto tit, Seleccionado
    TituloDiscoLabel.Renderizar
End Sub

Private Function GetMedidasTitulo(ByRef titAncho As Long, ByRef titAlto As Long, ByRef titX As Long, ByRef titY As Long)
    Dim fAncho As Long
    Dim fAlto As Long
    Dim mX As Long
    Dim mY As Long
        
'Ancho, Alto
    If mUseMaNumerik = True Then
        If mpSYS.Name = martino_MP6 Then
            fAncho = (mAncho * 0.85)
            fAlto = (mAlto * 0.762)
        End If
        
        If mpSYS.Name = eCHUgames Then
            fAncho = (mAncho * 0.49)
            fAlto = (mAlto * 0.1)
        End If
        
        If mpSYS.Name = MPR Then
            fAncho = (mAncho * 0.85)
            fAlto = (mAlto * 0.762)
        End If
    Else
        If mpSYS.Name = martino_MP6 Then
            fAncho = (mAncho * 0.8)
            fAlto = (mAlto * 0.762)
        End If
        
        If mpSYS.Name = eCHUgames Then
            fAncho = (mAncho * 0.91)
            fAlto = (mAlto * 0.12)
        End If
        
        If mpSYS.Name = MPR Then
            fAncho = (mAncho * 0.85)
            fAlto = (mAlto * 0.762)
        End If
    End If

'----------------------
'X, Y
    If mUseMaNumerik = True Then
        If mpSYS.Name = martino_MP6 Then
            mX = (mAncho / 2) - (AnchoFont / 2)
            mY = 20
        End If
        
        If mpSYS.Name = eCHUgames Then
            mX = ((mAncho * 0.5) / 2) - (fAncho / 2) + 16
            mY = (mAlto * 0.87)
            'mY = 145
        End If
        
        If mpSYS.Name = MPR Then
            mX = (mAncho / 2) - (AnchoFont / 2)
            mY = 20
        End If
    Else
        If mpSYS.Name = martino_MP6 Then
            mX = (mAncho / 2) - (AnchoFont / 2)
            mY = 20
        End If
        
        If mpSYS.Name = eCHUgames Then
            mX = ((mAncho * 1) / 2) - (fAncho / 2) + 5
            mY = (mAlto * 0.85)
            'mY = 145
        End If
        
        If mpSYS.Name = MPR Then
            mX = (mAncho / 2) - (AnchoFont / 2)
            mY = 20
        End If
    End If

titAncho = fAncho: titAlto = fAlto: titX = mX + qX: titY = mY + qY
End Function

'Public Function GetTitulo() As String
'End Function
Public Sub CambiarColorTitulo(qColor As Long)
    TituloDiscoLabel.CambiarColor qColor
    TituloDiscoLabel.Renderizar
End Sub
