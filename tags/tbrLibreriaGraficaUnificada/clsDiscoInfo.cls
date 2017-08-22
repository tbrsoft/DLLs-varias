VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiscoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim qFormularioAuxiliar As Form
Dim myAlpha As New tbrAlphaB

Public AnchoSprite As Long
Public AltoSprite As Long

Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Private Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Private Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Private Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Private Const LR_LOADFROMFILE = 16  ' used with LoadImage


Dim SpriteHDC As Long


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

Dim FondoHDC As Long
Dim FondoHDC_BUP As Long

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

'------------------------------------------------------------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


Public Sub Renderizar()
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
End Sub

Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long, FormularioAuxiliar As Object)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    zX = XFondo
    zY = YFondo
    zAncho = AnchoFondo
    zAlto = AltoFondo
    
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, zAncho, zAlto
    IniHDC FondoHDC_BUP, zAncho, zAlto
    
    IniHDC SpriteHDC, zAncho, zAlto
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, DestObjHdc, XFondo, YFondo, SRCCOPY
    BitBlt FondoHDC_BUP, 0, 0, zAncho, zAlto, DestObjHdc, XFondo, YFondo, SRCCOPY


    'con estas dos lineas es como que creo un marco de semitransparencia
    'negra para esta clase... como un fondo ;)
    myAlpha.CrearHDC_Color vbBlack
    myAlpha.MostrarAlpha FondoHDC_BUP, 0, 0, 45, zAncho, zAlto
    myAlpha.CrearHDC_Color vbBlack
    myAlpha.MostrarAlpha FondoHDC_BUP, 2, 2, 45, zAncho - 4, zAlto - 4
    myAlpha.CerrarGraficos
    
    Set qFormularioAuxiliar = FormularioAuxiliar
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


Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(FondoHDC)
    ret = DeleteDC(FondoHDC_BUP)
    myAlpha.CerrarGraficos
End Sub

Private Sub RestaurarFondo()
    'Subo el Back Up del Fondo
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, FondoHDC_BUP, 0, 0, SRCCOPY
End Sub
Public Sub CargarImagen(Archivo As String)
    
    Dim ret As Long
    Dim AnchoDestino As Long
    Dim AltoDestino As Long
    
    Dim stdPicAux As StdPicture
    Dim stdObjet As Long
    
    Dim BitmapData As BITMAP    ' data on the incoming bitmap
    Dim lresult As Long
    
    Dim aux_lenPixelIn As Long
    
    Dim tmpName As String
    
    If Dir(Archivo) = "" Then Exit Sub
    PuedoMostrar = True
    
    'Crear el HDC
    SpriteHDC = CreateCompatibleDC(0)
    
    Set stdPicAux = LoadPicture(Archivo)
    lresult = SelectObject(SpriteHDC, stdPicAux.Handle)
    
    AnchoSprite = qFormularioAuxiliar.ScaleX(stdPicAux.Width, vbHimetric, vbPixels)
    AltoSprite = qFormularioAuxiliar.ScaleX(stdPicAux.Height, vbHimetric, vbPixels)
    
    'ElDestinoHDC = DestinoHDC
    
    If AnchoSprite < 1 Then Ancho = 1
    If AnchoSprite < 1 Then Alto = 1
    
'DESCARGAR MEMORIA:--------------------------
    ret = DeleteObject(lresult)
'--------------------------------------------
End Sub

'------------------------------------------------------------------
Public Sub SetDiscoInfo(PathTapa As String)
    Dim AnchoFont As Long
    Dim AltoFont As Long
    Dim TxtPunt As String
    
    RestaurarFondo
    
    'Dibujar la tapa
    CargarImagen PathTapa
    tbrPintaNoPix.tbrPintarInHDC SpriteHDC, FondoHDC, 5, 5, (zAncho - 10), (zAlto - 10), AnchoSprite, AltoSprite
    
End Sub

