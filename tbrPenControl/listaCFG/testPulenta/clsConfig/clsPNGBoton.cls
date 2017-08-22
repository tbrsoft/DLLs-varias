VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPNGBoton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public IdRegion As Long
Public TagBoton As String
Private ClickHabilitado As Boolean

'Dim qLabelTouch As Label

Dim ShowPNG As New tbrPNG_3.tbrPNG3

Public AnchoSprite As Long
Public AltoSprite As Long

Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Private Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Private Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Private Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Private Const LR_LOADFROMFILE = 16  ' used with LoadImage


'Dim SpriteHDC As Long


Private Declare Function GetCurrentObject Lib "gdi32" (ByVal HDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

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

Dim qPNGpathUnSel As String
Dim qPNGpathSel As String

'------------------------------------------------------------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long

Public Sub Renderizar()
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
End Sub

Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long) ', LabelTouch As Label) ', PNGpathUnSel As String, PNGpathSel As String)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    'Set qLabelTouch = LabelTouch
    
    zX = XFondo
    zY = YFondo
    
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, zAncho, zAlto
    IniHDC FondoHDC_BUP, zAncho, zAlto
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, DestObjHdc, zX, zY, SRCCOPY
    BitBlt FondoHDC_BUP, 0, 0, zAncho, zAlto, DestObjHdc, zX, zY, SRCCOPY

    'Pongo el label en su lugar
    'qLabelTouch.BackStyle = 0
    
    'qLabelTouch.Left = GetX * 15
    'qLabelTouch.Top = GetY * 15
    'qLabelTouch.Width = GetAncho * 15
    'qLabelTouch.Height = GetAlto * 15
    
    'Creo la region para los clic
    'IdRegion = CreateRectRgn(GetX * 15, GetY * 15, GetAncho * 15, GetAlto * 15)

    IdRegion = CreateRectRgn(GetX, GetY, GetX + GetAncho, GetY + GetAlto)
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
End Sub

Private Sub RestaurarFondo()
    'Subo el Back Up del Fondo
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, FondoHDC_BUP, 0, 0, SRCCOPY
    
    'qLabelTouch.Visible = False
    ClickHabilitado = False
End Sub

Public Sub RestaurarFondoHDC()
    'Subo el Back Up del Fondo
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC_BUP, 0, 0, SRCCOPY
    
    'qLabelTouch.Visible = False
    ClickHabilitado = False
End Sub

'------------------------------------------------------------------
Private Sub mostrarPNG(path As String, mX As Long, mY As Long, Optional sel As Boolean)
    Dim PNG_Token As Long
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.LoadPictureGDIPlus path, 0, 0, FondoHDC
    'ShowPNG.LoadPictureGDIPlus path, mX, mY, FondoHDC
    ShowPNG.FreeGDIPlus PNG_Token

    'zAncho = ShowPNG.GetAncho
    'zAlto = ShowPNG.GetAlto
End Sub

Public Sub DibujarBoton()
    RestaurarFondo
    mostrarPNG qPNGpathUnSel, zX, zY
    
    'qLabelTouch.Visible = True
    ClickHabilitado = True
End Sub
'--------------------------------------------
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

Sub IniciarPNGs(PNGpathUnSel As String) ', PNGpathSel As String)
    Dim PNG_Token As Long
    qPNGpathUnSel = PNGpathUnSel
    'qPNGpathSel = PNGpathSel
    
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.GetAnchoAlto qPNGpathUnSel, zAncho, zAlto
    ShowPNG.FreeGDIPlus PNG_Token
End Sub
