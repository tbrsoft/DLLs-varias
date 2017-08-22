VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrTextoMovil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase es un label que se mueve cuando el texto es muy laro usando "dibujartexto" con algun timer


Private ElTexto As String

Dim qPixelSalteo As Long

Dim XActual As Long

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean

Dim qAlgunFormulario As Form
Dim DTxt As New pTextEx_Demo.clsMain

Dim qRect As RECT

Dim DestObjHdc As Long
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long

Dim SpriteHDC As Long
Dim FondoHDC As Long

'Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Sub IniciarGraficos(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long, PixelSalteo As Long)

    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    zX = XFondo
    zY = YFondo
    zAncho = AnchoFondo
    zAlto = AltoFondo
    
    'Creo dos HDC, en uno copio el fondo (lugar donde voy a imprimir)
    'el otro lo utilizo cada vez que imprimo algo, pego el fondo e imprimo
    'el texto en el x correspondiente
    
    DestObjHdc = DondeObjetoHdc
    qPixelSalteo = PixelSalteo
    
    SpriteHDC = CreateCompatibleDC(0)
    IniHDC SpriteHDC, AnchoFondo, AltoFondo
    
    FondoHDC = CreateCompatibleDC(0)
    IniHDC FondoHDC, AnchoFondo, AltoFondo
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
End Sub

Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(DestObjHdc)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    DeleteObject ObjCreado2
End Sub

Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColor As Long)
    DTxt.Iniciar FontName, FontSize, FontBold, FontItalic, FontStrikethrough, FontUnderLine, FontColor
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold

End Sub

Public Sub DibujarTexto()
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    XActual = XActual - qPixelSalteo
    If XActual < 0 - AnchoTexto Then
        XActual = zAncho
    End If
    
    qRect.qLeft = XActual
    qRect.qTop = 0
    qRect.qRight = zAncho
    qRect.qBottom = zAlto
    
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    DTxt.Dibujar ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
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

Public Sub SetTexto(qTexto As String)
    ElTexto = qTexto
    XActual = zAncho
End Sub

