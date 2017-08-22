VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal HDC As Long, ByVal hStretchMode As Long) As Long
Const STRETCHMODE = vbPaletteModeNone

Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

'------------------------------------------------------------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long

Dim qInAncho As Long
Dim qInAlto As Long
Dim qOutX As Long
Dim qOutY As Long
Dim qOutAncho As Long
Dim qOutAlto As Long

Dim SpriteHDC As Long
Dim DestObjHdc As Long


Public Sub Renderizar()
    tbrPintarInHDC SpriteHDC, DestObjHdc, qOutX, qOutY, qOutAncho, qOutAlto, qInAncho, qInAlto
End Sub

Public Sub IniciarGraficos(DestinoHDC As Long, InAncho As Long, InAlto As Long, OutX As Long, OutY As Long, OutAncho As Long, OutAlto As Long)
    qInAncho = InAncho
    qInAlto = InAlto
    qOutX = OutX
    qOutY = OutY
    qOutAncho = OutAncho
    qOutAlto = OutAlto

    DestObjHdc = DestinoHDC
    
    IniHDC SpriteHDC, qInAncho, qInAlto

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


Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
End Sub

Private Sub tbrPintarInHDC(HdcIn As Long, HdcOut As Long, X As Long, Y As Long, AnchoOut As Long, AltoOut As Long, AnchoIn As Long, AltoIn As Long)
    Call SetStretchBltMode(HdcOut, STRETCHMODE)
    Call StretchBlt(HdcOut, X, Y, AnchoOut, AltoOut, HdcIn, 0, 0, AnchoIn, AltoIn, vbSrcCopy)
End Sub


'===============================================================
'Propiedades solo lectura
'===============================================================

Public Property Get HDC() As Long
    HDC = SpriteHDC
End Property

Public Property Get AnchoIn() As Long
    AnchoIn = qInAncho
End Property

Public Property Get AltoIn() As Long
    AltoIn = qInAncho
End Property

Public Property Get Xout() As Long
    Xout = qOutX
End Property

Public Property Get Yout() As Long
    Yout = qOutY
End Property

Public Property Get AnchoOut() As Long
    AnchoOut = qOutAncho
End Property

Public Property Get AltoOut() As Long
    AltoOut = qOutAlto
End Property
