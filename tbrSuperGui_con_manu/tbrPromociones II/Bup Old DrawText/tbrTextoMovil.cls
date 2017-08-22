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
Attribute VB_Exposed = False
Const DT_WORDBREAK = &H10

Public ElTexto As String


Dim qRect As RECT

Dim DestObj As Object
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long

Dim SpriteHDC As Long
Dim FondoHDC As Long

'Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Sub Iniciar(DondeObjeto As Object, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    zX = XFondo
    zY = YFondo
    zAncho = AnchoFondo
    zAlto = AltoFondo
    
    'Creo dos HDC, en uno copio el fondo (lugar donde voy a imprimir)
    'el otro lo utilizo cada vez que imprimo algo, pego el fondo e imprimo
    'el texto en el x correspondiente
    
    Set DestObj = DondeObjeto
    
    SpriteHDC = CreateCompatibleDC(0)
    IniHDC SpriteHDC, AnchoFondo, AltoFondo
    
    FondoHDC = CreateCompatibleDC(0)
    IniHDC FondoHDC, AnchoFondo, AltoFondo
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObj.hdc, XFondo, YFondo, SRCCOPY
End Sub
 
Public Sub DibujarTexto(PorCiento As Long)
    Dim XDestino As Long
    
    qRect.qLeft = PorCiento
    qRect.qTop = 0
    qRect.qRight = zAncho
    qRect.qBottom = zAlto
    
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    DrawText SpriteHDC, ElTexto, Len(ElTexto), qRect, DT_WORDBREAK
    BitBlt DestObj.hdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
    
End Sub
 
Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObj.hdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(DestObj.hdc)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    DeleteObject ObjCreado2
End Sub

