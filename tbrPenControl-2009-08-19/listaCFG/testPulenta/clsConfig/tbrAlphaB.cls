VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrAlphaB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function Rectangle Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Const AC_SRC_OVER = &H0

Private Type AlphaOptions
  AlphaOption As Byte
  AlphaFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal HDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal HDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long

'******************************************************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim SpriteHDC As Long
Dim hbmImageOld As Long

Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Private Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Private Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Private Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Private Const LR_LOADFROMFILE = 16  ' used with LoadImage

Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Public Ancho As Long
Public Alto As Long

Public Sub AbrirImagen(Archivo As String)
    Dim hBitmap As Long
    Dim BitmapData As BITMAP    ' data on the incoming bitmap
    Dim lresult As Long
    
    Dim tmpName As String
    
    If Dir(Archivo) = "" Then Exit Sub
    '----------extra-----------
    tmpName = "C:\bmptemp.spr"
    JpgtoBmp Archivo, tmpName
    Archivo = tmpName
    '--------------------------
    
    
    hBitmap = LoadImage(0, Archivo, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    lresult = GetObject(hBitmap, Len(BitmapData), BitmapData)
    Ancho = BitmapData.bmWidth     ' Determine the bitmaps width
    Alto = BitmapData.bmHeight    ' Determine the bitmaps height
    SpriteHDC = CreateCompatibleDC(0)
    hbmImageOld = SelectObject(SpriteHDC, hBitmap)
End Sub

Private Sub JpgtoBmp(JpgPath As String, BmpPath As String)
    Dim x As StdPicture
    Set x = LoadPicture(JpgPath)
    If Dir(BmpPath) <> "" Then Kill (BmpPath)
    Call SavePicture(x, BmpPath)
End Sub

Private Sub Class_Terminate()
    DeleteObject (hbmImageOld)
    DeleteDC (SpriteHDC)
End Sub

Public Sub MostrarSimple(qHDC As Long, qX As Long, qY As Long)
    BitBlt qHDC, qX, qY, Ancho, Alto, SpriteHDC, 0, 0, vbSrcCopy
End Sub
'******************************************************************

Public Sub MostrarAlpha(qHDC As Long, qX As Long, qY As Long, qAlpha254 As Long, Optional qAncho As Long = -1, Optional qAlto As Long = -1)
    Dim AO As AlphaOptions
    Dim ptAO As Long
    
    If qAncho = -1 Then
        qAncho = Ancho
    End If
    If qAlto = -1 Then
        qAlto = Alto
    End If
    
    With AO
        .AlphaOption = AC_SRC_OVER
        .AlphaFlags = 0
        .SourceConstantAlpha = qAlpha254
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory ptAO, AO, 4
    AlphaBlend qHDC, qX, qY, qAncho, qAlto, SpriteHDC, 0, 0, Ancho, Alto, ptAO
End Sub

Public Sub CrearHDC_Color(qColor As Long)
    Dim qRect As RECT
    Dim hBitmap As Long
    
    SpriteHDC = CreateCompatibleDC(0)
    
    qRect.qLeft = 0
    qRect.qTop = 0
    qRect.qBottom = 10
    qRect.qRight = 10
    hBitmap = CreateCompatibleBitmap(SpriteHDC, 10, 10)
    Ancho = 10
    Alto = 10
    
    hbmImageOld = SelectObject(SpriteHDC, hBitmap)
    SetBkColor SpriteHDC, qColor
    Rectangle SpriteHDC, -1, -1, 11, 11
     
    DeleteObject hBitmap
    DeleteObject hbmImageOld
End Sub

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
End Sub
