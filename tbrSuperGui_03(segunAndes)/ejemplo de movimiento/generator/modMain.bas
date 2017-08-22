Attribute VB_Name = "Globales"
Public qFormularioAuxiliar As Form
Public Type RECT
  qLeft As Long
  qTop As Long
  qRight As Long
  qBottom As Long
End Type

'========================================================================
Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Public Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Public Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Public Const SRCERASE = &H440328 ' used to determine how a blit will turn out
Public Const SRCINVERT = &H660046  ' used to determine how a blit will turn out
Public Const SRCPAINT = &HEE0086   ' used to determine how a blit will turn out
Public Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap

Public Const LR_LOADFROMFILE = 16  ' used with LoadImage
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Sub CargarImagen(Archivo As String, EnQHDC As Long, retAnchoImagen As Long, retAltoImagen As Long)
    
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
    EnQHDC = CreateCompatibleDC(0)
    
    Set stdPicAux = LoadPicture(Archivo)
    lresult = SelectObject(EnQHDC, stdPicAux.Handle)
    
    AnchoSprite = qFormularioAuxiliar.ScaleX(stdPicAux.Width, vbHimetric, vbPixels)
    AltoSprite = qFormularioAuxiliar.ScaleX(stdPicAux.Height, vbHimetric, vbPixels)
    
    'ElDestinoHDC = DestinoHDC
    
    If AnchoSprite < 1 Then Ancho = 1
    If AnchoSprite < 1 Then Alto = 1
    
    retAnchoImagen = AnchoSprite
    retAltoImagen = AltoSprite
'DESCARGAR MEMORIA:--------------------------
    ret = DeleteObject(lresult)
'--------------------------------------------
End Sub


