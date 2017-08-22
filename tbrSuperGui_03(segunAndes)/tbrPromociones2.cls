VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrPromociones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------
Dim qTiemposEntreTextos As Long
Dim SumarTiempos As Boolean
Dim TiemposSumados As Long
'------------------------

'------------------------
Const AC_SRC_OVER = &H0

Private Type AlphaOptions
  AlphaOption As Byte
  AlphaFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
'------------------------


Dim DoPegarFondo As Boolean

Private qPromocion() As String

'------------------------
Dim qPixelSalteo As Long

Dim iActual As Long

Dim XActual As Long
Dim AlphaActual As Long
'------------------------

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColor As Long

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

Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, _
    AnchoFondo As Long, AltoFondo As Long, PixelSalteo As Long, TiemposEntreTextos As Long)
    
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
    
    DoPegarFondo = True
    
    qTiemposEntreTextos = TiemposEntreTextos
End Sub
 
Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColor As Long)
    
    DTxt.Iniciar FontName, FontSize, FontBold, FontItalic, FontStrikethrough, FontUnderLine, FontColor
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    rFntColor = FontColor
End Sub

Public Sub DibujarTexto()
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    
    If SumarTiempos = True Then
        TiemposSumados = TiemposSumados + 1
        If TiemposSumados > qTiemposEntreTextos Then
            TiemposSumados = 0
            SumarTiempos = False
        End If
        Exit Sub
    End If
    '---------------------
    Dim AO As AlphaOptions
    Dim ptAO As Long
    '---------------------
    
    If DoPegarFondo = True Then
        
        'Restaurar el fondo
        BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
        BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
        
        'Dibujar Texto------------------------
        qRect.qLeft = 0
        qRect.qTop = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, qPromocion(iActual), rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
        DTxt.Dibujar qPromocion(iActual), SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        '-------------------------------------
        
        DoPegarFondo = False
    End If
    
    'XActual = XActual - qPixelSalteo
    AlphaActual = AlphaActual + qPixelSalteo
    
    If AlphaActual > 50 Then
        iActual = iActual + 1
        If iActual > UBound(qPromocion) Then iActual = 1
        AlphaActual = 0
        
        'Restauro el fondo
        DoPegarFondo = True
        SumarTiempos = True
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------
    'If qAncho = -1 Then
    '    qAncho = Ancho
    'End If
    'If qAlto = -1 Then
    '    qAlto = Alto
    'End If
    
    With AO
        .AlphaOption = AC_SRC_OVER
        .AlphaFlags = 0
        .SourceConstantAlpha = 5
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory ptAO, AO, 4
    AlphaBlend DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, zAncho, zAlto, ptAO
    '-------------------------------------------------------------------------

    'BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
End Sub
 
Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(DestObjHdc)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    ret = DeleteObject(ObjCreado2)
    ret = DeleteObject(TempBMP)
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

'------------------------------------------------
Public Sub AgregarPromo(qTexto As String)
    Dim Ix As Long
    Ix = UBound(qPromocion) + 1
    ReDim Preserve qPromocion(Ix)
    qPromocion(Ix) = qTexto
    
    Reset
End Sub

Public Sub QuitarPromo(Index As Long)
    If UBound(qPromocion) = 0 Then
        ReDim qPromocion(0)
        Exit Sub
    End If
    
    Dim I As Long
    For I = Index To UBound(qPromocion) - 1
        qPromocion(I) = qPromocion(I + 1)
    Next I
    ReDim Preserve qPromocion(UBound(qPromocion) - 1)
End Sub

Public Sub SetPromo(Index As Long, qTexto As String)
    qPromocion(Index) = qTexto
End Sub

Public Function GetPromoCount() As Long
    GetPromoCount = UBound(qPromocion)
End Function

Public Function GetPromo(Index As Long) As String
    GetPromo = qPromocion(Index)
End Function

Public Sub ResetPromo()
    ReDim qPromocion(0)
End Sub

Private Sub Class_Initialize()
    ResetPromo
End Sub

Public Sub Reset()
    SumarTiempos = False
    AlphaActual = 0
    iActual = 1
End Sub

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
    ret = DeleteDC(FondoHDC)
    
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
    
    obSGO.SetProp "FontItalic", "0" 'no configurables por ahora
    obSGO.SetProp "FontStrike", "0"
    obSGO.SetProp "FontUnderLine", "0"
    
    obSGO.SetProp "TiemposEntreTextos", CStr(qTiemposEntreTextos)
    obSGO.SetProp "PixelSalteo", CStr(qPixelSalteo)

    
End Sub

'todas las clases del manu que tengan archivos externos
'deben devolver un empaquetado de las imagenes (predeterminado es solo la lista para que empaquete despues todo junto
Public Function GetFIMGs(Optional pathEmpaquetar As String = "") As clsFullImgs
    'ESTE NO TIENE ARCHIVOSA EXTERNOS
    
    Dim F As New clsFullImgs, F2 As clsFullImgs_FILE
    
    'Set F2 = F.AddFileByPath(qPNGpathUnSel)
    'F2.sID = "PNGUnSel"
    'F2.TypeFile = mImagePNG
    
    'If pathEmpaquetar <> "" Then
    '    If fso.FileExists(pathEmpaquetar) Then fso.DeleteFile pathEmpaquetar, True
    '    F.SaveFile pathEmpaquetar
    'End If
    
    Set GetFIMGs = F
    
End Function
