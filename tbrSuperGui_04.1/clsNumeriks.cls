VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumeriks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim qAlphaCantidad As Long
Dim qUsarAlpha As Boolean
Dim qAlphacolor As Long
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


Dim DTxt As New pTextEx_Demo.clsMain
Dim DTxtContorno As New pTextEx_Demo.clsMain

Dim SpriteHDC As Long


Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim PNG1 As String 'Archivo PNG Marco Disco

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColor As Long
Dim rFntColorContorno As Long

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

Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, _
    FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, _
    FontColor As Long, FntColorContorno As Long)
    
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    rFntColor = FontColor
    rFntColorContorno = FntColorContorno

    DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColor
    DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
End Sub

Public Function GetAncho()
    GetAncho = zAncho
End Function
Public Function GetAlto()
    GetAlto = zAlto
End Function

Public Sub IniciarGraficos(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, _
    qAncho As Long, qAlto As Long, AlphaHabilitado As Boolean, AlphaColor As Long, AlphaCantidad As Long)
    
    
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    zX = XFondo
    zY = YFondo
    zAncho = qAncho
    zAlto = qAlto
    
    qUsarAlpha = AlphaHabilitado
    qAlphaCantidad = AlphaCantidad
    qAlphacolor = AlphaColor
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, zAncho, zAlto
    IniHDC FondoHDC_BUP, zAncho, zAlto
    
    IniHDC SpriteHDC, zAncho, zAlto
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, DestObjHdc, XFondo, YFondo, SRCCOPY
    BitBlt FondoHDC_BUP, 0, 0, zAncho, zAlto, DestObjHdc, XFondo, YFondo, SRCCOPY

    
    If qUsarAlpha = True Then
        myAlpha.CrearHDC_Color AlphaColor
    End If
    
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
Private Sub GetAnchoAltoFont(aForm As Form, CadenaString As String, qFontName As String, _
    qFontSize As Long, qFontBold As Boolean, ByRef ElAncho As Long, ByRef ElAlto As Long)
    
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

Private Function ObtenerTextoConPuntitos(Texto As String, Ancho As Long) As String
    'Si se pasa del ancho dibuja ...
    Dim Ancho3p As Long
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim I As Long
    Dim TextoFor As String
    
    '1: Obtener el ancho de 3 puntitos
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, "...", rFntName, rFntSize, rFntBold, Ancho3p, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    
    '2 busco hasta pasarme del (ancho - Ancho de 3 puntitos)
    For I = 1 To Len(Texto)
        TextoFor = mID(Texto, 1, I)
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, TextoFor, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
        If AnchoTexto > (Ancho - Ancho3p) Then
            TextoFor = mID(Texto, 1, I - 1) + "..."
            ObtenerTextoConPuntitos = TextoFor
            Exit Function
        End If
        
    Next I
    
    'Si llegue aca, por que no hacen falta los ...
    ObtenerTextoConPuntitos = Texto
End Function

Private Sub DibujarTextContorno(Texto As String, HDC_Destino As Long, ttX As Long, ttY As Long, ttAncho As Long, ttAlto As Long)
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY + 1, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY + 1, ttAncho, ttAlto
    
    DTxt.Dibujar Texto, HDC_Destino, ttX, ttY, ttAncho, ttAlto
End Sub

'=================================================================================
Public Sub SetNumero(NuevoNumero As String, Optional SeleccionadoFX As Boolean, _
    Optional ForzarDigitosCantidad As Long = -1)
    
    Dim AnchoFont As Long
    Dim AltoFont As Long
    Dim TxtPunt As String
    Dim mX As Long
    Dim mY As Long
    
    Dim FormatoZero As String
    Dim I As Long
    
    RestaurarFondo
    
    If ForzarDigitosCantidad > 0 Then
        FormatoZero = ""
        For I = 1 To ForzarDigitosCantidad
            FormatoZero = FormatoZero + "0"
        Next I
        NuevoNumero = Format(NuevoNumero, FormatoZero)
    End If
    
    If qUsarAlpha = True Then
        myAlpha.MostrarAlpha FondoHDC, 0, 0, qAlphaCantidad, zAncho, zAlto
    End If
    
    'Dibujar el Texto
    TxtPunt = ObtenerTextoConPuntitos(NuevoNumero, (zAncho * 0.9))
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, TxtPunt, rFntName, rFntSize, False, AnchoFont, AltoFont
    qAlgunFormulario.ScaleMode = 3
    'Calculo Letras a mostrar
    mX = qX + (zAncho / 2 - AnchoFont / 2)
    mY = qY + (zAlto / 2 - AltoFont / 2)
    mY = mY - 2
    'mY = mY * 0.7
    
    If SeleccionadoFX = True Then
        mX = qX + 1
        mY = qY + 1
        DibujarTextContorno TxtPunt, FondoHDC, mX + 1, mY + 1, zAncho, zAlto
    Else
        DibujarTextContorno TxtPunt, FondoHDC, mX, mY, zAncho, zAlto
    End If
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
    obSGO.SetProp "FntColorContorno", CStr(rFntColorContorno)
    
    obSGO.SetProp "FontItalic", "0" 'no configurables por ahora
    obSGO.SetProp "FontStrike", "0"
    obSGO.SetProp "FontUnderLine", "0"
    
    obSGO.SetProp "AlphaCantidad", CStr(qAlphaCantidad)
    obSGO.SetProp "alphaHabilitado", CStr(CLng(qUsarAlpha))
    obSGO.SetProp "AlpahColorLong", CStr(qAlphacolor)
    
End Sub


'todas las clases del manu que tengan archivos externos
'deben devolver un empaquetado de las imagenes (predeterminado es solo la lista para que empaquete despues todo junto
Public Function GetFIMGs(Optional pathEmpaquetar As String = "") As clsFullImgs
    
    Dim F As New clsFullImgs, F2 As clsFullImgs_FILE
    
    'Set F2 = F.AddFileByPath(PNG1)
    'F2.sID = "PNGMarcoDisco"
    'F2.TypeFile = mImagePNG
    
    'If pathEmpaquetar <> "" Then
    '    If fso.FileExists(pathEmpaquetar) Then fso.DeleteFile pathEmpaquetar, True
    '    F.SaveFile pathEmpaquetar
    'End If
    
    Set GetFIMGs = F
    
End Function
