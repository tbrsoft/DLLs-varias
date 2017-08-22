VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRokolaIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Dim rFntColor As Long
Dim qAlgunFormulario As Form

Private Type PropiedadesFuentes
    qFntName As String
    qFntSize As Long
End Type

Private Type PropiedadesIndicador
    Valor As String
    Recta As RECT
End Type

'------------------------
Dim FntComun As PropiedadesFuentes
Dim FntCalculadora As PropiedadesFuentes
Dim FntCalculadoraCen As PropiedadesFuentes
'------------------------
Dim TemaActual As PropiedadesIndicador
Dim Precios As PropiedadesIndicador
Dim Publicidad As PropiedadesIndicador
Dim Creditos As PropiedadesIndicador
Dim Tiempo As PropiedadesIndicador
'------------------------

Dim DTxt As New pTextEx_Demo.clsMain
Dim DTxtCalculadora As New pTextEx_Demo.clsMain ':D por que parece una calculadora
Dim DTxtCalculadoraCen As New pTextEx_Demo.clsMain 'Centesimos de la calculadora

Dim ShowPNG As New tbrPNG_3.tbrPNG3

Dim myAlpha As New tbrAlphaB

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim PNG1 As String 'Archivo PNG Marco Disco

Dim qRect As RECT

Dim DestObjHdc As Long
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long
Dim zAnchoContenedor As Long
Dim zAltoContenedor As Long

Dim XCentradoPNG As Long
Dim YCentradoPNG As Long

Dim AnchoPNG As Long
Dim AltoPNG As Long

Dim FondoHDC As Long
Dim FondoHDC_BUP As Long

'------------------------------------------------------------------------------------------
'Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


Public Sub Renderizar()
    On Local Error Resume Next ' a veces se llama a renderizar antes de definir la fuente !!!
    'Restauro FondoHDC con el BUP
    BitBlt FondoHDC, 0, 0, zAncho, zAlto, FondoHDC_BUP, 0, 0, SRCCOPY
    
    'Imprimo todos los textos en FondoHDC
    ImprimirTexto TemaActual
    ImprimirTexto Precios
    ImprimirTexto Publicidad
    ImprimirTexto Tiempo
    'ImprimirTextoCalculadora Tiempo
    ImprimirTextoCalculadora Creditos
    
    'Form1.P.Cls
    
    'Muestro FondoHDC en Pantalla
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
End Sub

Public Sub IniciarFuente(AlgunFormulario As Object)
    Set qAlgunFormulario = AlgunFormulario
    
    FntComun.qFntName = "Arial": FntComun.qFntSize = 10
    'FntCalculadora.qFntName = "Pocket Calculator": FntCalculadora.qFntSize = 40
    'FntCalculadora.qFntName = "Digital-7": FntCalculadora.qFntSize = 40
    FntCalculadora.qFntName = "Transponder AOE": FntCalculadora.qFntSize = 30
    FntCalculadoraCen.qFntName = FntCalculadora.qFntName: FntCalculadoraCen.qFntSize = 16
    
    'FntComun.qFntName = "Arial": FntComun.qFntSize = 10
    'FntCalculadora.qFntName = "Arial": FntCalculadora.qFntSize = 22
    'FntCalculadoraCen.qFntName = FntCalculadora.qFntName: FntCalculadoraCen.qFntSize = 14
    
    rFntColor = RGB(220, 220, 220)
    'rFntColor = vbRed
    
    DTxt.Iniciar FntComun.qFntName, FntComun.qFntSize, True, False, False, False, rFntColor
    DTxtCalculadora.Iniciar FntCalculadora.qFntName, FntCalculadora.qFntSize, False, False, False, False, rFntColor
    DTxtCalculadoraCen.Iniciar FntCalculadoraCen.qFntName, FntCalculadoraCen.qFntSize, False, False, False, False, rFntColor
End Sub

Sub IniciarPNGs(PNG_Marco_file As String)
    Dim PNG_Token As Long
    
    PNG1 = PNG_Marco_file
    
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.GetAnchoAlto PNG1, AnchoPNG, AltoPNG
    ShowPNG.FreeGDIPlus PNG_Token
End Sub

Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long, AnchoContenedor As Long, AltoContenedor As Long, AlphaB As Boolean)
    Dim LetraColor As Long
    Dim TempBMP As Long
    Dim ObjCreado2 As Long
    
    
    zX = XFondo
    zY = YFondo
    zAncho = AnchoFondo
    zAlto = AltoFondo
    zAnchoContenedor = AnchoContenedor
    zAltoContenedor = AltoContenedor
    
    
    DestObjHdc = DondeObjetoHdc
    
    IniHDC FondoHDC, AnchoFondo, AltoFondo
    IniHDC FondoHDC_BUP, AnchoFondo, AltoFondo
    
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
    
    'con estas dos lineas es como que creo un marco de semitransparencia
    'negra para esta clase... como un fondo ;)
    If AlphaB = True Then
        myAlpha.CrearHDC_Color vbBlack
        myAlpha.MostrarAlpha FondoHDC, 0, 0, 50, AnchoFondo, AltoFondo
        myAlpha.CerrarGraficos
    End If

    'Cargo el PNG de fondo en FondoHDC
    MostrarRI
    
    'Creo un Back up de FondoHDC
    BitBlt FondoHDC_BUP, 0, 0, AnchoFondo, AltoFondo, FondoHDC, 0, 0, SRCCOPY

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

Private Sub MostrarRI()
    'Se llama solo 1 vez!!!
    
    '1. Subo el Back Up del Fondo
    'BitBlt DestObjHdc, zX, zY, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
    XCentradoPNG = ((zAncho / 2) - (AnchoPNG / 2)) + 3
    YCentradoPNG = ((zAlto / 2) - (AltoPNG / 2)) + 3
    mostrarPNG PNG1, XCentradoPNG, YCentradoPNG
End Sub


'=================================================================================
Sub mostrarPNG(path As String, mX As Long, mY As Long, Optional sel As Boolean)
    Dim PNG_Token As Long
    PNG_Token = ShowPNG.InitGDIPlus
    ShowPNG.LoadPictureGDIPlus path, mX, mY, FondoHDC
    ShowPNG.FreeGDIPlus PNG_Token
End Sub
'=================================================================================

Public Function GetPNGAncho() As Long
    GetPNGAncho = AnchoPNG
End Function

Public Function GetPNGAlto() As Long
    GetPNGAlto = AltoPNG
End Function

Private Sub Class_Initialize()
    If mpSYS.Name = eCHUgames Then
        TemaActual.Recta.qLeft = 8: TemaActual.Recta.qTop = 8
        TemaActual.Recta.qRight = 430: TemaActual.Recta.qBottom = 23
        
        Publicidad.Recta.qLeft = 175: Publicidad.Recta.qTop = 45
        Publicidad.Recta.qRight = 200: Publicidad.Recta.qBottom = 33
        
        Precios.Recta.qLeft = -35: Precios.Recta.qTop = 35
        Precios.Recta.qRight = 180: Precios.Recta.qBottom = 49
    
        Creditos.Recta.qLeft = 393: Creditos.Recta.qTop = 33
        Creditos.Recta.qRight = 100: Creditos.Recta.qBottom = 49
        
        Tiempo.Recta.qLeft = 441: Tiempo.Recta.qTop = 8
        Tiempo.Recta.qRight = 52: Tiempo.Recta.qBottom = 23
    Else
        TemaActual.Recta.qLeft = 8: TemaActual.Recta.qTop = 8
        TemaActual.Recta.qRight = 430: TemaActual.Recta.qBottom = 23
        
        Publicidad.Recta.qLeft = 190: Publicidad.Recta.qTop = 33
        Publicidad.Recta.qRight = 200: Publicidad.Recta.qBottom = 33
        
        Precios.Recta.qLeft = 8: Precios.Recta.qTop = 33
        Precios.Recta.qRight = 180: Precios.Recta.qBottom = 49
    
        Creditos.Recta.qLeft = 393: Creditos.Recta.qTop = 33
        Creditos.Recta.qRight = 100: Creditos.Recta.qBottom = 49
        
        Tiempo.Recta.qLeft = 441: Tiempo.Recta.qTop = 8
        Tiempo.Recta.qRight = 52: Tiempo.Recta.qBottom = 23
    End If
End Sub


'==========================================================================
'Los datos que se muesran
'==========================================================================
'Estas dos funciones escriben los textos en FondoHDC, no se ven en
'pantalla hasta que se llama a la funcion que renderiza

'Esta funcion se usa para setear todos los textos menos
'los dos que son onda calculiadora... digo calculadora :$
Private Sub ImprimirTexto(EnQIndicador As PropiedadesIndicador)
    Dim fntAncho As Long
    Dim fntAlto As Long
    Dim fntX As Long
    Dim fntY As Long

    Dim txtPuntitos As String
    txtPuntitos = GetTextoConPuntitos(EnQIndicador.Valor, fntX, fntY, EnQIndicador.Recta.qRight, zAnchoContenedor)
    
    GetAnchoAltoFont qAlgunFormulario, txtPuntitos, FntComun.qFntName, FntComun.qFntSize, False, fntAncho, fntAlto
    
    fntX = (EnQIndicador.Recta.qRight / 2) - (fntAncho / 2)
    fntY = (EnQIndicador.Recta.qBottom / 2) - (fntAlto / 2)
    fntX = fntX + EnQIndicador.Recta.qLeft + XCentradoPNG
    fntY = fntY + EnQIndicador.Recta.qTop + YCentradoPNG
    
    DTxt.Dibujar txtPuntitos, FondoHDC, fntX, fntY, zAnchoContenedor, zAltoContenedor
End Sub

'Esta funcion se usa para setear dos que son onda calculiadora... digo calculadora :$
Private Sub ImprimirTextoCalculadora(EnQIndicador As PropiedadesIndicador)
    Dim fntAncho2 As Long ' Necesito recordar el ancho de la letra grande para saber el X de la letra grande
    Dim fntAncho As Long
    Dim fntAlto As Long
    Dim fntX As Long
    Dim fntY As Long
    
    Dim Valor As Double
    Dim nEnteroI As Double
    Dim nDecimalI As Double
    Dim nEnteroS As String
    Dim nDecimalS As String
    
    Valor = CDbl(EnQIndicador.Valor)
    Valor = Round(Valor, 2)
    
    nEnteroI = Fix(Valor)
    Dim digs As Long
    digs = Len(CStr(nEnteroI))
    If digs < 2 Then
        digs = 2 'minimo 2 digitos para los enteros del credito
    End If
    nEnteroS = Format(nEnteroI, String(digs, "0"))
    
    nDecimalI = (Valor - Fix(Valor)) * 100
    nDecimalS = Format(nDecimalI, "00")
    
    Dim showDecimals As Boolean
    If (digs > 2) Then
        If (nDecimalI = 0) Then showDecimals = False
    Else
        showDecimals = True
    End If

    Dim X_SumAux As Long
    If Len(Str(nEnteroI)) = 1 Then X_SumAux = 30
    If Len(Str(nEnteroI)) = 2 Then X_SumAux = 20
    If Len(Str(nEnteroI)) = 3 Then X_SumAux = 10
    If Len(Str(nEnteroI)) = 4 Then X_SumAux = 3
    If Len(Str(nEnteroI)) = 5 Then X_SumAux = -2
    
    'Primero la letra Grande
    fntX = 4 + X_SumAux: fntY = -1
    GetAnchoAltoFont qAlgunFormulario, nEnteroS, FntCalculadora.qFntName, FntCalculadora.qFntSize, False, fntAncho2, fntAlto
    
    fntX = fntX + EnQIndicador.Recta.qLeft + XCentradoPNG
    fntY = fntY + EnQIndicador.Recta.qTop + YCentradoPNG
    
    DTxtCalculadora.Dibujar nEnteroS, FondoHDC, fntX, fntY, zAnchoContenedor, zAltoContenedor
    
    'si no hay decimales debo salir iaaaa, esta listo el entero
    If showDecimals = False Then Exit Sub
    
    'Ahora la letra Chica
    fntX = 6 + X_SumAux: fntY = fntY + 6
    GetAnchoAltoFont qAlgunFormulario, nDecimalS, FntCalculadoraCen.qFntName, FntCalculadora.qFntSize, False, fntAncho2, fntAlto
    
    fntX = fntX + EnQIndicador.Recta.qLeft + XCentradoPNG + fntAncho2
    
    DTxtCalculadoraCen.Dibujar nDecimalS, FondoHDC, fntX, fntY, zAnchoContenedor, zAltoContenedor

End Sub
'---------------------------------------------
Public Sub SetTemaActual(Texto As String)
    TemaActual.Valor = Texto
End Sub
Public Function GetTemaActual() As String
    GetTemaActual = TemaActual.Valor
End Function
'---------------------------------------------
Public Sub SetPromocion(Texto As String)
    Precios.Valor = Texto
End Sub

Public Sub SetPromocion2(PrecioMusica As Double, PrecioVideos As Double)
    Dim TMP As String
    
    Select Case PrecioMusica
        Case 0
            TMP = "MUSICA GRATIS"
        Case Else
            TMP = "Música $ " + CStr(Round(PrecioMusica, 2))
    End Select
    
    Select Case PrecioVideos
        Case 0
            TMP = TMP + vbCrLf + "VIDEOS GRATIS"
        Case Else
            TMP = TMP + vbCrLf + "Videos $ " + CStr(Round(PrecioVideos, 2))
    End Select
    
    Precios.Valor = TMP
End Sub

Public Function GetPromocion() As String
    GetPromocion = Precios.Valor
End Function
'---------------------------------------------
Public Sub SetPublicidad(Texto As String)
    Publicidad.Valor = Texto
End Sub
Public Function GetPublicidad() As String
    GetPublicidad = Publicidad.Valor
End Function
'---------------------------------------------
Public Sub SetCreditos(Texto As String)
    Creditos.Valor = Texto
End Sub
Public Function GetCreditos() As String
    GetCreditos = Creditos.Valor
End Function
'---------------------------------------------
Public Sub SetTiempo(Texto As String)
    Tiempo.Valor = Texto
End Sub
Public Function GetTiempo() As String
    GetTiempo = Tiempo.Valor
End Function

Private Function GetTextoConPuntitos(Texto As String, X As Long, Y As Long, Ancho As Long, Alto As Long) As String
    'Si se pasa del ancho dibuja ...
    Dim Ancho3p As Long
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim i As Long
    Dim TextoFor As String
    
    '1: Obtener el ancho de 3 puntitos
    'GetAnchoAltoFont qAlgunFormulario, "...", rFntName, rFntSize, rFntBold, Ancho3p, AltoTexto
    GetAnchoAltoFont qAlgunFormulario, "...", FntComun.qFntName, FntComun.qFntSize, False, Ancho3p, AltoTexto
    
    '2 busco hasta pasarme del (ancho - Ancho de 3 puntitos)
    For i = 1 To Len(Texto)
        TextoFor = Mid(Texto, 1, i)
        'GetAnchoAltoFont qAlgunFormulario, TextoFor, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        GetAnchoAltoFont qAlgunFormulario, TextoFor, FntComun.qFntName, FntComun.qFntSize, False, AnchoTexto, AltoTexto
        If AnchoTexto > (Ancho - Ancho3p) Then
            TextoFor = Mid(Texto, 1, i - 1) + "..."
            
            'DTxt.Dibujar TextoFor, FondoHDC, x, y, Ancho, Alto
            'DibujarTextContorno TextoFor, SpriteHDC, x, y, Ancho, Alto
            GetTextoConPuntitos = TextoFor
            Exit Function
        End If
        
    Next i
    
    'Si llegue aca, por que no hacen falta los ...
    'DTxt.Dibujar Texto, FondoHDC, x, y, Ancho, Alto
    'DibujarTextContorno Texto, SpriteHDC, x, y, Ancho, Alto
    
    GetTextoConPuntitos = Texto
End Function
