VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrTextoSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'listbox que puede ir horizontar NO GIRA

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private mElemento() As String

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColorContorno As Long

Dim rFntCol1 As Long
Dim rFntCol2 As Long

Dim DTxt As New pTextEx_Demo.clsMain
Dim DTxtContorno As New pTextEx_Demo.clsMain

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

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long)
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
    
    SpriteHDC = CreateCompatibleDC(0)
    IniHDC SpriteHDC, AnchoFondo, AltoFondo
    
    FondoHDC = CreateCompatibleDC(0)
    IniHDC FondoHDC, AnchoFondo, AltoFondo
    
    'Aqui Cargo FondoHDC
    BitBlt FondoHDC, 0, 0, AnchoFondo, AltoFondo, DestObjHdc, XFondo, YFondo, SRCCOPY
End Sub
 
Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColorUnsel As Long, FontColorSel As Long, FntColorContorno As Long)
    
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    
    rFntCol1 = FontColorUnsel
    rFntCol2 = FontColorSel
    
    rFntColorContorno = FntColorContorno

End Sub

Public Sub DibujarTexto(ElemenSelecIndex As Long, Optional ModoVertical As Boolean = False)
    Dim Conta As Long
    
    Dim XActual As Long 'Calculo en que x empiezo a dibujar
    Dim ElTexto As String
    
    Dim dDibujar As Boolean
    
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim i As Long
    
    Dim lAncho As Long
    Dim lAlto As Long
    
    If ModoVertical = True Then
        DoModoVertical ElemenSelecIndex
        Exit Sub
    End If
    
    If ElemenSelecIndex < 1 Or ElemenSelecIndex > UBound(mElemento) Then Exit Sub
    
    '>Primero Dibujo el elemento seleccionado EN EL MEDIO
    
    ElTexto = mElemento(ElemenSelecIndex)
    'Obtengo el ancho de "ElTexto" para imprimir el resto despues
    
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    
    XActual = zAncho / 2 - AnchoTexto / 2
    qRect.qLeft = XActual
    qRect.qTop = 0
    qRect.qRight = zAncho
    qRect.qBottom = zAlto

    DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
    DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
    
    'Reinicio el fondo
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    'Dibujo el texto
    'DTxt.Dibujar ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
    DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
    
    'SI SOLO TIENE UN ELEMENTO NO LO REPITO ATRAS Y ADELATE!!
    If UBound(mElemento) = 1 Then GoTo FINDIB
    
    XActual = XActual + AnchoTexto
    
    '>Segundo Dibujo elementos a la derecha hasta pasar el borde
    Conta = ElemenSelecIndex
    dDibujar = True
    While dDibujar = True
        Conta = Conta + 1
        If Conta > UBound(mElemento) Then
            Conta = 1
        End If
        ElTexto = " - " + mElemento(Conta)
        
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, lAncho, lAlto
        qAlgunFormulario.ScaleMode = 3
        
        qRect.qLeft = XActual
        qRect.qTop = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        
        'Lo dibujo
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        'DTxt.Dibujar ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        
        If XActual > zAncho Then dDibujar = False
        XActual = XActual + lAncho
    Wend
    
    
    'XActual = XActual - AnchoTexto
    XActual = zAncho / 2 - AnchoTexto / 2
    
    '>Tercero Dibujo elementos a la izquierda hasta pasar el borde
    Conta = ElemenSelecIndex
    dDibujar = True
    While dDibujar = True
        Conta = Conta - 1
        If Conta < 1 Then
            Conta = UBound(mElemento)
        End If
        ElTexto = mElemento(Conta) + " - "
        
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, lAncho, lAlto
        qAlgunFormulario.ScaleMode = 3
        
        XActual = XActual - lAncho
        
        qRect.qLeft = XActual
        qRect.qTop = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        
        'Lo dibujo
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        'DTxt.Dibujar ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        
        If XActual < 0 Then dDibujar = False
    Wend
    
    
FINDIB:
    'Y por ultimo... Render!
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
End Sub

Private Sub DoModoVertical(ElemenSelecIndex As Long)
    Dim ElemEntran As Long
    Dim DesdeElem As Long
    
    Dim i As Long
    
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    
    Dim YActual As Long 'Calculo en que y empiezo a dibujar
    
    If ElemenSelecIndex < 1 Or ElemenSelecIndex > UBound(mElemento) Then Exit Sub
    
    'Reinicio el fondo
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
    '>Primero Dibujo del 1er elemento hasta el seleccionado
    
    'Obtengo el alto de mElemento(1),  para imprimir el resto despues
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, mElemento(1), rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    
    
    
    
    
    If (AltoTexto) * (ElemenSelecIndex) < zAlto Then
    'Si el elemento que selecciono NO esta por debajo del limite de alto
    
        
        
        YActual = 0
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        For i = 1 To ElemenSelecIndex - 1
            qRect.qTop = YActual
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next i
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        qRect.qTop = YActual
        '-----------------------------------
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
        'Dibujo el texto
        DibujarTextoConPuntitos mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        '-----------------------------------
        YActual = YActual + AltoTexto
        
        
        '>Tercero, dibujo el resto
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, mElemento(ElemenSelecIndex), rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
        
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        For i = ElemenSelecIndex + 1 To UBound(mElemento)
            If YActual + AltoTexto > zAlto Then GoTo HacerRender
            
            qRect.qTop = YActual
            
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next i
    
    Else
    'Si el elemento que selecciono ESTA por debajo del limite de alto
    
        'Calcular desde que elemento deberia empezar a mostrar, para que el elemento seleccionado se vea
        'Cuantos elementos entran???
        ElemEntran = zAlto / AltoTexto
        'desde que elemento empiezo a mostrar???
        DesdeElem = ElemenSelecIndex - ElemEntran + 1
        
        
        YActual = 0
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        'For i = DesdeElem + 1 To ElemenSelecIndex - 1
        For i = DesdeElem To ElemenSelecIndex - 1
            qRect.qTop = YActual
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next i
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        qRect.qTop = YActual
        '-----------------------------------
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
        'Dibujo el texto
        DibujarTextoConPuntitos mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        '-----------------------------------
        YActual = YActual + AltoTexto
        
    End If
    
HacerRender:
    'Y por ultimo... Render!
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY

End Sub
 
Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(DestObjHdc)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

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


'---------------------Heredado-------------------
Private Sub Class_Initialize()
    ReDim mElemento(0)
End Sub

Public Function AgregarElemento(NombreElemento As String) As Long 'Devuelvo index agregado
    Dim Ix As Long
    Ix = UBound(mElemento) + 1
    ReDim Preserve mElemento(Ix)
    mElemento(Ix) = NombreElemento
    AgregarElemento = Ix
End Function

Public Sub QuitarElemento(Index As Long)
    If UBound(mElemento) = 0 Then
        ReDim mElemento(0)
        Exit Sub
    End If
    Dim i As Long
    For i = Index To UBound(mElemento) - 1
        mElemento(i) = mElemento(i + 1)
    Next i
    ReDim Preserve mElemento(UBound(mElemento) - 1)
End Sub

Public Sub SetElemento(Index As Long, NombreElemento As String)
    mElemento(Index) = NombreElemento
End Sub

Public Function GetElemento(Index As Long) As String
    If UBound(mElemento) > 0 Then
        GetElemento = mElemento(Index)
    Else
        GetElemento = "-1"
    End If
End Function

Public Function GetElementoCount() As Long
    GetElementoCount = UBound(mElemento)
End Function

'--------------------------------------------------

Public Sub ResetElementos()
    ReDim mElemento(0)
End Sub

Private Sub DibujarTextoConPuntitos(Texto As String, x As Long, y As Long, Ancho As Long, Alto As Long)
    'Si se pasa del ancho dibuja ...
    Dim Ancho3p As Long
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim i As Long
    Dim TextoFor As String
    
    '1: Obtener el ancho de 3 puntitos
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, "...", rFntName, rFntSize, rFntBold, Ancho3p, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    
    '2 busco hasta pasarme del (ancho - Ancho de 3 puntitos)
    For i = 1 To Len(Texto)
        TextoFor = Mid(Texto, 1, i)
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, TextoFor, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
    
        If AnchoTexto > (Ancho - Ancho3p) Then
            TextoFor = Mid(Texto, 1, i - 1) + "..."
            'DTxt.Dibujar TextoFor, SpriteHDC, X, Y, Ancho, Alto
            DibujarTextContorno TextoFor, SpriteHDC, x, y, Ancho, Alto
            Exit Sub
        End If
        
    Next i
    
    'Si llegue aca, por que no hacen falta los ...
    'DTxt.Dibujar Texto, SpriteHDC, X, Y, Ancho, Alto
    DibujarTextContorno Texto, SpriteHDC, x, y, Ancho, Alto
End Sub
'el private lo pone andresv en set09 por que me parece que es privado
Private Sub DibujarTextContorno(Texto As String, HDC_Destino As Long, ttX As Long, ttY As Long, ttAncho As Long, ttAlto As Long)
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY + 1, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY + 1, ttAncho, ttAlto
    
    
    DTxt.Dibujar Texto, HDC_Destino, ttX, ttY, ttAncho, ttAlto
End Sub

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
    ret = DeleteDC(FondoHDC)
End Sub



