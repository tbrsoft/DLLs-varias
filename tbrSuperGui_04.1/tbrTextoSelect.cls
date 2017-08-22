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
Dim CACA As New tbrAlphaB 'BORRAR TODO LO RELACIONADO A ESTO (bah, sirve para depurar)

'--------------------------------------------------
Private Type RegionData
    IdRegion As Long
    IndexElemento As Long
    X As Long
    Y As Long
    X2 As Long
    Y2 As Long
End Type
'Regiones
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Dim IdRegiones() As RegionData
'--------------------------------------------------

'listbox que puede ir horizontar NO GIRAPrivate Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private mElemento() As String

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColorContorno As Long

Dim rFntCol1 As Long
Dim rFntCol2 As Long

Dim DTxt As New pTextEx_Demo.clsMain 'Fuente comun
Dim DTxtSelect As New pTextEx_Demo.clsMain 'Fuente de texto seleccionado
Dim DTxtContorno As New pTextEx_Demo.clsMain 'Fuente del contorno

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

Public Sub IniciarGraficos(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long)
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
 
Public Sub IniciarFuente(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, _
    FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColorUnsel As Long, FontColorSel As Long, FntColorContorno As Long)
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    
    rFntCol1 = FontColorUnsel
    rFntCol2 = FontColorSel
    
    rFntColorContorno = FntColorContorno

    DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
    DTxtSelect.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
    DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno

End Sub

Public Sub DibujarTexto(ElemenSelecIndex As Long, Optional ModoVertical As Boolean = False)
    
    On Local Error GoTo errDIBU
    
    terr.Anotar "texEl-001", ElemenSelecIndex, ModoVertical
    
    Dim Conta As Long
    
    Dim XActual As Long 'Calculo en que x empiezo a dibujar
    Dim ElTexto As String
    
    Dim dDibujar As Boolean
    
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    Dim I As Long
    
    Dim lAncho As Long
    Dim lAlto As Long
    
    If ModoVertical = True Then
        DoModoVertical ElemenSelecIndex
        Exit Sub
    End If
    
    
    'Reinicio las Regiones, para crear nuevas
    BorrarRegiones
    ReDim IdRegiones(0) 'choreado por andres del mpRock, el manu se lo debe haber olvidado
    
    If ElemenSelecIndex < 1 Or ElemenSelecIndex > UBound(mElemento) Then
        terr.Anotar "texEl-003"
        Exit Sub
    End If
    
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

    'Reinicio el fondo
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
    'Dibujo el elemento del medio
    DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxt
    
    'Crear la primer Region
    'Hace falta la region del elemento que YA ESTA ELEGIDO??? yo creo que no ¬¬ .
    'AgregarRegion qRect.qLeft, qRect.qTop, (qRect.qLeft + AnchoTexto), (qRect.qTop + AnchoTexto), ElemenSelecIndex

    'SI SOLO TIENE UN ELEMENTO NO LO REPITO ATRAS Y ADELATE!!
    If UBound(mElemento) = 1 Then
        terr.Anotar "texEl-009"
        GoTo FINDIB
    End If
    
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
        
        'AQUIIIII
        qAlgunFormulario.ScaleMode = 1
        
        GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, lAncho, lAlto
        
        qAlgunFormulario.ScaleMode = 3
        
        qRect.qLeft = XActual
        qRect.qTop = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        
        '-----
        If XActual > zAncho Then dDibujar = False
        '-----
        
        If dDibujar = True Then
            'Lo dibujo
            DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxtSelect
            
            'Creo Region correspondiente
            AgregarRegion zX + qRect.qLeft, zY + qRect.qTop, ((zX + qRect.qLeft) + lAncho), ((zY + qRect.qTop) + lAlto), Conta
        End If
        
        
        XActual = XActual + lAncho
    Wend
    
    terr.Anotar "texEl-011", AnchoTexto
    'XActual = XActual - AnchoTexto
    XActual = zAncho / 2 - AnchoTexto / 2
    
    '>Tercero Dibujo elementos a la izquierda hasta pasar el centro
    Conta = ElemenSelecIndex
    dDibujar = True
    While dDibujar = True
        Conta = Conta - 1
        terr.Anotar "texEl-011b", Conta, ElTexto
        If Conta < 1 Then
            Conta = UBound(mElemento)
        End If
        ElTexto = mElemento(Conta) + " - "
        qAlgunFormulario.ScaleMode = 1
        
        terr.Anotar "texEl-011c", XActual, lAncho
        GetAnchoAltoFont qAlgunFormulario, ElTexto, rFntName, rFntSize, rFntBold, lAncho, lAlto
        qAlgunFormulario.ScaleMode = 3
        XActual = XActual - lAncho
        
        terr.Anotar "texEl-011d"
        qRect.qLeft = XActual
        qRect.qTop = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        
        '-----
        If XActual < 0 Then dDibujar = False
        '-----
        
        If dDibujar = True Then
            'Lo dibujo
            DibujarTextContorno ElTexto, SpriteHDC, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxtSelect
            
            'Creo Region correspondiente
            AgregarRegion zX + qRect.qLeft, zY + qRect.qTop, ((zX + qRect.qLeft) + lAncho), ((zY + qRect.qTop) + lAlto), Conta
        End If
        
    Wend
    
    terr.Anotar "texEl-012", XActual
    
FINDIB:
    'Y por ultimo... Render!
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
    
    Exit Sub
    
errDIBU:
    terr.AppendLog "erDIBU", terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub DibujarRegiones()
    Dim Ix As Long
    For Ix = 1 To UBound(IdRegiones)
        Randomize Timer
        CACA.CerrarGraficos
        CACA.CrearHDC_Color QBColor(Rnd * 15)
        CACA.MostrarAlpha DestObjHdc, IdRegiones(Ix).X, IdRegiones(Ix).Y, 70, (IdRegiones(Ix).X2 - IdRegiones(Ix).X), (IdRegiones(Ix).Y2 - IdRegiones(Ix).Y)
    Next Ix
End Sub


Private Sub DoModoVertical(ElemenSelecIndex As Long)
    Dim ElemEntran As Long
    Dim DesdeElem As Long
    
    Dim I As Long
    
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
        'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        For I = 1 To ElemenSelecIndex - 1
            qRect.qTop = YActual
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(I), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxtSelect
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next I
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        qRect.qTop = YActual
        '-----------------------------------
        'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
        'Dibujo el texto
        DibujarTextoConPuntitos mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxt
        '-----------------------------------
        YActual = YActual + AltoTexto
        
        
        '>Tercero, dibujo el resto
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, mElemento(ElemenSelecIndex), rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
        
        'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        For I = ElemenSelecIndex + 1 To UBound(mElemento)
            If YActual + AltoTexto > zAlto Then GoTo HacerRender
            
            qRect.qTop = YActual
            
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(I), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxtSelect
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next I
    
    Else
    'Si el elemento que selecciono ESTA por debajo del limite de alto
    
        'Calcular desde que elemento deberia empezar a mostrar, para que el elemento seleccionado se vea
        'Cuantos elementos entran???
        ElemEntran = zAlto / AltoTexto
        'desde que elemento empiezo a mostrar???
        DesdeElem = ElemenSelecIndex - ElemEntran + 1
        
        
        YActual = 0
        'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        'For i = DesdeElem + 1 To ElemenSelecIndex - 1
        For I = DesdeElem To ElemenSelecIndex - 1
            qRect.qTop = YActual
            '-----------------------------------
            'Dibujo el texto
            DibujarTextoConPuntitos mElemento(I), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxtSelect
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next I
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        qRect.qTop = YActual
        '-----------------------------------
        'DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
        
        'Dibujo el texto
        DibujarTextoConPuntitos mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom, DTxt
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

    On Local Error GoTo errGAF

    terr.Anotar "texEl-010o", aForm.Name, qFontName, qFontSize, CadenaString
    ElAncho = 0
    ElAlto = 0

    aForm.FontName = qFontName
    terr.Anotar "texEl-010o2"
    
    aForm.FontSize = qFontSize
    
    terr.Anotar "texEl-010o3"
    aForm.FontBold = qFontBold
    
    terr.Anotar "texEl-010o4"
    aForm.FontItalic = False
    aForm.FontStrikethru = False
    
    terr.Anotar "texEl-010o5"
    aForm.FontUnderLine = False
    
    terr.Anotar "texEl-010o6", CadenaString
    ElAncho = (CLng(aForm.TextWidth(CadenaString)) / 15)
    ElAlto = (CLng(aForm.TextHeight(CadenaString)) / 15)
    
    terr.Anotar "texEl-010p", ElAncho, ElAlto
    
    
    Exit Sub
    
errGAF:
    terr.AppendLog "****GRAVE*****-GAF09809"
    ' que no de ceroooooooo, negradaaaaaaaaaaaa
    ElAncho = 10
    ElAlto = 10
End Sub


'---------------------Heredado-------------------
Private Sub Class_Initialize()
    ReDim mElemento(0)
    ReDim IdRegiones(0)
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
    Dim I As Long
    For I = Index To UBound(mElemento) - 1
        mElemento(I) = mElemento(I + 1)
    Next I
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

Private Sub DibujarTextoConPuntitos(Texto As String, X As Long, Y As Long, Ancho As Long, Alto As Long, _
    qDTxt As pTextEx_Demo.clsMain)        'Si se pasa del ancho dibuja ...
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
            DibujarTextContorno TextoFor, SpriteHDC, X, Y, Ancho, Alto, qDTxt
            Exit Sub
        End If
        
    Next I
    
    'Si llegue aca, por que no hacen falta los ...
    DibujarTextContorno Texto, SpriteHDC, X, Y, Ancho, Alto, qDTxt
End Sub
'el private lo pone andresv en set09 por que me parece que es privado
Private Sub DibujarTextContorno(Texto As String, HDC_Destino As Long, ttX As Long, ttY As Long, ttAncho As Long, ttAlto As Long, qDTxt As pTextEx_Demo.clsMain)
    
    terr.Anotar "texEl-010q"
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX, ttY + 1, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY, ttAncho, ttAlto
    
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX - 1, ttY - 1, ttAncho, ttAlto
    DTxtContorno.Dibujar Texto, HDC_Destino, ttX + 1, ttY + 1, ttAncho, ttAlto
    
    qDTxt.Dibujar Texto, HDC_Destino, ttX, ttY, ttAncho, ttAlto
End Sub

Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(SpriteHDC)
    ret = DeleteDC(FondoHDC)
    BorrarRegiones
End Sub

Private Sub BorrarRegiones()
    For I = 1 To UBound(IdRegiones)
        ret = DeleteObject(IdRegiones(I).IdRegion)
    Next I
    ReDim IdRegiones(0)
End Sub

Private Sub AgregarRegion(X1Region As Long, Y1Region As Long, X2Region As Long, Y2Region As Long, qIndexElemento As Long)
    
    terr.Anotar "texEl-010u", X1Region, qIndexElemento
    Dim Ix As Long
    Ix = UBound(IdRegiones) + 1
    ReDim Preserve IdRegiones(Ix)
    terr.Anotar "texEl-010v", Ix
    
    IdRegiones(Ix).IdRegion = CreateRectRgn(X1Region, Y1Region, X2Region, Y2Region)
    IdRegiones(Ix).IndexElemento = qIndexElemento
    'lo agrega andres por que la funcion que detecta solo anda como el ocote
    IdRegiones(Ix).X = X1Region
    IdRegiones(Ix).Y = Y1Region
    IdRegiones(Ix).X2 = X2Region
    IdRegiones(Ix).Y2 = Y2Region
    
    terr.Anotar "texEl-010w"
    'AgregarRegion = Ix
End Sub

Public Function DoClick_GetElementoIndex(clX As Long, clY As Long) As Long
    Dim I As Long
    DoClick_GetElementoIndex = -1
    For I = 1 To UBound(IdRegiones)
        
        If clX >= IdRegiones(I).X And clX <= IdRegiones(I).X2 Then
            If clY >= IdRegiones(I).Y And clY <= IdRegiones(I).Y2 Then
                DoClick_GetElementoIndex = IdRegiones(I).IndexElemento
                Exit Function
            End If
        End If
        
        'If PtInRegion(IdRegiones(i).IdRegion, clX, clY) <> 0 Then
        '    DoClick_GetElementoIndex = IdRegiones(i).IndexElemento
        '    'DibujarTexto i
        '    Exit Function
        'End If
    Next I
End Function

Private Function GetIxRegionByElemento(ixElemento As Long) As Long
    GetIxRegionByElemento = -1
    
    Dim M As Long
    For M = 1 To UBound(IdRegiones)
        If ixElemento = IdRegiones(M).IndexElemento Then
            GetIxRegionByElemento = M
            Exit Function
        End If
    Next M
End Function

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
    obSGO.SetProp "FontColorSel", CStr(fFntColorSel)
    obSGO.SetProp "FontColorUnSel", CStr(fFntColorUnSel)
    obSGO.SetProp "FntColorContorno", CStr(rFntColorContorno)
    
    obSGO.SetProp "FontItalic", "0" 'no configurables por ahora
    obSGO.SetProp "FontStrike", "0"
    obSGO.SetProp "FontUnderLine", "0"
    
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

Public Function ToStringRgns()
    Dim ret As String
    ret = ""
    
    Dim M As Long
    For M = 1 To UBound(IdRegiones)
        
        ret = ret + "RGN:" + CStr(M) + ": EL: (" + CStr((IdRegiones(M).IndexElemento)) + ")" + mElemento((IdRegiones(M).IndexElemento)) + vbCrLf
               
        ret = ret + " rect: " + _
            CStr(IdRegiones(M).X) + " " + CStr(IdRegiones(M).Y) + " " + _
            CStr(IdRegiones(M).X2) + " " + CStr(IdRegiones(M).Y2) + " " + vbCrLf + vbCrLf
    
        
    Next M
    
    ToStringRgns = ret

End Function

