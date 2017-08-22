VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTemasManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------
'Regiones
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Dim IdRegiones() As Long
'--------------------------------------------------

Dim myAlphaSel As New tbrAlphaB

Dim myAlpha As New tbrAlphaB

Private mElemento() As clsTemaList

Dim rFntName As String
Dim rFntSize As Long
Dim rFntBold As Boolean
Dim rFntColorContorno As Long

Dim rFntCol1 As Long
Dim rFntCol2 As Long

Dim qAlgunFormulario As Form

Dim DTxt As New pTextEx_Demo.clsMain
Dim DTxtContorno As New pTextEx_Demo.clsMain

Dim qRect As RECT

Dim DestObjHdc As Long
Dim zX As Long
Dim zY As Long
Dim zAncho As Long
Dim zAlto As Long

Dim AuxAltoTexto As Long

Dim SpriteHDC As Long
Dim FondoHDC As Long

'Private Const SRCAND = &H8800C6 ' used to determine how a blit will turn out
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

Private mIndiceElegido As Long 'ANDRESV 'permiso, ahora yo tmb lo uso =) MANUV



Public Sub IniciarGrafios(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long, AnchoFondo As Long, AltoFondo As Long, AlphaB As Boolean, ColorSel As Long)
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

    If AlphaB = True Then
        myAlpha.CrearHDC_Color vbBlack
        myAlpha.MostrarAlpha FondoHDC, 0, 0, 50, AnchoFondo, AltoFondo
        myAlpha.CerrarGraficos
    End If
    
    myAlphaSel.CrearHDC_Color ColorSel
End Sub
 
Public Sub IniciarFuente(AlgunFormulario As Form, FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColorUnsel As Long, FontColorSel As Long, FntColorContorno As Long)
    Set qAlgunFormulario = AlgunFormulario
    rFntName = FontName
    rFntSize = FontSize
    rFntBold = FontBold
    
    rFntCol1 = FontColorUnsel
    rFntCol2 = FontColorSel
    
    rFntColorContorno = FntColorContorno

End Sub

Public Sub SelNext(rotate As Boolean)

    Dim MM As Long, TrySel As Long
    TrySel = mIndiceElegido
    
    For MM = 1 To UBound(mElemento)
        If TrySel + 1 > UBound(mElemento) Then
            If rotate Then
                TrySel = 1
            Else
                'NI BOSTA
            End If
        Else
            TrySel = TrySel + 1
        End If
        
        'gira hasta encotrar uno no bloqueado
        If mElemento(TrySel).Bloqueado = False Then
            mIndiceElegido = TrySel
            'solo si cambio se dibuja
            DibujarTexto mIndiceElegido
            Exit For
        End If
    Next MM
    
    
End Sub

Public Sub DefineElegidoByIndex(sID As Long)
    mIndiceElegido = sID
    DibujarTexto mIndiceElegido
End Sub

Public Sub SelPrevious(rotate As Boolean)
    
    Dim MM As Long, TrySel As Long
    TrySel = mIndiceElegido
    
    For MM = 1 To UBound(mElemento)
        If TrySel - 1 < 1 Then
            If rotate Then
                TrySel = UBound(mElemento)
            Else
                'NI BOSTA
            End If
        Else
            TrySel = TrySel - 1
        End If
        
        'gira hasta encotrar uno no bloqueado
        If mElemento(TrySel).Bloqueado = False Then
            mIndiceElegido = TrySel
            'solo si cambio se dibuja
            DibujarTexto mIndiceElegido
            Exit For
        End If
    Next MM
End Sub

Public Function GetElegido() As clsTemaList
    If mIndiceElegido > UBound(mElemento) Then
        Set GetElegido = mElemento(0) 'supongo que es uno vacio!
    Else
        If mElemento(mIndiceElegido) Is Nothing Then
            Set GetElegido = mElemento(0) 'supongo que es uno vacio!
        Else
            Set GetElegido = mElemento(mIndiceElegido)
        End If
    End If
End Function

Public Function GetIndiceElegido() As Long
    GetIndiceElegido = mIndiceElegido
End Function

Public Sub DibujarTexto(ElemenSelecIndex As Long)
    Dim DibujarSeleccionAlpha As Boolean
    
    Dim ElemEntran As Long
    Dim DesdeElem As Long
    
    Dim i As Long
    
    Dim AnchoTexto As Long
    Dim AltoTexto As Long
    
    Dim YActual As Long 'Calculo en que y empiezo a dibujar
    
    
    If ElemenSelecIndex = -1 Then
        ElemenSelecIndex = 0
        DibujarSeleccionAlpha = False
    Else
        DibujarSeleccionAlpha = True
    End If
    If UBound(mElemento) = 0 Then Exit Sub
    If ElemenSelecIndex < 0 Or ElemenSelecIndex > UBound(mElemento) Then Exit Sub
    
    mIndiceElegido = ElemenSelecIndex 'ANDRESV
    
    'Reinicio el fondo
    BitBlt SpriteHDC, 0, 0, zAncho, zAlto, FondoHDC, 0, 0, SRCCOPY
    
    '>Primero Dibujo del 1er elemento hasta el seleccionado
    
    'Obtengo el alto de mElemento(1),  para imprimir el resto despues
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, mElemento(1).Titulo, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
    qAlgunFormulario.ScaleMode = 3
    AuxAltoTexto = AltoTexto
    
    
    
    'Si elije uno que esta bloqueado... no hago nada... Y QUE!
    If ElemenSelecIndex <> 0 Then
        If mElemento(ElemenSelecIndex).Bloqueado = True Then
            'ElemenSelecIndex = ElemenSelecIndex + 1
            Exit Sub
        End If
    End If
    
    If (AltoTexto) * (ElemenSelecIndex) < zAlto Then
    'Si el elemento que selecciono NO esta por debajo del limite de alto
    
        
        
        YActual = 0
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        For i = 1 To ElemenSelecIndex - 1
            qRect.qTop = YActual
            
            If mElemento(i).Bloqueado = True Then
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, RGB(50, 50, 50)
            Else
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
            End If
            '-----------------------------------
            'Dibujo el texto
            'DibujarTextoConPuntitos mElemento(i).Titulo, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            DibjarElemento mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next i
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        If ElemenSelecIndex <> 0 Then
            qRect.qTop = YActual
            '-----------------------------------
            DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
            DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
            'Dibujo el texto
            'DibujarTextoConPuntitos mElemento(ElemenSelecIndex).Titulo, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            If DibujarSeleccionAlpha = True Then
                myAlphaSel.MostrarAlpha SpriteHDC, qRect.qLeft, qRect.qTop - 2, 100, qRect.qRight, AltoTexto + 2
            End If
            DibjarElemento mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        End If
        
        '>Tercero, dibujo el resto
        
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        For i = ElemenSelecIndex + 1 To UBound(mElemento)
            If YActual + AltoTexto > zAlto Then GoTo HacerRender
            
            qRect.qTop = YActual
            
            If mElemento(i).Bloqueado = True Then
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, RGB(50, 50, 50)
            Else
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
            End If
            
            '-----------------------------------
            'Dibujo el texto
            'DibujarTextoConPuntitos mElemento(i).Titulo, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            DibjarElemento mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
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
        'DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
        
        qRect.qLeft = 0
        qRect.qRight = zAncho
        qRect.qBottom = zAlto
        'For i = DesdeElem + 1 To ElemenSelecIndex - 1
        For i = DesdeElem To ElemenSelecIndex - 1
            qRect.qTop = YActual
            
            If mElemento(i).Bloqueado = True Then
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, RGB(50, 50, 50)
            Else
                DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol1
            End If
            '-----------------------------------
            'Dibujo el texto
            'DibujarTextoConPuntitos mElemento(i).Titulo, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            DibjarElemento mElemento(i), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
            '-----------------------------------
            YActual = YActual + AltoTexto
        Next i
        
    
        
        '>Segundo Dibujo el elemento seleccionado
        qRect.qTop = YActual
        '-----------------------------------
        DTxtContorno.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntColorContorno
        DTxt.Iniciar rFntName, rFntSize, rFntBold, False, False, False, rFntCol2
        'Dibujo el texto
        'DibujarTextoConPuntitos mElemento(ElemenSelecIndex).Titulo, qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        If DibujarSeleccionAlpha = True Then
            myAlphaSel.MostrarAlpha SpriteHDC, qRect.qLeft, qRect.qTop - 2, 100, qRect.qRight, AltoTexto + 2
        End If
        DibjarElemento mElemento(ElemenSelecIndex), qRect.qLeft, qRect.qTop, qRect.qRight, qRect.qBottom
        '-----------------------------------
        YActual = YActual + AltoTexto
        
    End If

HacerRender:
    'Y por ultimo... Render!
    BitBlt DestObjHdc, zX, zY, zAncho, zAlto, SpriteHDC, 0, 0, SRCCOPY
End Sub
 
Sub DibjarElemento(qElemento As clsTemaList, dibX As Long, dibY As Long, dibAncho As Long, dibAlto As Long)
    'Dibujo Numero
    If qElemento.Numero > -1 Then
        DibujarTextoConPuntitos Format(qElemento.Numero, "00"), dibX + 0, dibY, (dibAncho * 0.1) + 3, dibAlto
        'Dibujo Titulo
        DibujarTextoConPuntitos qElemento.Titulo, dibX + (dibAncho * 0.06) + 3, dibY, (dibAncho * 0.8) - 3, dibAlto
    Else
        'Si no hay número entonces dibujo el Titulo mas a la izquierda
        DibujarTextoConPuntitos qElemento.Titulo, dibX + 3, dibY, (dibAncho * 0.8) - 3, dibAlto
    End If
    'Dibujo Duracion
    If qElemento.Duracion <> "" Then
        DibujarTextoConPuntitos qElemento.Duracion, (dibAncho * 0.88) - 0, dibY, (dibAncho), dibAlto
    End If
    '-----------------------
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

Public Function AgregarElemento() As clsTemaList 'Devuelvo index agregado
    Dim Ix As Long
    Ix = UBound(mElemento) + 1
    ReDim Preserve mElemento(Ix)
    Set mElemento(Ix) = New clsTemaList
    Set AgregarElemento = mElemento(Ix)
    
End Function

Public Sub QuitarElemento(index As Long)
    If UBound(mElemento) = 0 Then
        ReDim mElemento(0)
        Exit Sub
    End If
    Dim i As Long
    For i = index To UBound(mElemento) - 1
        mElemento(i) = mElemento(i + 1)
    Next i
    ReDim Preserve mElemento(UBound(mElemento) - 1)
End Sub

Public Sub SetElemento(index As Long, NombreElemento As String)
    mElemento(index) = NombreElemento
End Sub

Public Function GetElem(index As Long) As clsTemaList
    If UBound(mElemento) > 0 Then
        If index <= UBound(mElemento) Then
            Set GetElem = mElemento(index)
        End If
    Else
        Set GetElem = mElemento(0) 'esta vacio!
    End If
End Function

Public Function GetElementoString(index As Long) As String
    If UBound(mElemento) > 0 Then
        GetElemento = mElemento(index)
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
    mIndiceElegido = 0
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
        TextoFor = mID(Texto, 1, i)
        qAlgunFormulario.ScaleMode = 1
        GetAnchoAltoFont qAlgunFormulario, TextoFor, rFntName, rFntSize, rFntBold, AnchoTexto, AltoTexto
        qAlgunFormulario.ScaleMode = 3
        If AnchoTexto > (Ancho - Ancho3p) Then
            TextoFor = mID(Texto, 1, i - 1) + "..."
            'DTxt.Dibujar TextoFor, SpriteHDC, X, Y, Ancho, Alto
            DibujarTextContorno TextoFor, SpriteHDC, x, y, Ancho, Alto
            Exit Sub
        End If
        
    Next i
    
    'Si llegue aca, por que no hacen falta los ...
    'DTxt.Dibujar Texto, SpriteHDC, X, Y, Ancho, Alto
    DibujarTextContorno Texto, SpriteHDC, x, y, Ancho, Alto
End Sub

Sub DibujarTextContorno(Texto As String, HDC_Destino As Long, ttX As Long, ttY As Long, ttAncho As Long, ttAlto As Long)
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
    ret = DeleteDC(FondoHDC)
    ret = DeleteDC(SpriteHDC)
    myAlpha.CerrarGraficos
End Sub

Public Function GetX() As Long
    GetX = zX
End Function
Public Function GetY() As Long
    GetY = zY
End Function
Public Function GetAncho() As Long
    GetAncho = zAncho
End Function
Public Function GetAlto() As Long
    GetAlto = zAlto
End Function
Public Function GetAltoCadaElemento() As Long
    GetAltoCadaElemento = AuxAltoTexto
End Function

Public Function GetCantidadElementosEntran() As Long
    GetCantidadElementosEntran = zAlto / AuxAltoTexto
End Function

Public Function GetElementoVisibleNro(index As Long) As clsTemaList
    Set GetElementoVisibleNro = mElemento(index)
End Function

Public Sub IniciarTouchScreen()
    Dim i As Long
    Dim aux As Long
    Dim RegAltoFuente As Long
    
    Dim X1Region As Long
    Dim Y1Region As Long
    Dim X2Region As Long
    Dim Y2Region As Long
    
    'Cuantos elementos caben en la lista?
    'primero obtengo el alto de la fuente
    qAlgunFormulario.ScaleMode = 1
    GetAnchoAltoFont qAlgunFormulario, "A", rFntName, rFntSize, rFntBold, aux, RegAltoFuente
    qAlgunFormulario.ScaleMode = 3
    
    'alto total / alto fuente = cantidad de elementos que entran
    aux = zAlto / RegAltoFuente
    
    'creo las regiones
    ReDim IdRegiones(0)
    For i = 1 To aux
        ReDim Preserve IdRegiones(i)
        
        X1Region = zX
        Y1Region = zY + (RegAltoFuente * (i - 1))
        X2Region = zAncho
        Y2Region = Y1Region + RegAltoFuente
        
        IdRegiones(i) = CreateRectRgn(X1Region, Y1Region, X2Region, Y2Region)
    Next i
End Sub

Public Function DoClick_GetElementoIndex(clX As Long, clY As Long) As Long
    Dim i As Long
    DoClick_GetElementoIndex = -1
    For i = 1 To UBound(IdRegiones)
        If PtInRegion(IdRegiones(i), clX, clY) <> 0 Then
            DoClick_GetElementoIndex = i
            DibujarTexto i
            Exit Function
        End If
    Next i
End Function
