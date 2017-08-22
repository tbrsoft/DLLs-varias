VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrDrawVUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Event VerError(Descr As String)

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'mueve el puntero grafico a algun lugar
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
'linea desde el puntero grafico hasta el indicado
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Enum ModoVum
    DesdeArriba = 0
    DesdeAbajo = 1
    Estereo = 2
    TresColoresEstereo = 3 'p1 es un canal y p2 el otro
    UnaImagenSobreOtra = 4 'usa p2 sobre p1 y p3 sobre p4
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private R1 As RECT

Dim WithEvents V As tbrVUMETRO
Attribute V.VB_VarHelpID = -1

Private BRocha(10) As Long 'brocha para dibujar actual

'lo saco
'Private LastBrocha As Long 'id de la brocha para borrarla de memoria

Private LastBrochaCol As Long 'color de la ultima (para ver si necsita de nuevo el mismo color y no cargar una nueva brocha)
Private Modo As ModoVum 'modo de dibujo

'colres a usar en vumetros de colores
Private VerdeOFF As Long, AmarilloOFF As Long, RojoOFF As Long
Private VerdeON As Long, AmarilloON As Long, RojoON As Long
Private mColorBase As Long 'color de las barras en modos simples

Private P As Object 'lugar donde todo se pinta
Private P2 As Object 'segundo picture para otras cosas (otro canal)
Private P3 As Object
Private P4 As Object

'coordenadas que se usan dentro del picturebox
Private X1enP As Long, X2enP As Long
Private Y1enP As Long, Y2enP As Long

Private X1enP2 As Long, X2enP2 As Long
Private Y1enP2 As Long, Y2enP2 As Long

Private X1enP3 As Long, X2enP3 As Long
Private Y1enP3 As Long, Y2enP3 As Long

Private X1enP4 As Long, X2enP4 As Long
Private Y1enP4 As Long, Y2enP4 As Long


Private mCantPic As Long 'cantidad de barras para los modos simples (0 1 2)
Private mCantCuadros As Long 'cantidad de bloques en modo 3 colores estereo
Private mFramePorSeg As Long 'lecturas del reloj por segundo de las señales

'ultimos x valores de cada canal para que los cambios no sean bruscos y el movimiento sea armonico
Private LastValrR() As Byte
Private LastValrL() As Byte

Private p1HDC As Long, p2HDC As Long, p3HDC As Long, p4HDC As Long

Private mEnabled As Boolean
Private mCarFantastic As Boolean

'ver las proporciones del destino

Dim AnchoBarra As Single 'barras en modos 0, 1, 2 en pixeles
Dim AnchoBarra2 As Single 'barras en modos 0, 1, 2 en pixeles
Dim AltoBarra As Long 'barras en modos 0, 1, 2 en pixeles
Dim AltoBarra2 As Long 'barras en modos 0, 1, 2 en pixeles

Dim ThisPT_L As Long 'valor usado para cada canal
Dim ThisPT_R As Long
    
Public Sub DoPause(Continuar As Boolean)
    V.DoPause Continuar
End Sub

Public Property Get DispositivosCant() As Long
    DispositivosCant = V.Dispositivos
End Property

Public Property Get NombrePlacas(i As Long)
    NombrePlacas = V.NombreDispositivo(i)
End Property

Public Sub NotifyResizeVUM() 'reviso los dos por las dudas
    X1enP = P.Width * 0.02: Y1enP = P.Height * 0.02
    X2enP = P.Width - X1enP: Y2enP = P.Height - Y1enP
    
    X1enP2 = P2.Width * 0.02: Y1enP2 = P2.Height * 0.02
    X2enP2 = P2.Width - X1enP2: Y2enP2 = P2.Height - Y1enP2
    
    'ver las proporciones del destino
    AnchoBarra = (CSng(X2enP - X1enP) / 15) / mCantPic
    AnchoBarra2 = (CSng(X2enP2 - X1enP2) / 15) / mCantPic
    AltoBarra = (Y2enP - Y1enP) / 15
    AltoBarra2 = (Y2enP2 - Y1enP2) / 15
End Sub

Public Function Empezar() As Long
    'cero ok
    'uno mal!!!!

    CrearBrochas
    Empezar = V.DoStart

End Function

Public Sub Terminar()
    V.DoStop
    
    DeleteObject BRocha(0)
    DeleteObject BRocha(1)
    DeleteObject BRocha(2)
    DeleteObject BRocha(3)
    DeleteObject BRocha(4)
    DeleteObject BRocha(5)
    DeleteObject BRocha(6)
    
End Sub

Public Property Get IsPlaying() As Boolean
    IsPlaying = V.IsPlaying
End Property

Private Sub Class_Initialize()
    'On Local Error GoTo ErrIniVU
    
    Set V = New tbrVUMETRO
    
    ReDim LastValrR(2) 'armoniza con 11 valores (no en los modos con multiples barras)
    ReDim LastValrL(2) 'todos los valores quedan en cero
    
    'If V.Dispositivos = 0 Then RaiseEvent VerError("Sin Dispositivos")
    
    VerdeOFF = &H808000
    AmarilloOFF = &H8080&
    RojoOFF = &H80&
    VerdeON = &HFF00&
    AmarilloON = &HFFFF&
    RojoON = &HFF&
    mColorBase = RGB(0, 0, 255)
    
    CrearBrochas
    
    mFramePorSeg = 10
    mCantPic = 15
    mCantCuadros = 20
    Modo = Estereo
    
    V.SetIntervalLectura (1000 / mFramePorSeg)
    
    Exit Sub
ErrIniVU:
    'MsgBox Err.Description + " " + CStr(Err.Number)
    Resume Next
End Sub

Private Sub CrearBrochas()
    BRocha(0) = CreateSolidBrush(mColorBase)
    BRocha(1) = CreateSolidBrush(VerdeOFF)
    BRocha(2) = CreateSolidBrush(AmarilloOFF)
    BRocha(3) = CreateSolidBrush(RojoOFF)
    BRocha(4) = CreateSolidBrush(VerdeON)
    BRocha(5) = CreateSolidBrush(AmarilloON)
    BRocha(6) = CreateSolidBrush(RojoON)
End Sub

Private Sub sDraw(PP As Object, _
                  L As Long, t As Long, R As Long, B As Long, _
                  indexBR As Long)
    
    'BRocha(indexBR) = CreateSolidBrush(mColorBase)
    SelectObject PP.hdc, BRocha(indexBR)
'    RoundRect PP.hdc, L, T, R, B, 2, 2

    Dim rcBand As RECT
    rcBand.Left = L
    rcBand.Top = t
    rcBand.Bottom = B
    rcBand.Right = R

    FillRect PP.hdc, rcBand, BRocha(indexBR)
'    DeleteObject BRocha(indexBR)
    
End Sub

Public Property Let CantCuadros(newCant As Long)
    mCantCuadros = newCant
End Property

Public Property Get CantCuadros() As Long
    CantCuadros = mCantCuadros
End Property

Public Property Let CantPic(newCant As Long)
    mCantPic = newCant
End Property

Public Property Get CantPic() As Long
    CantPic = mCantPic
End Property

Public Property Let ModoVumetro(newModo As ModoVum)
    Modo = newModo
    'en este modo los bordes negro mejor
    If Modo = TresColoresEstereo Then ColorBase = vbBlack
End Property

Public Property Get ModoVumetro() As ModoVum
    ModoVumetro = Modo
End Property

Private Sub DrawBarra3PM(PB As Object, X As Long, X2 As Long, Y As Long, Y2 As Long, Porc As Long)
    
    Dim AltoBarra2b As Long
    AltoBarra2b = ((Y2 - Y) / 15) / mCantCuadros
    
    'pintar una por una del color que corresponda
    
    Dim tmpBr As Long
    Dim PorcEsteCuadro As Long, a As Long
    For a = 0 To mCantCuadros - 1
        PorcEsteCuadro = (100 / mCantCuadros) * a
        If PorcEsteCuadro > 80 Then
            If PorcEsteCuadro > Porc Then
                tmpBr = 3 'RojoOFF
            Else
                tmpBr = 6 'RojoON
            End If
        Else
            If PorcEsteCuadro > 20 Then
                If PorcEsteCuadro > Porc Then
                    tmpBr = 2 'AmarilloOFF
                Else
                    tmpBr = 5 'AmarilloON
                End If
            Else
                If PorcEsteCuadro > Porc Then
                    tmpBr = 1 'VerdeOFF
                Else
                    tmpBr = 4 'VerdeON
                End If
            End If
        End If
        'pintar lo elegido
        
        sDraw PB, X / 15, (Y2 / 15) - ((a + 1) * AltoBarra2b), X2 / 15, _
             ((Y2 / 15) - ((a) * AltoBarra2b) - 1), _
            tmpBr
    Next a
End Sub

Private Sub V_Pintar(D() As Byte)
    
    Dim a As Long
    'los pares son un canal e impares otros
    
    Dim C As Long: C = 0 'contador ya que no entra en los 512 valores!
    
    'coordenada de cada uno de los parametro de la funcion para dibujar
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    'los valores de la matriz byte son en general entre 128 y 256 pero a veces son menores a 128 ??
    'entonces tomo como cero el 128 y los menores a este los considereo positivos
    Dim MaxThis As Long, Descontar As Long
    Descontar = 128
    MaxThis = 256 - Descontar  'el sonido cero es 120
    'borro todo en cada vuelta
    

    'en los modos 0 1 y 2 debo leer mas de una vez segun la cantiad de pics
    If Modo = DesdeAbajo Or Modo = DesdeArriba Or Modo = Estereo Then
        P.Cls: P2.Cls
         'en otros modos puede limpiar menos
        For a = 0 To 510 Step CLng(512 / mCantPic) + 1 'que sea un poco de todo el espectro de valores
            
            'A TIENE QUE SER PAR PARA RESPETAR LOS CANALES!!
            If a / 2 <> a \ 2 Then
                If a < 510 Then
                    a = a + 1
                Else
                    a = 510
                End If
            End If
            'armonizar con el anterior
            ThisPT_L = Abs(D(a) - Descontar)
            ThisPT_R = Abs(D(a + 1) - Descontar)
            
            If ThisPT_L > 255 Then ThisPT_L = 255
            If ThisPT_R > 255 Then ThisPT_R = 255
            
            'asegurar que las divisiones con decimal no se vayan a cero
            Dim TMPsng As Single
            Select Case Modo
                Case 0 'de arriba para abajo
                    X1 = (C * AnchoBarra)
                    Y1 = 0
                    X2 = (C * AnchoBarra) + AnchoBarra
                    TMPsng = (ThisPT_L / MaxThis) * AltoBarra
                    Y2 = CLng(TMPsng)
                    
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_L, 50))
                    sDraw P, (X1enP / 15) + X1, (Y1enP / 15) + Y1, _
                          (X1enP / 15) + X2, (Y1enP / 15) + Y2, 0
                          
                    X1 = (C * AnchoBarra2)
                    Y1 = 0
                    X2 = (C * AnchoBarra2) + AnchoBarra2
                    TMPsng = (ThisPT_R / MaxThis) * AltoBarra2
                    Y2 = CLng(TMPsng)
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_R, 50))
                    sDraw P2, (X1enP2 / 15) + X1, (Y1enP2 / 15) + Y1, _
                          (X1enP2 / 15) + X2, (Y1enP2 / 15) + Y2, 0 ' mColorBase
                          
                Case 1 'desde abajo
                    X1 = (C * AnchoBarra)
                    TMPsng = AltoBarra - _
                        (ThisPT_L / MaxThis) * AltoBarra
                    Y1 = CLng(TMPsng)
                    X2 = (C * AnchoBarra) + AnchoBarra
                    Y2 = AltoBarra
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_L, 50))
                    sDraw P, (X1enP / 15) + X1, (Y1enP / 15) + Y1, _
                          (X1enP / 15) + X2, (Y1enP / 15) + Y2, 0 ' mColorBase
                          
                    X1 = (C * AnchoBarra2)
                    TMPsng = AltoBarra2 - _
                        ((ThisPT_R) / (MaxThis) * AltoBarra2)
                    Y1 = CLng(TMPsng)
                    X2 = (C * AnchoBarra2) + AnchoBarra2
                    Y2 = AltoBarra2
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_R, 50))
                    sDraw P2, (X1enP2 / 15) + X1, (Y1enP2 / 15) + Y1, _
                          (X1enP2 / 15) + X2, (Y1enP2 / 15) + Y2, 0 'mColorBase
                          
                Case 2 'desde el medio en estereo
                    X1 = (C * AnchoBarra)
                    Y1 = (AltoBarra / 2)
                    X2 = (C * AnchoBarra) + AnchoBarra ' - 2
                    TMPsng = (AltoBarra / 2) + _
                        (CSng(ThisPT_L / MaxThis) * CSng((AltoBarra / 2)))
                    Y2 = CLng(TMPsng)
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_L, 50))
                    sDraw P, _
                        (X1enP / 15) + X1, (Y1enP / 15) + Y1, _
                        (X1enP / 15) + X2, (Y1enP / 15) + Y2, _
                        0 'mColorBase
                    
                    X1 = (C * AnchoBarra)
                    Y1 = AltoBarra / 2
                    X2 = (C * AnchoBarra) + AnchoBarra
                    TMPsng = CSng(AltoBarra / 2) - _
                        (CSng(ThisPT_R / MaxThis) * CSng((AltoBarra2 / 2)))
                    Y2 = CLng(TMPsng)
                    'BRocha(0) = CreateSolidBrush(RGB(255, ThisPT_R, 50))
                    sDraw P, _
                        (X1enP / 15) + X1, (Y1enP / 15) + Y1, _
                        (X1enP / 15) + X2, (Y1enP / 15) + Y2, _
                        0 'mColorBase
                
            End Select
            
            C = C + 1
        Next a
    End If
    
    If Modo = TresColoresEstereo Then
        P.Cls: P2.Cls
        
        ThisPT_R = Abs(D(150) - Descontar) * 2
        ThisPT_L = Abs(D(151) - Descontar) * 2
        
        'Armonizar ThisPT_R, ThisPT_L, ThisPT_R, ThisPT_L
        If ThisPT_R > 255 Then ThisPT_R = 255
        If ThisPT_L > 255 Then ThisPT_L = 255
        
        DrawBarra3PM P, X1enP, X2enP, Y1enP, Y2enP, (100 * ThisPT_R) / MaxThis
        DrawBarra3PM P2, X1enP2, X2enP2, Y1enP2, Y2enP2, (100 * ThisPT_L) / MaxThis
    End If
    
    If Modo = UnaImagenSobreOtra Then
        ThisPT_R = Abs(D(350) - Descontar) * 2
        ThisPT_L = Abs(D(351) - Descontar) * 2
        
        'Armonizar ThisPT_R, ThisPT_L, ThisPT_R, ThisPT_L
        If ThisPT_R > 255 Then ThisPT_R = 255
        If ThisPT_L > 255 Then ThisPT_L = 255
    
        'p1 esta por debajo, es la imagen activa del vumetro y esta atras de la p2 que es la inactiva
        ' de igual forma el 3 esta detras del 4
        'las imagenes y los tamaños ya estan definidos en el ejecutable
        'en este caso tono como referencia maxthis que es lo maximo, calculo un porcentaje y de ahi
        'se que alto hay que dibujar
        'asegurarse en el ejecutable de hacer los zorders correspondientes
        'ademas de que ambos tengan el mismo top y left, no vamos a cambiar eso 50 veces por segundo
        
        P2.Height = Abs(1 - (ThisPT_R / MaxThis)) * P.Height  'p1 y p3.height no cambian nunca!
        P4.Height = Abs(1 - (ThisPT_L / MaxThis)) * P3.Height
        
        'P2.Refresh: P4.Refresh haxe que todo titile!!!
    End If
    
End Sub

Private Sub Armonizar(AddCH_R As Long, AddCH_L As Long, RetCH_R As Long, RetCH_L As Long)
    'entran dos valores y devuelve el valor armonizado
    Dim J As Long
    'correr todos los valores ....
    For J = UBound(LastValrR) To 1 Step -1
        LastValrR(J) = LastValrR(J - 1)
        LastValrL(J) = LastValrL(J - 1)
    Next J
    
    'queda el cero libre para el nuevo valor ...
    LastValrR(0) = AddCH_R
    LastValrL(J) = AddCH_L
    
    'calcular el valor armonizado
    RetCH_R = 0: RetCH_L = 0
    For J = 0 To UBound(LastValrR)
        RetCH_R = RetCH_R + LastValrR(J)
        RetCH_L = RetCH_L + LastValrL(J)
    Next J
    
    RetCH_R = RetCH_R / (UBound(LastValrR)) 'deberia ser +1 pero lo dejo asi para que de numeros mas altos
    RetCH_L = RetCH_L / (UBound(LastValrR))
    
    If RetCH_R > 255 Then RetCH_R = 255
    If RetCH_L > 255 Then RetCH_L = 255
    
End Sub

Public Sub DefinePictureBox(pic As Object)
    Set P = pic
    P.AutoRedraw = True
    '10 % margen en todos lados
    X1enP = P.Width * 0.02: Y1enP = P.Height * 0.02
    X2enP = P.Width - X1enP: Y2enP = P.Height - Y1enP
    
    P.ForeColor = mColorBase
    p1HDC = P.hdc
End Sub

Public Sub DefinePictureBox2(pic As Object)
    Set P2 = pic
    P2.AutoRedraw = True
    '10 % margen en todos lados
    X1enP2 = P2.Width * 0.02: Y1enP2 = P2.Height * 0.02
    X2enP2 = P2.Width - X1enP2: Y2enP2 = P2.Height - Y1enP2
    
    P2.ForeColor = mColorBase
    p2HDC = P2.hdc
End Sub

Public Sub DefinePictureBox3(pic As Object)
    Set P3 = pic
    P3.AutoRedraw = True
    '10 % margen en todos lados
    X1enP3 = P3.Width * 0.02: Y1enP3 = P3.Height * 0.02
    X2enP3 = P3.Width - X1enP3: Y2enP3 = P3.Height - Y1enP3
    
    P3.ForeColor = mColorBase
    p3HDC = P3.hdc
End Sub

Public Sub DefineImage(nPic As Long, ArchImg As String, aSIZE As Boolean)
    Dim pMil As Object
    
    If nPic = 1 Then Set pMil = P
    If nPic = 2 Then Set pMil = P2
    If nPic = 3 Then Set pMil = P3
    If nPic = 4 Then Set pMil = P4
    
    If aSIZE Then pMil.AutoSize = True
    pMil.Picture = LoadPicture(ArchImg)
    If aSIZE Then pMil.AutoSize = False
    
    pMil.Cls
End Sub

Public Sub DefinePictureBox4(pic As Object)
    Set P4 = pic
    P4.AutoRedraw = True
    '10 % margen en todos lados
    X1enP4 = P4.Width * 0.02: Y1enP4 = P4.Height * 0.02
    X2enP4 = P4.Width - X1enP4: Y2enP4 = P4.Height - Y1enP4
    
    P4.ForeColor = mColorBase
    p4HDC = P4.hdc
End Sub

Public Property Get ColorBase() As Long
    ColorBase = mColorBase
End Property

Public Property Let ColorBase(NewColor As Long)
    mColorBase = NewColor
    'el borde del mismo color
    DeleteObject BRocha(0)
    BRocha(0) = CreateSolidBrush(mColorBase)
    
    P.ForeColor = NewColor
    P2.ForeColor = NewColor
End Property

Public Property Get FramePorSeg() As Long
    FramePorSeg = mFramePorSeg
End Property

Public Property Let FramePorSeg(newFrames As Long)
    If newFrames > 200 Then newFrames = 200
    
    mFramePorSeg = newFrames
    V.SetIntervalLectura (1000 / mFramePorSeg)
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(newEnabled As Boolean)
    mEnabled = newEnabled
End Property

Public Property Get CarFantastic() As Boolean
    CarFantastic = mCarFantastic
End Property

Public Property Let CarFantastic(newCarFantastic As Boolean)
    mCarFantastic = newCarFantastic
End Property
