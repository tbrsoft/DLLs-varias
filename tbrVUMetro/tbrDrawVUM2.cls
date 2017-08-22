VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrDrawVUM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'mueve el puntero grafico a algun lugar
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As Any) As Long
'linea desde el puntero grafico hasta el indicado
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Private Const FFT_MAXAMPLITUDE       As Double = 0.2
Private Const FFT_BANDLOWER          As Double = 0.07
Private Const FFT_BANDS              As Long = 22
Private Const FFT_BANDSPACE          As Long = 1
Private Const FFT_BANDWIDTH          As Long = 3
Private Const FFT_STARTINDEX         As Long = 1
Private Const FFT_SAMPLES            As Long = 1024

Private Const DRW_BARXOFF            As Long = 4
Private Const DRW_BARYOFF            As Long = 2
Private Const DRW_BARWIDTH           As Long = 3
Private Const DRW_BARSPACE           As Long = 1

Public Enum ModoVum
    DesdeArriba = 0
    DesdeAbajo = 1
    Estereo = 2
    TresColoresEstereo = 3
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

Private p As Object 'lugar donde todo se pinta
Private P2 As Object 'segundo picture para otras cosas (otro canal)
'coordenadas que se usan dentro del picturebox
Private X1enP As Long, X2enP As Long
Private Y1enP As Long, Y2enP As Long

Private X1enP2 As Long, X2enP2 As Long
Private Y1enP2 As Long, Y2enP2 As Long

Private mCantPic As Long 'cantidad de barras para los modos simples (0 1 2)
Private mCantCuadros As Long 'cantidad de bloques en modo 3 colores estereo
Private mFramePorSeg As Long 'lecturas del reloj por segundo de las señales
Private LastD(511) As Byte

Private p1HDC As Long, p2HDC As Long
Private clsFFT As clsFourier
Private band(FFT_BANDS - 1) As Double


Public Sub NotifyResizeVUM() 'reviso los dos por las dudas
    X1enP = p.Width * 0.02: Y1enP = p.Height * 0.02
    X2enP = p.Width - X1enP: Y2enP = p.Height - Y1enP
    
    X1enP2 = P2.Width * 0.02: Y1enP2 = P2.Height * 0.02
    X2enP2 = P2.Width - X1enP2: Y2enP2 = P2.Height - Y1enP2
End Sub

Public Sub Empezar()
    V.DoStart
End Sub

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

Private Sub Class_Initialize()
    Set V = New tbrVUMETRO
    
    VerdeOFF = &H808000
    AmarilloOFF = &H8080&
    RojoOFF = &H80&
    VerdeON = &HFF00&
    AmarilloON = &HFFFF&
    RojoON = &HFF&
    mColorBase = 4444444 'RGB(0, 0, 255)
    
    BRocha(0) = CreateSolidBrush(mColorBase)
    BRocha(1) = CreateSolidBrush(VerdeOFF)
    BRocha(2) = CreateSolidBrush(AmarilloOFF)
    BRocha(3) = CreateSolidBrush(RojoOFF)
    BRocha(4) = CreateSolidBrush(VerdeON)
    BRocha(5) = CreateSolidBrush(AmarilloON)
    BRocha(6) = CreateSolidBrush(RojoON)
    
    mFramePorSeg = 10
    mCantPic = 15
    Modo = Estereo
    
    V.SetIntervalLectura (1000 / mFramePorSeg)
    
End Sub

Public Sub DrawFrequencies(intSamples() As Integer, picbox As PictureBox)

    Dim sngRealOut(FFT_SAMPLES - 1) As Single
    Dim sngBand                     As Single
    Dim hBrush                      As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim intRed                      As Integer
    Dim intGreen                    As Integer
    Dim intBlue                     As Integer
    Dim rcBand                      As RECT

    If UBound(intSamples) < FFT_SAMPLES - 1 Then Exit Sub

    'se define solo una vez
    If clsFFT Is Nothing Then
        Set clsFFT = New clsFourier
        clsFFT.NumberOfSamples = FFT_SAMPLES
        clsFFT.WithTimeWindow = 1
    End If

    'paso todos los valores al furier
    For i = 0 To FFT_SAMPLES - 1
        clsFFT.RealIn(i + 1) = intSamples(i)
    Next

    ' lower band amplitudes
    For i = 0 To FFT_BANDS - 1
        band(i) = band(i) - FFT_BANDLOWER
        If band(i) < 0 Then band(i) = 0
    Next

    ' normalize values and cut them at FFT_MAXAMPLITUDE
    For i = 0 To FFT_SAMPLES / 2
        ' Ausgabe auf [0;1] normalisieren
        sngRealOut(i) = clsFFT.ComplexOut(i + 1) / (FFT_SAMPLES / 4) / 32767

        ' cut the output to FFT_MAXAMPLITUDE, so
        ' the spectrum doesn't get too small
        If sngRealOut(i) > FFT_MAXAMPLITUDE Then
            sngRealOut(i) = FFT_MAXAMPLITUDE
        End If

        sngRealOut(i) = sngRealOut(i) / FFT_MAXAMPLITUDE
    Next

    j = FFT_STARTINDEX

    For i = 0 To FFT_BANDS - 1
        ' average for the current band
        For j = j To j + FFT_BANDWIDTH
            sngBand = sngBand + sngRealOut(j)
        Next

        ' boost frequencies in the middle with a hanning window,
        ' because they got less power then the low ones
        sngBand = (sngBand * (Hanning(i + 3, FFT_BANDS + 3) + 1)) / FFT_BANDWIDTH

        If band(i) < sngBand Then band(i) = sngBand
        If band(i) > 1 Then band(i) = 1

        ' skip some bands
        j = j + FFT_BANDSPACE
    Next

    ' draw bars
    picbox.Cls

    intRed = 255
    intBlue = 50

    For i = 0 To FFT_BANDS - 1
        intGreen = (band(i) * 255)

        hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))

        With rcBand
            .Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
            .Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
            .Top = max(DRW_BARYOFF, Min(picbox.ScaleHeight, picbox.ScaleHeight - (picbox.ScaleHeight * band(i))) - DRW_BARYOFF) ' - 1)
            .Bottom = picbox.ScaleHeight - DRW_BARYOFF
        End With
        FillRect picbox.hdc, rcBand, hBrush

        DeleteObject hBrush
    Next
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

Private Sub DrawBarra3PM(PP As Object, Porc As Long)

    'le digo el porcentaje y la coordenada X
    
    Dim x As Long, X2 As Long, Y As Long, Y2 As Long
    If PP.hWnd = p.hWnd Then x = X1enP: X2 = X2enP: Y = Y1enP: Y2 = Y2enP
    If PP.hWnd = P2.hWnd Then x = X1enP2: X2 = X2enP2: Y = Y1enP2: Y2 = Y2enP2
    
    Dim AltoBarra2 As Long
    AltoBarra2 = ((Y2 - Y) / 15) / mCantCuadros
    
    'pintar una por una del color que corresponda
    
    Dim tmpBr As Long
    Dim PorcEsteCuadro As Long, a As Long
    For a = 1 To mCantCuadros
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
        
        'sDraw PP, x / 15, (Y2 / 15) - ((a + 1) * AltoBarra2), X2 / 15, _
             (Y2 / 15) - ((a) * AltoBarra2), _
            tmpBr
    Next a
End Sub

Private Sub V_Pintar(D() As Byte)
    DrawFrequencies D(), p
End Sub

Public Sub DefinePictureBox(pic As Object)
    Set p = pic
    p.AutoRedraw = True
    '10 % margen en todos lados
    X1enP = p.Width * 0.02: Y1enP = p.Height * 0.02
    X2enP = p.Width - X1enP: Y2enP = p.Height - Y1enP
    
    p.ForeColor = mColorBase
    p1HDC = p.hdc
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

Public Property Get ColorBase() As Long
    ColorBase = mColorBase
End Property

Public Property Let ColorBase(NewColor As Long)
    mColorBase = NewColor
    'el borde del mismo color
    DeleteObject BRocha(0)
    BRocha(0) = CreateSolidBrush(mColorBase)
    
    p.ForeColor = NewColor
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

' boost frequencies in the middle with a hanning window,
        ' because they got less power then the low ones
Private Function Hanning(ByVal x As Single, ByVal Length As Long) As Single
    Hanning = 0.5 * (1 - Cos((2 * Pi * x) / Length))
End Function

Private Function Min(ByVal val1 As Long, ByVal val2 As Long) As Long
    Min = IIf(val1 < val2, val1, val2)
End Function

Private Function max(ByVal val1 As Long, ByVal val2 As Long) As Long
    max = IIf(val1 > val2, val1, val2)
End Function


