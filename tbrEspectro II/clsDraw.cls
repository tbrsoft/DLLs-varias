VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' draws a frequency spectrum (Winamp style)
' and a amplitude curve

Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long

Private Declare Function MoveToEx Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal D As Long _
) As Long

Private Declare Function LineTo Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Private Declare Function Rectangle Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long

Private Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Private Const Pi                    As Single = 3.14159265358979

'Private band(FFT_BANDS - 1)          As Double
Private band() As Double

Private exband() As Double

Private clsFFT                      As clsFourier

Public Sub DrawRect( _
    ByVal hdc As Long, _
    ByVal lngLeft As Long, _
    ByVal lngTop As Long, _
    ByVal lngRight As Long, _
    ByVal lngBottom As Long, _
    ByVal lngColor As Long _
)

    Dim udtRect As RECT
    Dim hBrush  As Long

    With udtRect
        .Left = lngLeft
        .Right = lngRight
        .Top = lngTop
        .Bottom = lngBottom
    End With

    hBrush = CreateSolidBrush(lngColor)
    FillRect hdc, udtRect, hBrush
    DeleteObject hBrush
End Sub

Public Sub DrawAmplitudes(data() As Integer, picVis As PictureBox, Coloraso As Long, ColorasoFondo As Long)
    
    ReDim band(FFT_BANDS - 1)
    
    Dim dx              As Long, dy         As Long
    Dim X               As Long, K          As Long
    Dim dy2             As Long
    Dim dc0             As Long
    Dim lngPoints       As Long
    Dim lngMaxAmpl      As Long, lngAmpl    As Long
    Dim dblAmpl         As Double

    dx = picVis.ScaleWidth
    dy = picVis.ScaleHeight
    dy2 = dy \ 2
    dc0 = picVis.hdc

    'picVis.ForeColor = vbBlack
    picVis.ForeColor = ColorasoFondo
    Rectangle dc0, 0, 0, dx, dy
    'picVis.ForeColor = vbWhite
    picVis.ForeColor = Coloraso
    MoveToEx dc0, 0, dy2, 0

    For X = 0 To UBound(data)
        lngAmpl = Abs(CLng(data(X)))
        If lngAmpl > lngMaxAmpl Then
            lngMaxAmpl = lngAmpl
        End If
    Next

    'If lngMaxAmpl = 0 Then lngMaxAmpl = 32767
    lngMaxAmpl = 32767

    ' points per pixel
    lngPoints = UBound(data) / picVis.ScaleWidth

    For X = 1 To picVis.ScaleWidth - 3
        ' average of some points
        dblAmpl = 0
        For K = K To K + lngPoints - 1
            'XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
            'se pasa de data(k) en modo 2
            On Local Error Resume Next
            dblAmpl = dblAmpl + data(K)
        Next

        ' normalize points
        dblAmpl = (dblAmpl / lngPoints) / lngMaxAmpl
        If dblAmpl > 1 Then
            dblAmpl = 1
        ElseIf dblAmpl < -1 Then
            dblAmpl = -1
        End If

        ' draw a line to the new point
        LineTo dc0, X, dblAmpl * (dy2 - 2) + dy2
    Next

    ' return to the middle
    LineTo dc0, X + 0, dy2
    LineTo dc0, X + 1, dy2
End Sub

Public Sub DrawFrequencies(intSamples() As Integer, picbox As PictureBox, _
    Coloraso As Long, DivisorMagico As Double, ElCls As Boolean)
    
    If DRW_BARWIDTH < 1 Then DRW_BARWIDTH = 3
    
    'SE LO SAQUE YO ANDRES PARA QUE PERMITA 0,50. Desconozco por que estaba
    'If DivisorMagico < 1 Then DivisorMagico = 1
    
    ReDim band(FFT_BANDS - 1)
    
    Dim sngRealOut(FFT_SAMPLES - 1) As Single
    Dim sngBand                     As Single
    Dim hBrush                      As Long
    Dim i                           As Long
    Dim j                           As Long
    'Dim intRed                      As Integer
    'Dim intGreen                    As Integer
    'Dim intBlue                     As Integer
    Dim rcBand                      As RECT

    If UBound(intSamples) < FFT_SAMPLES - 1 Then Exit Sub

    'se define solo una vez
    If clsFFT Is Nothing Then
        Set clsFFT = New clsFourier
        clsFFT.NumberOfSamples = FFT_SAMPLES
        clsFFT.WithTimeWindow = 1
    End If

    For i = 0 To FFT_SAMPLES - 1
        clsFFT.RealIn(i + 1) = intSamples(i)
    Next

    'METI MANO
    ReDim exband(UBound(band))
    For i = 0 To UBound(band)
        exband(i) = band(i)
    Next
    '----------------------------
    
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
    If ElCls = True Then
        picbox.Cls
    Else
        picbox.Refresh
    End If

    'intGreen = 125
    'intRed = 125
    'intBlue = 250

    Dim Promedio As Single
    'Dim e As Long
    For i = 0 To FFT_BANDS - 1
        'intGreen = (band(i) * 255)

        'hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))
        hBrush = CreateSolidBrush(Coloraso)

'        With rcBand
'            .Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
'            .Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
'            .Top = max(DRW_BARYOFF, Min(picbox.ScaleHeight, picbox.ScaleHeight - (picbox.ScaleHeight * band(i))) - (DRW_BARYOFF)) ' - 1)
'            '.Top = max(DRW_BARYOFF, Min(picbox.ScaleHeight, picbox.ScaleHeight - (picbox.ScaleHeight * (band(i) * DivisorMagico))) - (DRW_BARYOFF)) ' - 1)
'            .Bottom = (picbox.ScaleHeight - DRW_BARYOFF)
'        End With
        
        
        'If UBound(band) = UBound(exband) Then
        '    Promedio = ((band(3) + band(i) + exband(i)) / 3)
        'Else
        '    Promedio = ((band(3) + band(i)) / 2)
        'End If
        If i > UBound(band) - 4 Then
            Promedio = (band(i) + exband(i) + band(5)) / 3
        Else
            Promedio = (band(i + 4) + exband(i + 4) + band(5)) / 3
        End If
        'Promedio = Log(Promedio * 10)
        
        With rcBand
            .Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
            .Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
            '.Top = max(DRW_BARYOFF, Min(picbox.ScaleHeight, picbox.ScaleHeight - (picbox.ScaleHeight * (band(i) * (i * 2)))) - (DRW_BARYOFF)) ' - 1)
            .Top = max(DRW_BARYOFF, Min(picbox.ScaleHeight, picbox.ScaleHeight - (picbox.ScaleHeight * (Promedio * DivisorMagico))) - (DRW_BARYOFF)) ' - 1)
            .Bottom = (picbox.ScaleHeight - DRW_BARYOFF)
        End With
        
        FillRect picbox.hdc, rcBand, hBrush
        
        DeleteObject hBrush
    Next
End Sub

Private Function Hanning( _
    ByVal X As Single, _
    ByVal Length As Long _
) As Single

    Hanning = 0.5 * (1 - Cos((2 * Pi * X) / Length))
End Function

Private Function Min( _
    ByVal val1 As Long, _
    ByVal val2 As Long _
) As Long

    Min = IIf(val1 < val2, val1, val2)
End Function

Private Function max( _
    ByVal val1 As Long, _
    ByVal val2 As Long _
) As Long

    max = IIf(val1 > val2, val1, val2)
End Function

