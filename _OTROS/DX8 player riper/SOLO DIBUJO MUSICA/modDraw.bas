Attribute VB_Name = "modDraw"
'dibujar las ondas segun frecuencias o canales o como osicloscopio

Option Explicit

Public Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal Y As Long, _
    ByVal D As Long _
) As Long

Public Declare Function LineTo Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal Y As Long _
) As Long

Public Declare Function Rectangle Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long

Private Declare Function DrawText Lib "user32" _
Alias "DrawTextA" ( _
    ByVal hdc As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long _
) As Long

Public Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Const DT_BOTTOM             As Long = &H8
Private Const DT_CENTER             As Long = &H1
Private Const DT_LEFT               As Long = &H0
Private Const DT_RIGHT              As Long = &H2
Private Const DT_TOP                As Long = &H0
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORDBREAK          As Long = &H10

Public Const FFT_MAXAMPLITUDE       As Double = 50
Public Const FFT_BANDLOWER          As Double = 0.12
Public Const FFT_BANDS              As Long = 19
Public Const FFT_BANDSPACE          As Long = 1
Public Const FFT_BANDWIDTH          As Long = 3
Public Const FFT_STARTINDEX         As Long = 1
Public Const FFT_SAMPLES            As Long = 512

Public Const DRW_BARXOFF            As Long = 4
Public Const DRW_BARYOFF            As Long = 2
Public Const DRW_BARWIDTH           As Long = 4
Public Const DRW_BARSPACE           As Long = 1

Private band(FFT_BANDS - 1)         As Double
Private BarData()                   As Double

Public Sub DrawFrequencies( _
    intSamples() As Integer, _
    picVis As PictureBox _
)

    Dim sngRealOut(FFT_SAMPLES - 1) As Single
    Dim sngBand                     As Single
    Dim hBrush                      As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim intRed                      As Integer
    Dim intGreen                    As Integer
    Dim intBlue                     As Integer
    Dim rcBand                      As RECT

    ' decrease bands amplitudes
    ' creates an effect similar to Winamp
    For i = 0 To FFT_BANDS - 1
        band(i) = band(i) - FFT_BANDLOWER
        If band(i) < 0 Then band(i) = 0
    Next

    ' Fast Fourier Transformation
    ' with Hanning Window
    clsDSP.FastFourierTransform FFT_SAMPLES, _
                                intSamples, _
                                sngRealOut, _
                                WINDOW_FUNC_HANNING

    ' normalize values and cut them at maxampl
    For i = 0 To FFT_SAMPLES / 2
        sngRealOut(i) = Abs(sngRealOut(i) / Tan(1 / Sqr(i + 1))) * 0.00003
        If sngRealOut(i) > FFT_MAXAMPLITUDE Then sngRealOut(i) = FFT_MAXAMPLITUDE
        sngRealOut(i) = sngRealOut(i) / FFT_MAXAMPLITUDE
    Next

    j = FFT_STARTINDEX

    For i = 0 To FFT_BANDS - 1
        ' calculate average for the current band
        For j = j To j + FFT_BANDWIDTH
            sngBand = sngBand + sngRealOut(j)
        Next
        ' boost higher freqs, as they got less
        ' power then the lower ones
        sngBand = (sngBand / FFT_BANDWIDTH) * Cos(1 / (i + 1))
        ' if the current band is smaller then
        ' the new one, replace the old with the new
        If band(i) < sngBand Then band(i) = sngBand
        If band(i) > 1 Then band(i) = 1
        ' leave out some bands
        j = j + FFT_BANDSPACE
    Next

    ' draw bars
    picVis.Cls

    intRed = 255
    intBlue = 0

    For i = 0 To FFT_BANDS - 1
        intGreen = 255 - (band(i) * 255)

        hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))

        With rcBand
            .Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
            .Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
            .Top = Max(DRW_BARYOFF, Min(picVis.ScaleHeight, picVis.ScaleHeight - (picVis.ScaleHeight * band(i))) - DRW_BARYOFF - 1)
            .Bottom = picVis.ScaleHeight - DRW_BARYOFF
        End With
        FillRect picVis.hdc, rcBand, hBrush

        DeleteObject hBrush
    Next
End Sub

' L/R Channel Peaks
' http://www.activevb.de/tipps/vb6tipps/tipp0406.html
Public Sub DrawPeaks( _
    intSamples() As Integer, _
    picVis As PictureBox _
)

    Dim i               As Long
    Dim j               As Long
    Dim maxL            As Long
    Dim maxR            As Long
    Dim rcBand          As RECT
    Dim rcText          As RECT

    Dim strDB           As String

    Dim intRed          As Integer
    Dim intGreen        As Integer
    Dim intBlue         As Integer

    Dim hBrush          As Long

    Static LastL        As Long
    Static LastR        As Long

    ' mono?
    If clsStream.Info.channels = 1 Then

        For i = LBound(intSamples) To UBound(intSamples)
            If Abs(CLng(intSamples(i))) > maxL Then
                maxL = Abs(CLng(intSamples(i)))
                maxR = maxL
            End If
        Next

    ' stereo!
    Else

        For i = LBound(intSamples) To UBound(intSamples)
            If i Mod 2 Then
                If Abs(CLng(intSamples(i))) > maxR Then
                    maxR = Abs(CLng(intSamples(i)))
                End If
            Else
                If Abs(CLng(intSamples(i))) > maxL Then
                    maxL = Abs(CLng(intSamples(i)))
                End If
            End If
        Next

    End If

    ' smoother value
    maxL = (LastL + maxL) / 2
    maxR = (LastR + maxR) / 2

    '*********************************************
    ' draw bars
    '*********************************************

    picVis.Cls

    intRed = 255
    intBlue = 0
    intGreen = 255 - (maxL / 32767 * 255)

    hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))

    ' Peak for left channel
    With rcBand
        .Right = maxL / 40000 * picVis.ScaleWidth - DRW_BARXOFF * 2
        .Left = DRW_BARXOFF
        .Top = DRW_BARYOFF * 2
        .Bottom = .Top + 6
    End With
    FillRect picVis.hdc, rcBand, hBrush

    DeleteObject hBrush

    '*********************************************
    '*********************************************

    intRed = 255
    intBlue = 0
    intGreen = 255 - (maxR / 32767 * 255)

    hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))

    ' Peak for right channel
    With rcBand
        .Right = maxR / 40000 * picVis.ScaleWidth - DRW_BARXOFF * 2
        .Left = DRW_BARXOFF
        .Top = .Bottom + 4
        .Bottom = .Top + 6
    End With
    FillRect picVis.hdc, rcBand, hBrush

    DeleteObject hBrush

    '*********************************************
    ' dBFS display
    '*********************************************

    ' font color
    picVis.ForeColor = vbWhite
    
    '*********************************************
    '*********************************************
    ' left ch

    strDB = Fix(dBFS(maxL)) & " dB"

    With rcText
        .Left = picVis.ScaleWidth - picVis.TextWidth(strDB)
        .Right = picVis.ScaleWidth
        .Top = DRW_BARYOFF - 1
        .Bottom = .Top + picVis.TextHeight(strDB)
    End With

    DrawText picVis.hdc, strDB, Len(strDB), rcText, DT_CENTER

    '*********************************************
    '*********************************************
    ' right ch

    strDB = Fix(dBFS(maxR)) & " dB"

    With rcText
        .Left = picVis.ScaleWidth - picVis.TextWidth(strDB)
        .Right = picVis.ScaleWidth
        .Top = .Bottom
        .Bottom = .Top + picVis.TextHeight(strDB)
    End With

    DrawText picVis.hdc, strDB, Len(strDB), rcText, DT_CENTER

    '*********************************************
    '*********************************************

    LastL = maxL
    LastR = maxR
End Sub

' Oscilloscope
Public Sub DrawOsc( _
    Data() As Integer, _
    picVis As PictureBox _
)

    Dim dx              As Long, dy         As Long
    Dim x               As Long
    Dim dy2             As Long
    Dim dc0             As Long
    Dim j               As Long, k          As Long
    Dim lngPoints       As Long
    Dim lngMaxAmpl      As Long
    Dim lngAmpl         As Long
    Dim dblAmpl         As Double

    dx = picVis.ScaleWidth
    dy = picVis.ScaleHeight
    dy2 = dy \ 2
    dc0 = picVis.hdc

    picVis.ForeColor = vbBlack
    Rectangle dc0, 0, 0, dx, dy
    picVis.ForeColor = RGB(255, 255, 113)
    MoveToEx dc0, 0, dy2, 0

    For x = 0 To UBound(Data)
        lngAmpl = Abs(CLng(Data(x)))
        If lngAmpl > lngMaxAmpl Then
            lngMaxAmpl = lngAmpl
        End If
    Next

    If lngMaxAmpl = 0 Then lngMaxAmpl = 10000

    ' points per pixel
    lngPoints = UBound(Data) / picVis.ScaleWidth

    For x = 1 To picVis.ScaleWidth - 3
        ' combine some points
        dblAmpl = 0
        'TBR
        'K NO PUEDE SER 512 O MAS!
        Dim Tope As Long
        If (k + lngPoints - 1) >= 512 Then
            Tope = 511
        Else
            Tope = k + lngPoints - 1
        End If
        For k = k To Tope
            dblAmpl = dblAmpl + Data(k)
        Next

        ' normalize the point
        dblAmpl = (dblAmpl / lngPoints) / lngMaxAmpl
        If dblAmpl > 1 Then
            dblAmpl = 1
        ElseIf dblAmpl < -1 Then
            dblAmpl = -1
        End If

        LineTo dc0, x, dblAmpl * (dy2 - 2) + dy2
    Next

    LineTo dc0, x + 0, dy2
    LineTo dc0, x + 1, dy2
End Sub
