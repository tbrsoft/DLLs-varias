VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SignalProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 16 Bit DSP functions

Private Const Pi        As Single = 3.14159265358979
Private Const LN2       As Single = 0.693147180559945
Private Const AngleNum  As Single = 2 * Pi

Private lngPow2(31)     As Long

Public Enum FFT_WINDOW
    WINDOW_FUNC_NONE
    WINDOW_FUNC_HANNING
    WINDOW_FUNC_HAMMING
    WINDOW_FUNC_BLACKMAN
End Enum

Public Enum VOL_UNIT
    VOL_DECIBEL
    VOL_PERCENT
    VOL_FACTOR
End Enum

Public Enum IIR_TYPE
    IIR_LOW_PASS
    IIR_HIGH_PASS
    IIR_BAND_PASS
    IIR_NOTCH
    IIR_PEAK_EQ
    IIR_LOW_SHELVE
    IIR_HIGH_SHELVE
End Enum

' http://www.fullspectrum.com/deeth/main.html
Public Sub FastFourierTransform( _
    NumSamples As Long, _
    RealIn() As Integer, _
    RealOut() As Single, _
    Optional wnd As FFT_WINDOW = WINDOW_FUNC_NONE _
)

    Dim NumBits     As Long

    Dim Rev         As Long
    Dim index       As Long

    Dim i           As Long, j          As Long
    Dim k           As Long, n          As Long

    Dim BlockSize   As Long, BlockEnd   As Long

    Dim DeltaAngle  As Single, DeltaAr  As Single
    Dim Alpha       As Single, Beta     As Single

    Dim TR          As Single, TI       As Single
    Dim AR          As Single, AI       As Single

    Dim dblWnd()    As Single

    Dim imagout()   As Single
    ReDim imagout(NumSamples - 1) As Single

    dblWnd = CreateWindow(wnd, NumSamples)

    For i = 0 To 16
        If (NumSamples And lngPow2(i)) <> 0 Then
            NumBits = i
            Exit For
        End If
    Next

    For i = 0 To (NumSamples - 1)
        index = i
        Rev = 0
        For k = 0 To NumBits - 1
            Rev = (Rev * 2) Or (index And 1)
            index = index \ 2
        Next
        j = Rev

        RealOut(j) = RealIn(i) * dblWnd(i)
    Next

    BlockEnd = 1
    BlockSize = 2

    Do While BlockSize <= NumSamples
        DeltaAngle = AngleNum / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2# * Alpha * Alpha
        Beta = Sin(DeltaAngle)

        For i = 0 To NumSamples - 1 Step BlockSize
            AR = 1#
            AI = 0#

            j = i
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * imagout(k)
                TI = AI * RealOut(k) + AR * imagout(k)
                RealOut(k) = RealOut(j) - TR
                imagout(k) = imagout(j) - TI
                RealOut(j) = RealOut(j) + TR
                imagout(j) = imagout(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1&
            Next

        Next

        BlockEnd = BlockSize
        BlockSize = BlockSize * 2&
    Loop
End Sub

Private Sub InitPower2()
    Dim i   As Long
    For i = 0 To 30
        lngPow2(i) = 2 ^ i
    Next
    lngPow2(31) = &H80000000
End Sub

Public Function CreateWindow( _
    wnd As FFT_WINDOW, _
    ByVal Length As Long _
) As Single()

    Dim dblOut()    As Single
    Dim i           As Long

    ReDim dblOut(Length - 1) As Single

    Select Case wnd

        Case WINDOW_FUNC_NONE
            For i = 0 To Length - 1
                dblOut(i) = 1
            Next

        Case WINDOW_FUNC_HANNING
            For i = 0 To Length - 1
                dblOut(i) = 0.5 * (1 - Cos(i * 2 * Pi / (Length - 1)))
            Next

        Case WINDOW_FUNC_HAMMING
            For i = 0 To Length - 1
                dblOut(i) = 0.54 - (0.46 * Cos((i) * 2 * Pi / (Length - 1)))
            Next

        Case WINDOW_FUNC_BLACKMAN
            For i = 0 To Length - 1
                dblOut(i) = 0.42 - (0.5 * Cos((i) * 2 * Pi / (Length - 1))) + (0.08 * Cos((i) * 4 * Pi / (Length - 1)))
            Next

    End Select

    CreateWindow = dblOut
End Function

' various Biquad IIR filters
' http://www.dspguru.com/sw/lib/biquad.c
Public Function CreateBiquadIIR( _
    ByVal iirtype As IIR_TYPE, _
    ByVal dBGain As Single, _
    ByVal freq As Single, _
    ByVal srate As Single, _
    ByVal bandwidth As Single _
) As IIRFilter

    Dim A       As Single, omega    As Single
    Dim sn      As Single, cs       As Single
    Dim Alpha   As Single, Beta     As Single

    Dim a0      As Single, a1       As Single, a2   As Single
    Dim b0      As Single, b1       As Single, b2   As Single

    Dim filter  As IIRFilter
    Set filter = New IIRFilter

    A = 10 ^ (dBGain / 40)
    omega = 2 * Pi * freq / srate
    sn = Sin(omega)
    cs = Cos(omega)
    Alpha = sn * sinh(LN2 / 2 * bandwidth * omega / sn)
    Beta = Sqr(A + A)

    Select Case iirtype
        Case IIR_LOW_PASS
            b0 = (1 - cs) / 2
            b1 = 1 - cs
            b2 = (1 - cs) / 2
            a0 = 1 + Alpha
            a1 = -2 * cs
            a2 = 1 - Alpha
        Case IIR_HIGH_PASS
            b0 = (1 + cs) / 2
            b1 = -(1 + cs)
            b2 = (1 + cs) / 2
            a0 = 1 + Alpha
            a1 = -2 * cs
            a2 = 1 - Alpha
        Case IIR_BAND_PASS
            b0 = Alpha
            b1 = 0
            b2 = -Alpha
            a0 = 1 + Alpha
            a1 = -2 * cs
            a2 = 1 - Alpha
        Case IIR_NOTCH
            b0 = 1
            b1 = -2 * cs
            b2 = 1
            a0 = 1 + Alpha
            a1 = -2 * cs
            a2 = 1 - Alpha
        Case IIR_PEAK_EQ
            b0 = 1 + (Alpha * A)
            b1 = -2 * cs
            b2 = 1 - (Alpha * A)
            a0 = 1 + (Alpha / A)
            a1 = -2 * cs
            a2 = 1 - (Alpha / A)
        Case IIR_LOW_SHELVE
            b0 = A * ((A + 1) - (A - 1) * cs + Beta * sn)
            b1 = 2 * A * ((A - 1) - (A + 1) * cs)
            b2 = A * ((A + 1) - (A - 1) * cs - Beta * sn)
            a0 = (A + 1) + (A - 1) * cs + Beta * sn
            a1 = -2 * ((A - 1) + (A + 1) * cs)
            a2 = (A + 1) + (A - 1) * cs - Beta * sn
        Case IIR_HIGH_SHELVE
            b0 = A * ((A + 1) + (A - 1) * cs + Beta * sn)
            b1 = -2 * A * ((A - 1) + (A + 1) * cs)
            b2 = A * ((A + 1) + (A - 1) * cs - Beta * sn)
            a0 = (A + 1) - (A - 1) * cs + Beta * sn
            a1 = 2 * ((A - 1) - (A + 1) * cs)
            a2 = (A + 1) - (A - 1) * cs - Beta * sn
    End Select

    filter.a0 = b0 / a0
    filter.a1 = b1 / a0
    filter.a2 = b2 / a0
    filter.a3 = a1 / a0
    filter.a4 = a2 / a0

    filter.x1 = 0: filter.x2 = 0
    filter.y1 = 0: filter.y2 = 0

    Set CreateBiquadIIR = filter
End Function

Public Sub ChangeVolume( _
    intSamples() As Integer, _
    ByVal datalen As Long, _
    ByVal value As Single, _
    ByVal unit As VOL_UNIT _
)

    Dim sngFactor   As Single
    Dim sngResult   As Single
    Dim i           As Long

    Select Case unit
        Case VOL_DECIBEL
            sngFactor = 10 ^ (value / 20)
        Case VOL_PERCENT
            sngFactor = value / 100
        Case VOL_FACTOR
            sngFactor = value
    End Select

    For i = 0& To datalen
        sngResult = intSamples(i) * sngFactor

        If sngResult > 32767# Then
            intSamples(i) = 32767
        ElseIf sngResult < -32768# Then
            intSamples(i) = -32768
        Else
            intSamples(i) = CInt(sngResult)
        End If
    Next
End Sub

Private Sub Class_Initialize()
    InitPower2
End Sub

Private Function sinh(angle As Single) As Single
    sinh = (Exp(angle) - Exp(-angle)) / 2#
End Function
