VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Equalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' parametric equalizer
' (peaking EQ biquad IIR Filter)
'
' from rb-j`s Audio EQ Cookbook:
' http://www.musicdsp.org/files/Audio-EQ-Cookbook.txt

Private Type IIRCoeff
    a0      As Single
    a1      As Single
    a2      As Single
    a3      As Single
    a4      As Single
    x1      As Single
    x2      As Single
    y1      As Single
    y2      As Single
    d0      As Single
    d1      As Single
    gain    As Single
    freq    As Single
    bw      As Single
End Type

Private udtFilters()    As IIRCoeff

Private lngSamplerate   As Long

Private Const LN2       As Single = 0.693147180559945
Private Const Pi        As Single = 3.14159265358979

Public Sub SetBandCount( _
    ByVal bands As Byte _
)

    ReDim udtFilters(bands - 1)
End Sub

Public Property Get samplerate() As Long
    samplerate = lngSamplerate
End Property

Public Property Let samplerate( _
    ByVal lng As Long _
)

    On Error Resume Next

    lngSamplerate = lng
    Update
End Property

Public Sub ProcessSamples( _
    intSamples() As Integer, _
    ByVal datalen As Long _
)

    Dim result  As Single
    Dim sample  As Single
    Dim i       As Long
    Dim j       As Long
    Dim ubnd    As Long

    ubnd = UBound(udtFilters)

    For i = 0& To datalen
        sample = intSamples(i) / 32768

        For j = 0& To ubnd
            With udtFilters(j)
                ' Direct Form 2
'                result = .a0 * sample + .d0
'                .d0 = .a1 * sample - .a3 * result + .d1
'                .d1 = .a2 * sample - .a4 * result

               ' Direct Form 1
               result = .a0 * sample + .a1 * .x1 + .a2 * .x2 - .a3 * .y1 - .a4 * .y2

               .x2 = .x1
               .x1 = sample

               .y2 = .y1
               .y1 = result
            End With

            If result > 1# Then
                result = 1#
            ElseIf result < -1# Then
                result = -1#
            End If

            sample = result
        Next

        intSamples(i) = CInt(result * 32767)
    Next
End Sub

Private Sub Update()
    Dim i   As Integer

    For i = 0 To UBound(udtFilters)
        With udtFilters(i)
            SetBandValues i, .gain, .freq, .bw
        End With
    Next
End Sub

Public Sub SetBandValues( _
    ByVal band As Byte, _
    ByVal dBGain As Single, _
    ByVal Frequency As Single, _
    ByVal bandwidth As Single _
)

    udtFilters(band) = CreatePeakEQ(dBGain, Frequency, lngSamplerate, bandwidth)
End Sub

Private Function CreatePeakEQ( _
    ByVal dBGain As Single, _
    ByVal freq As Single, _
    ByVal srate As Single, _
    ByVal bandwidth As Single _
) As IIRCoeff

    Dim A       As Single, omega    As Single
    Dim sn      As Single, cs       As Single
    Dim Alpha   As Single, Beta     As Single

    Dim a0      As Single, a1       As Single, a2   As Single
    Dim b0      As Single, b1       As Single, b2   As Single

    Dim filter  As IIRCoeff

    A = 10 ^ (dBGain / 40)
    omega = 2 * Pi * freq / srate
    sn = Sin(omega)
    cs = Cos(omega)
    Alpha = sn * sinh(LN2 / 2 * bandwidth * omega / sn)
    Beta = Sqr(A + A)

    b0 = 1 + (Alpha * A)
    b1 = -2 * cs
    b2 = 1 - (Alpha * A)
    a0 = 1 + (Alpha / A)
    a1 = -2 * cs
    a2 = 1 - (Alpha / A)

    filter.a0 = b0 / a0
    filter.a1 = b1 / a0
    filter.a2 = b2 / a0
    filter.a3 = a1 / a0
    filter.a4 = a2 / a0

    filter.x1 = 0: filter.x2 = 0
    filter.y1 = 0: filter.y2 = 0

    filter.gain = dBGain
    filter.freq = freq
    filter.bw = bandwidth

    CreatePeakEQ = filter
End Function

Private Function sinh( _
    angle As Single _
) As Single

    sinh = (Exp(angle) - Exp(-angle)) / 2#
End Function
