VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IIRFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Pi        As Single = 3.14159265358979
Private Const LN2       As Single = 0.693147180559945

Public Enum IIR_TYPE
    IIR_LOW_PASS
    IIR_HIGH_PASS
    IIR_BAND_PASS
    IIR_PEAK_EQ
End Enum

Private fa0             As Single
Private fa1             As Single
Private fa2             As Single
Private fa3             As Single
Private fa4             As Single

Private fX1             As Single
Private fX2             As Single
Private fY1             As Single
Private fY2             As Single

Private lngSR           As Long
Private udeType         As IIR_TYPE
Private sngDB           As Single
Private sngFreq         As Single
Private sngBW           As Single

Private blnEnabled      As Boolean

Public Property Get Enabled( _
) As Boolean

    Enabled = blnEnabled
End Property

Public Property Let Enabled( _
    ByVal blnVal As Boolean _
)

    blnEnabled = blnVal
End Property

Public Property Get Bandwidth( _
) As Single

    Bandwidth = sngBW
End Property

Public Property Let Bandwidth( _
    ByVal sngVal As Single _
)

    sngBW = sngVal
    UpdateSettings
End Property

Public Property Get IIRType( _
) As IIR_TYPE

    IIRType = udeType
End Property

Public Property Get GainDB( _
) As Single

    GainDB = sngDB
End Property

Public Property Let GainDB( _
    ByVal sngVal As Single _
)

    sngDB = sngVal
    UpdateSettings
End Property

Public Property Get Frequency( _
) As Single

    Frequency = sngFreq
End Property

Public Property Let Frequency( _
    ByVal sngVal As Single _
)

    sngFreq = sngVal
    UpdateSettings
End Property

Public Property Get samplerate( _
) As Long

    samplerate = lngSR
End Property

Public Property Let samplerate( _
    ByVal lngVal As Long _
)

    lngSR = lngVal
    UpdateSettings
End Property

Private Sub UpdateSettings()
    If lngSR > 0 Then
        CreateBiquadIIR udeType, sngDB, sngFreq, lngSR, sngBW
    End If
End Sub

Public Sub ProcessSamples_8bit( _
    btSamples() As Byte _
)

    Dim result  As Single
    Dim sample  As Single
    Dim i       As Long

    If Not blnEnabled Then Exit Sub

    For i = 0 To UBound(btSamples)
        sample = (btSamples(i) - 128) / 127
        result = fa0 * sample + fa1 * fX1 + fa2 * fX2 - fa3 * fY1 - fa4 * fY2

        fX2 = fX1
        fX1 = sample

        fY2 = fY1
        fY1 = result

        If result > 1# Then
            btSamples(i) = 255
        ElseIf result < -1# Then
            btSamples(i) = 0
        Else
            btSamples(i) = CByte((result * 127) + 128)
        End If
    Next
End Sub

Public Sub ProcessSamples_16bit( _
    intSamples() As Integer _
)

    Dim result  As Single
    Dim sample  As Single
    Dim i       As Long

    If Not blnEnabled Then Exit Sub

    For i = 0 To UBound(intSamples)
        sample = intSamples(i) / 32767
        result = fa0 * sample + fa1 * fX1 + fa2 * fX2 - fa3 * fY1 - fa4 * fY2

        fX2 = fX1
        fX1 = sample

        fY2 = fY1
        fY1 = result

        If result > 1# Then
            intSamples(i) = 32767
        ElseIf result < -1# Then
            intSamples(i) = -32768
        Else
            intSamples(i) = CInt(result * 32767)
        End If
    Next
End Sub

' various Biquad IIR Filters
' http://www.dspguru.com/sw/lib/biquad.c

' Low/High/Bandpass: dBGain not used
'                    freq is cutoff frequency
'                    srate is samplerate
'                    bandwidth in octaves
'
' Peak EQ: dBGain is boost in dB
'          freq is center frequency of the band
'          srate is samplerate
'          bandwidth in octaves
'
Public Sub CreateBiquadIIR( _
    ByVal IIRType As IIR_TYPE, _
    ByVal dBGain As Single, _
    ByVal freq As Single, _
    ByVal srate As Single, _
    ByVal bw As Single _
)

    Dim A       As Single, omega    As Single
    Dim sn      As Single, cs       As Single
    Dim Alpha   As Single, Beta     As Single

    Dim a0      As Single, a1       As Single, a2   As Single
    Dim b0      As Single, b1       As Single, b2   As Single

    lngSR = srate
    sngBW = bw
    sngDB = dBGain
    sngFreq = freq
    udeType = IIRType

    A = 10 ^ (dBGain / 40)
    omega = 2 * Pi * freq / srate
    sn = Sin(omega)
    cs = Cos(omega)
    Alpha = sn * sinh(LN2 / 2 * bw * omega / sn)
    Beta = Sqr(A + A)

    Select Case IIRType
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
        Case IIR_PEAK_EQ
            b0 = 1 + (Alpha * A)
            b1 = -2 * cs
            b2 = 1 - (Alpha * A)
            a0 = 1 + (Alpha / A)
            a1 = -2 * cs
            a2 = 1 - (Alpha / A)
    End Select

    fa0 = b0 / a0
    fa1 = b1 / a0
    fa2 = b2 / a0
    fa3 = a1 / a0
    fa4 = a2 / a0

    fX1 = 0: fX2 = 0
    fY1 = 0: fY2 = 0
End Sub

Private Function sinh( _
    angle As Single _
) As Single

    sinh = (Exp(angle) - Exp(-angle)) / 2#
End Function

Private Sub Class_Initialize()
    blnEnabled = True
End Sub
