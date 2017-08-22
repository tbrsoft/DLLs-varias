VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Digital Signal Processing Class for 16 bit samples
'
' Echo, Amplifier, Phase Shift, Graphical Equalizer

Private Const Pi                    As Single = 3.14159265358979

Private Enum DSP_VOL_UNIT
    DSP_VOL_DECIBEL = 0
    DSP_VOL_PERCENT
    DSP_VOL_FACTOR
End Enum

Public Enum DSP_FX
    DSP_FX_ECHO = 2 ^ 0
    DSP_FX_PHASE_SHIFT = 2 ^ 1
    DSP_FX_AMPLIFY = 2 ^ 2
    DSP_FX_EQUALIZER = 2 ^ 3
End Enum

Private udeFXFlags                  As DSP_FX

'''' ECHO
Private intEcho16()                 As Integer
Private lngEchoPos16                As Long
Private sngEchoLength16             As Single
Private lngEchoLen                  As Long
Private sngEchoAmp                  As Single

'''' AMPLIFIER
Private lngVolume                   As Long

'''' PHASE SHIFT
Private sngShX1                     As Single
Private sngShX2                     As Single
Private sngShX3                     As Single
Private sngShX4                     As Single

Private sngShY1                     As Single
Private sngShY2                     As Single
Private sngShY3                     As Single
Private sngShY4                     As Single

Private sngShMinWP                  As Single
Private sngShMaxWP                  As Single
Private sngShWp                     As Single
Private sngShRate                   As Single
Private sngShSweepFac               As Single

Private sngSetShDry                 As Single
Private sngSetShWet                 As Single
Private sngSetShFeedback            As Single
Private sngSetShSweepRate           As Single
Private sngSetShSweepRange          As Single
Private sngSetShFreq                As Single

'''' EQUALIZER
Private lngEQBands                  As Long
Private sngBandGain()               As Single
Private clsEQUBands()               As IIRFilter

'''' FORMAT
Private lngSamplerate               As Long
Private intChannels                 As Integer

Public Property Get samplerate( _
) As Long

    samplerate = lngSamplerate
End Property

Public Property Let samplerate( _
    ByVal lngSR As Long _
)

    If lngSR < 1 Or lngSR > 96000 Then
        Err.Raise 32000, "invalid samplerate"
    Else
        lngSamplerate = lngSR
        UpdateFX
    End If
End Property

Public Property Get Channels( _
) As Integer

    Channels = intChannels
End Property

Public Property Let Channels( _
    ByVal intCh As Integer _
)

    If intCh < 1 Or intCh > 2 Then
        Err.Raise 32000, "invalid channels"
    Else
        intChannels = intCh
        UpdateFX
    End If
End Property

Public Property Get EffectsUsed( _
) As DSP_FX

    EffectsUsed = udeFXFlags
End Property

Public Property Let EffectsUsed( _
    ByVal ude As DSP_FX _
)

    udeFXFlags = ude
End Property

Public Property Get EchoLength( _
) As Long

    EchoLength = lngEchoLen
End Property

Public Property Let EchoLength( _
    ByVal lngMS As Long _
)

    lngEchoLen = lngMS
    UpdateEcho
End Property

Public Property Get EchoAmp( _
) As Single

    EchoAmp = sngEchoAmp
End Property

Public Property Let EchoAmp( _
    ByVal sngAmp As Single _
)

    If sngAmp > 0.9 Then sngAmp = 0.9

    sngEchoAmp = sngAmp
    UpdateEcho
End Property

Public Property Get AmplifyDB( _
) As Long

    AmplifyDB = lngVolume
End Property

Public Property Let AmplifyDB( _
    ByVal lngVol As Long _
)

    lngVolume = lngVol
End Property

Public Property Get PhaseShiftWet( _
) As Single

    PhaseShiftWet = sngSetShWet
End Property

Public Property Let PhaseShiftWet( _
    ByVal sng As Single _
)

    sngSetShWet = sng
    UpdatePhaseShift
End Property

Public Property Get PhaseShiftDry( _
) As Single

    PhaseShiftDry = sngSetShDry
End Property

Public Property Let PhaseShiftDry( _
    ByVal sng As Single _
)

    sngSetShDry = sng
    UpdatePhaseShift
End Property

Public Property Get PhaseShiftFeedback( _
) As Single

    PhaseShiftFeedback = sngSetShFeedback
End Property

Public Property Let PhaseShiftFeedback( _
    ByVal sng As Single _
)

    sngSetShFeedback = sng
    UpdatePhaseShift
End Property

Public Property Get PhaseShiftSweepRate( _
) As Single

    PhaseShiftSweepRate = sngSetShSweepRate
End Property

Public Property Let PhaseShiftSweepRate( _
    ByVal sng As Single _
)

    sngSetShSweepRate = sng
    UpdatePhaseShift
End Property

Public Property Get PhaseShiftSweepRange( _
) As Single

    PhaseShiftSweepRange = sngSetShSweepRange
End Property

Public Property Let PhaseShiftSweepRange( _
    ByVal sng As Single _
)

    sngSetShSweepRange = sng
    UpdatePhaseShift
End Property

Public Property Get PhaseShiftFrequency( _
) As Single

    PhaseShiftFrequency = sngSetShFreq
End Property

Public Property Let PhaseShiftFrequency( _
    ByVal sng As Single _
)

    sngSetShFreq = sng
    UpdatePhaseShift
End Property

Public Property Get EqualizerBandFrequency( _
    ByVal Index As Long _
) As Long

    EqualizerBandFrequency = clsEQUBands(Index).Frequency
End Property

Public Property Get EqualizerBandGainDB( _
    ByVal Index As Long _
) As Single

    EqualizerBandGainDB = sngBandGain(Index)
End Property

Public Property Let EqualizerBandGainDB( _
    ByVal Index As Long, _
    ByVal gain_db As Single _
)

    sngBandGain(Index) = gain_db
    UpdateEqualizer
End Property

Public Property Get EqualizerBands( _
) As Long

    EqualizerBands = lngEQBands
End Property

Public Property Let EqualizerBands( _
    ByVal lngBands As Long _
)

    Dim i   As Long

    If lngBands < 1 Or lngBands > 10 Then
        Err.Raise 32000, "Out Of Bounds!"
    Else
        lngEQBands = lngBands
        ReDim clsEQUBands(lngEQBands - 1) As IIRFilter
        ReDim sngBandGain(lngEQBands - 1) As Single

        For i = 0 To lngEQBands - 1
            Set clsEQUBands(i) = New IIRFilter
        Next

        UpdateEqualizer
    End If
End Property

Private Sub UpdateFX()
    If lngSamplerate = 0 Then Exit Sub
    If intChannels = 0 Then Exit Sub

    UpdateEcho
    UpdatePhaseShift
    UpdateEqualizer
End Sub

' http://www.codeproject.com/cs/media/cswavplayfx.asp
Private Sub UpdatePhaseShift()
    Dim range   As Double

    If lngSamplerate = 0 Then Exit Sub

    sngShMinWP = Pi * sngSetShFreq / lngSamplerate
    range = 2 ^ sngSetShSweepRange
    sngShMaxWP = Pi * sngSetShFreq * range / lngSamplerate
    sngShRate = range ^ (2 * sngSetShSweepRate / lngSamplerate)

    sngShSweepFac = sngShRate
    sngShWp = sngShMinWP
End Sub

Private Sub UpdateEcho()
    Dim lngEchoPoints   As Long

    If lngSamplerate = 0 Then Exit Sub

    lngEchoPoints = lngSamplerate / 1000 * lngEchoLen
    ReDim intEcho16(lngEchoPoints - 1) As Integer

    sngEchoLength16 = sngEchoAmp
    lngEchoPos16 = 0
End Sub

Public Sub ProcessSamples( _
    intSamples() As Integer _
)

    Dim i   As Long

    If lngSamplerate = 0 Then Exit Sub
    If intChannels = 0 Then Exit Sub

    If (udeFXFlags And DSP_FX_AMPLIFY) = DSP_FX_AMPLIFY Then
        DSP_Process_ChangeVolume intSamples, lngVolume, DSP_VOL_DECIBEL
    End If

    If (udeFXFlags And DSP_FX_ECHO) = DSP_FX_ECHO Then
        DSP_Process_Echo intSamples
    End If

    If (udeFXFlags And DSP_FX_PHASE_SHIFT) = DSP_FX_PHASE_SHIFT Then
        DSP_Process_PhaseShift intSamples
    End If

    If (udeFXFlags And DSP_FX_EQUALIZER) = DSP_FX_EQUALIZER Then
        For i = 0 To lngEQBands - 1
            clsEQUBands(i).ProcessSamples_16bit intSamples
        Next
    End If
End Sub

Private Sub UpdateEqualizer()
    Dim i           As Long
    Dim lngFreq     As Long
    Dim sngBW       As Single
    Dim sngBFreq    As Single

    If lngSamplerate = 0 Then Exit Sub

    For i = 0 To lngEQBands - 1
        lngFreq = lngSamplerate / 3

        sngBW = Log(lngFreq / 80#) / 4

        With clsEQUBands(i)
            sngBFreq = 80# * (lngFreq / 80#) ^ (i / (lngEQBands - 1))

            .CreateBiquadIIR IIR_PEAK_EQ, _
                             sngBandGain(i), _
                             sngBFreq, _
                             lngSamplerate, _
                             sngBW
        End With
    Next
End Sub

Private Sub DSP_Process_Echo( _
    intSamples() As Integer _
)

    Dim i   As Long

    For i = 0 To UBound(intSamples)
        intSamples(i) = norm16(CLng(intSamples(i)) + intEcho16(lngEchoPos16))
        intEcho16(lngEchoPos16) = intSamples(i) * sngEchoLength16

        lngEchoPos16 = lngEchoPos16 + 1
        If lngEchoPos16 > UBound(intEcho16) Then
            lngEchoPos16 = 0
        End If
    Next
End Sub

Private Sub DSP_Process_ChangeVolume( _
    intSamples() As Integer, _
    ByVal value As Single, _
    ByVal unit As DSP_VOL_UNIT _
)

    Dim sngFactor   As Single
    Dim sngResult   As Single
    Dim i           As Long

    Select Case unit
        Case DSP_VOL_DECIBEL
            sngFactor = 10 ^ (value / 20)
        Case DSP_VOL_PERCENT
            sngFactor = value / 100
        Case DSP_VOL_FACTOR
            sngFactor = value
    End Select

    For i = 0& To UBound(intSamples)
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

' http://www.codeproject.com/cs/media/cswavplayfx.asp
Private Sub DSP_Process_PhaseShift( _
    intSamples() As Integer _
)

    Dim i   As Long
    Dim k   As Single
    Dim X1  As Single
    Dim x   As Single
    Dim y   As Single

    For i = 0& To UBound(intSamples)
        x = intSamples(i) / 32767

        k = (1# - sngShWp) / (1# + sngShWp)

        X1 = x + sngSetShFeedback * sngShY4
        sngShY1 = k * (sngShY1 + X1) - sngShX1
        sngShX1 = X1
        sngShY2 = k * (sngShY2 + sngShY1) - sngShX2
        sngShX2 = sngShY1
        sngShY3 = k * (sngShY3 + sngShY2) - sngShX3
        sngShX3 = sngShY2
        sngShY4 = k * (sngShY4 + sngShY3) - sngShX4
        sngShX4 = sngShY3

        y = sngShY4 * sngSetShWet + x * sngSetShDry

        sngShWp = sngShWp * sngShSweepFac
        If (sngShWp > sngShMaxWP) Then
            sngShSweepFac = 1# / sngShRate
        ElseIf (sngShWp < sngShMinWP) Then
            sngShSweepFac = sngShRate
        End If

        y = y * 32767

        If y > 32767 Then
            intSamples(i) = 32767
        ElseIf y < -32768 Then
            intSamples(i) = -32768
        Else
            intSamples(i) = CInt(y)
        End If
    Next
End Sub

Private Function norm16( _
    ByVal dbl As Single _
) As Integer

    If dbl > 32767 Then
        norm16 = 32767
    ElseIf dbl < -32768 Then
        norm16 = -32768
    Else
        norm16 = CInt(dbl)
    End If
End Function

Private Sub Class_Initialize()
    Dim i   As Long

    sngEchoAmp = 0.4
    lngEchoLen = 100                    ' ms
    UpdateEcho

    lngVolume = 3                       ' dB

    sngSetShDry = 1
    sngSetShWet = 1
    sngSetShFeedback = 0.5
    sngSetShSweepRate = 1
    sngSetShSweepRange = 4
    sngSetShFreq = 100                  ' Hz

    lngEQBands = 7
    ReDim clsEQUBands(lngEQBands - 1) As IIRFilter
    ReDim sngBandGain(lngEQBands - 1) As Single

    For i = 0 To lngEQBands - 1
        Set clsEQUBands(i) = New IIRFilter
    Next
End Sub
