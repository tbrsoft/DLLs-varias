Attribute VB_Name = "modDSP"
Option Explicit

Public clsDSP               As tbrDX8.SignalProcessor

Private Const FLABUFLEN     As Long = 350

Private intEcho()           As Integer
Private lngEchoPos          As Long
Private lngEchoLength       As Single

Public Sub DSPEcho( _
    intSamples() As Integer, _
    ByVal datalength As Long _
)

    Dim i   As Long

    For i = 0 To datalength
        intSamples(i) = norm(CLng(intSamples(i)) + intEcho(lngEchoPos))
        intEcho(lngEchoPos) = intSamples(i) * lngEchoLength

        lngEchoPos = lngEchoPos + 1
        If lngEchoPos > UBound(intEcho) Then
            lngEchoPos = 0
        End If
    Next
End Sub

Public Sub DSPEchoSettings( _
    ByVal samplerate As Long, _
    ByVal echo_length_ms As Long, _
    ByVal echo_length As Single _
)

    Dim lngEchoPoints   As Long

    lngEchoPoints = samplerate / 1000 * echo_length_ms
    ReDim intEcho(lngEchoPoints - 1) As Integer

    lngEchoLength = echo_length

    lngEchoPos = 0
End Sub

' http://www.un4seen.com/
Public Sub DSPFlange( _
    intSamples() As Integer, _
    ByVal datalen As Long _
)

    Static flapos               As Single
    Static flabuf(FLABUFLEN, 2) As Single
    Static flas                 As Single
    Static flasinc              As Single

    Dim i                       As Long
    Dim p1                      As Long
    Dim p2                      As Long
    Dim F                       As Single
    Dim s                       As Single

    If flas = 0 Then flas = FLABUFLEN / 2
    If flasinc = 0 Then flasinc = 0.002

    For i = 0 To datalen Step 2
        p1 = (flapos + Int(flas)) Mod FLABUFLEN
        p2 = (p1 + 1) Mod FLABUFLEN
        F = flas - Fix(flas / 1) * 1

        s = intSamples(i) + ((flabuf(p1, 0) * (1 - F)) + (flabuf(p2, 0) * F))
        flabuf(flapos, 0) = intSamples(i)
        intSamples(i) = norm(s)

        s = intSamples(i + 1) + ((flabuf(p1, 1) * (1 - F)) + (flabuf(p2, 1) * F))
        flabuf(flapos, 1) = intSamples(i + 1)
        intSamples(i + 1) = norm(s)

        flapos = flapos + 1
        If (flapos = FLABUFLEN) Then flapos = 0
        flas = flas + flasinc
        If ((flas < 0#) Or (flas > FLABUFLEN)) Then flasinc = -flasinc
    Next
End Sub

Private Function norm( _
    ByVal dbl As Double _
) As Integer

    If dbl > 32767 Then
        norm = 32767
    ElseIf dbl < -32768 Then
        norm = -32768
    Else
        norm = CInt(dbl)
    End If
End Function

Public Function dBFS(ByVal amplitude As Long) As Double
    If amplitude = 0 Then
        dBFS = -96
    Else
        dBFS = 20 * ((Log(Abs(amplitude) / 32768)) / Log(10))
    End If
End Function
