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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' http://www.dspguru.com/sw/lib/biquad.c

Public a0   As Single
Public a1   As Single
Public a2   As Single
Public a3   As Single
Public a4   As Single
Public x1   As Single
Public x2   As Single
Public y1   As Single
Public y2   As Single

Public Sub ProcessSamples( _
    intSamples() As Integer, _
    ByVal datalen As Long _
)

    Dim result  As Single
    Dim sample  As Single
    Dim i       As Long

    For i = 0 To datalen
        sample = intSamples(i) / 32767
        result = a0 * sample + a1 * x1 + a2 * x2 - a3 * y1 - a4 * y2

        x2 = x1
        x1 = sample

        y2 = y1
        y1 = result

        If result > 1# Then
            intSamples(i) = 32767
        ElseIf result < -1# Then
            intSamples(i) = -32768
        Else
            intSamples(i) = CInt(result * 32767)
        End If
    Next
End Sub
