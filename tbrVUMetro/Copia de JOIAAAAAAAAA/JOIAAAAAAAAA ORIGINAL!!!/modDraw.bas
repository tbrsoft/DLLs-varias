Attribute VB_Name = "modDraw"
Option Explicit

' some constants shared by Form1 and clsDraw
' (actually only FFT_SAMPLES...)

Public Const FFT_MAXAMPLITUDE       As Double = 0.2
Public Const FFT_BANDLOWER          As Double = 0.07
Public Const FFT_BANDS              As Long = 22
Public Const FFT_BANDSPACE          As Long = 1
Public Const FFT_BANDWIDTH          As Long = 3
Public Const FFT_STARTINDEX         As Long = 1
Public Const FFT_SAMPLES            As Long = 1024

Public Const DRW_BARXOFF            As Long = 4
Public Const DRW_BARYOFF            As Long = 2
Public Const DRW_BARWIDTH           As Long = 3
Public Const DRW_BARSPACE           As Long = 1

Private Type vbcolor
    red                             As Single
    green                           As Single
    blue                            As Single
End Type

Public Function GetGradColor( _
    ByVal max As Single, _
    ByVal value As Single, _
    ByVal colstart As Long, _
    ByVal colmiddle As Long, _
    ByVal colend As Long _
) As Long

    Dim udtCol1 As vbcolor
    Dim udtCol2 As vbcolor
    Dim udtCol3 As vbcolor
    Dim udtColS As vbcolor
    Dim udtColE As vbcolor
    Dim udtDrw  As vbcolor
    Dim x       As Long

    udtCol1 = translate_color(colstart)
    udtCol2 = translate_color(colmiddle)
    udtCol3 = translate_color(colend)

    If value < max / 2 Then
        udtDrw = udtCol1
        udtColS = udtCol1
        udtColE = udtCol2
    Else
        value = value - (max / 2)
        udtDrw = udtCol2
        udtColS = udtCol2
        udtColE = udtCol3
    End If

    max = max / 2

    With udtDrw
        .red = .red + (((udtColE.red - udtColS.red) / max) * value)
        .green = .green + (((udtColE.green - udtColS.green) / max) * value)
        .blue = .blue + (((udtColE.blue - udtColS.blue) / max) * value)

        GetGradColor = RGB(.red, .green, .blue)
    End With
End Function

Private Function translate_color( _
    ByVal olecol As Long _
) As vbcolor

    With translate_color
        .blue = (olecol \ &H10000) And &HFF
        .green = (olecol \ &H100) And &HFF
        .red = olecol And &HFF
    End With
End Function
