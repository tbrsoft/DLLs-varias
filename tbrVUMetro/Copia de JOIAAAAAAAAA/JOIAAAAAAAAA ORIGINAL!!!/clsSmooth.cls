VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSmooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' translated from the XMMS volume normalization plugin
' http://volnorm.sourceforge.net/
'
' returns the average of an array of values
' (the XMMS version searches for the maximum
'  and once it got it, it will return it forever,
'  that's not what I want)

Private sngData()   As Single
Private sngMax      As Single
Private lngSize     As Long
Private lngUsed     As Long
Private lngCurrent  As Long

Public Sub SmoothNew( _
    ByVal size As Long _
)

    ReDim sngData(size - 1) As Single

    lngSize = size
    lngCurrent = 0
    lngUsed = 0
    sngMax = 0
End Sub

Public Sub SmoothAdd( _
    ByVal Sample As Single _
)

    ' /* Put the sample in the buffer */
    sngData(lngCurrent) = Sample

    ' /* Adjust the sample stats */
    lngCurrent = lngCurrent + 1

    If lngCurrent > lngUsed Then
        lngUsed = lngUsed + 1
    End If

    If lngCurrent >= lngSize Then
        lngCurrent = lngCurrent Mod lngSize
    End If
End Sub

Public Function SmoothGetMax( _
) As Single

    Dim smoothed    As Single
    Dim i           As Long

    ' /* Calculate the smoothed value */
    For i = 0 To lngUsed - 1
        smoothed = smoothed + sngData(i)
    Next

    smoothed = smoothed / lngUsed

    ' /* If we haven't filled the smoothing buffer, dont save the max value. */
    If lngUsed < lngSize Then
        SmoothGetMax = smoothed
        Exit Function
    End If

    If sngMax < smoothed Then
        sngMax = smoothed
    End If

    SmoothGetMax = smoothed
End Function
