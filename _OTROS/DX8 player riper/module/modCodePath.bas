Attribute VB_Name = "modCodePath"
Option Explicit

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long _
) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long _
) As Long

Private objTarget       As Object
Private objParameter    As Object
Private strTargetMethod As String
Private hCPTimer        As Long

Public Sub LooseCodePath1P( _
    ByVal targetobj As Object, _
    targetmethod As String, _
    param As Object _
)

    Set objTarget = targetobj
    Set objParameter = param
    strTargetMethod = targetmethod

    hCPTimer = SetTimer(0, 0, 1, _
                        AddressOf CodePathProc)
End Sub

Private Sub CodePathProc( _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal hTimer As Long, _
    ByVal dwTime As Long _
)

    KillTimer 0, hCPTimer

    CallByName objTarget, _
               strTargetMethod, _
               VbMethod, _
               objParameter
End Sub
