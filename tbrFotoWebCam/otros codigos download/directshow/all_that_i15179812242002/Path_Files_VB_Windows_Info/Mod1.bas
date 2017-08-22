Attribute VB_Name = "Mod1"
Option Explicit

Private Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Private Const zFormOrPictBoxStr = "Must pass in the name of either a Form or a PictureBox."

Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal HWND As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal HWND As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Function Trim2(ByVal cString As String) As String
  Dim t As Long, Z As Long
  For t = 1 To Len(cString)
    If Mid$(cString, t, 1) <> " " And Mid$(cString, t, 1) <> Chr$(0) Then Exit For
  Next t
  For Z = Len(cString) To 1 Step -1
    If Mid$(cString, Z, 1) <> " " And Mid$(cString, Z, 1) <> Chr$(0) Then Exit For
  Next Z
  If Z < t Then
    Trim2 = ""
  ElseIf Z = t Then
    Trim2 = Mid$(cString, t, 1)
  Else
    Trim2 = Mid$(cString, t, (Z - t) + 1)
  End If
End Function

Private Function isTransparent(zForm As Form) As TransType
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.HWND, cKey, vTrans, ALPHA
  If Err Then
    isTransparent = -1
  Else
    isTransparent = ALPHA
  End If
End Function

Private Function GetTrans(zForm As Form) As Long
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.HWND, cKey, vTrans, ALPHA
  If ALPHA = LWA_ALPHA Then
    GetTrans = vTrans
  ElseIf ALPHA = LWA_COLORKEY Then
    GetTrans = cKey
  Else
    GetTrans = -1
  End If
  If Err Then
    GetTrans = -1
  End If
End Function

Public Function FadeIn(zForm As Form, Optional ByVal Final As Byte = 255, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 0
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 0
    vTrans = 0
  End If
  If vTrans > Final Then
    FadeIn = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp < Final
    DoEvents
    VarTmp = VarTmp + vStep
    If VarTmp > Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Err Then
    FadeIn = False
  Else
    FadeIn = True
  End If
End Function

Public Function FadeOut(zForm As Form, Optional ByVal Final As Byte = 0, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 255
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 255
    vTrans = 255
  End If
  If vTrans < Final Then
    FadeOut = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp > Final
    DoEvents
    VarTmp = VarTmp - vStep
    If VarTmp < Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Final = 0 Then zForm.Hide
  If Err Then
    FadeOut = False
  Else
    FadeOut = True
  End If
End Function

Public Function SetTrans(zForm As Form, Optional ByVal vTrans As Byte = 127) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.HWND, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.HWND, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.HWND, 0, vTrans, LWA_ALPHA
  If Err Then
    SetTrans = False
  Else
    SetTrans = True
  End If
End Function

Public Function DrawLine(zFormOrPictBox As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Dim TMP2(0 To 1) As POINTAPI
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    TMP2(0).X = X1
    TMP2(0).Y = Y1
    TMP2(1).X = X2
    TMP2(1).Y = Y2
    DrawLine = Polygon(zFormOrPictBox.hdc, TMP2(0), 2)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "Super"
    DrawLine = 0
  End If
End Function
