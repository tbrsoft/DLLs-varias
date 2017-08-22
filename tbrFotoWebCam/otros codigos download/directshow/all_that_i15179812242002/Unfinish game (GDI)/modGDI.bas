Attribute VB_Name = "modGDI"
Option Explicit

Public Enum Align
  TA_LEFT = 0
  TA_RIGHT = 2
  TA_CENTER = 6
  TA_TOP = 0
  TA_BOTTOM = 8
  TA_BASELINE = 24
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RGBColor
  cRed As Byte
  cGreen As Byte
  cBlue As Byte
End Type

Public Const zFormOrPictBoxStr = "Must pass in the name of either a Form or a PictureBox."

Private Declare Function GetCurrentPositionEx Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo2 Lib "gdi32.dll" Alias "LineTo" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetPixel2 Lib "gdi32.dll" Alias "GetPixel" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel2 Lib "gdi32.dll" Alias "SetPixel" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FloodFill2 Lib "gdi32.dll" Alias "FloodFill" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function HighByte Lib "tlbinf32.dll" Alias "hibyte" (ByVal Word As Integer) As Byte
Private Declare Function LowByte Lib "tlbinf32.dll" Alias "lobyte" (ByVal Word As Integer) As Byte
Private Declare Function HighWord Lib "tlbinf32.dll" Alias "hiword" (ByVal DWord As Long) As Integer
Private Declare Function LowWord Lib "tlbinf32.dll" Alias "loword" (ByVal DWord As Long) As Integer
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
  
Public Function GetCurrentX(zFormOrPictBox As Object) As Variant
  Dim TMP As POINTAPI
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If GetCurrentPositionEx(zFormOrPictBox.hdc, TMP) = 0 Then
      GetCurrentX = "ERROR"
    Else
      GetCurrentX = TMP.X
    End If
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    GetCurrentX = "ERROR"
  End If
End Function

Public Function GetCurrentY(zFormOrPictBox As Object) As Variant
  Dim TMP As POINTAPI
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If GetCurrentPositionEx(zFormOrPictBox.hdc, TMP) = 0 Then
      GetCurrentY = "ERROR"
    Else
      GetCurrentY = TMP.Y
    End If
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    GetCurrentY = "ERROR"
  End If
End Function

Public Function GetCurrentPosition(zFormOrPictBox As Object, ByRef X As Long, ByRef Y As Long) As Long
  Dim TMP As POINTAPI, var1 As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    var1 = GetCurrentPositionEx(zFormOrPictBox.hdc, TMP)
    If var1 = 0 Then
      GetCurrentPosition = 0
    Else
      X = TMP.X
      Y = TMP.Y
      GetCurrentPosition = var1
    End If
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    GetCurrentPosition = 0
  End If
End Function

Public Function MoveTo(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long) As Long
  Dim TMP As POINTAPI
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    MoveTo = MoveToEx(zFormOrPictBox.hdc, X, Y, TMP)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    MoveTo = 0
  End If
End Function

Public Function LineTo(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    LineTo = LineTo2(zFormOrPictBox.hdc, X, Y)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    LineTo = 0
  End If
End Function

Public Function GetPixel(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    GetPixel = GetPixel2(zFormOrPictBox.hdc, X, Y)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    GetPixel = -1
  End If
End Function

Public Function SetPixel(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    SetPixel = SetPixel2(zFormOrPictBox.hdc, X, Y, crColor)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    SetPixel = 0
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
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawLine = 0
  End If
End Function

Public Function DrawTriangle(zFormOrPictBox As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
  Dim TMP2(0 To 2) As POINTAPI
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    TMP2(0).X = X1
    TMP2(0).Y = Y1
    TMP2(1).X = X2
    TMP2(1).Y = Y2
    TMP2(2).X = X3
    TMP2(2).Y = Y3
    DrawTriangle = Polygon(zFormOrPictBox.hdc, TMP2(0), 3)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawTriangle = 0
  End If
End Function

Public Function DrawAngleCircle(zFormOrPictBox As Object, ByVal X As Single, ByVal Y As Single, ByVal dwRadius As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1) As Boolean
  Const PI As Single = 3.14159265
  Dim SM As Integer, FC As Long, DW As Integer
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If StartAngle < 0 Or EndAngle < 0 Or StartAngle > 360 Or EndAngle > 360 Then
      MsgBox "StartAngle and EndAngle must be between 0 and 360 !", vbExclamation, "SuperDLL"
      DrawAngleCircle = False
      Exit Function
    End If
    If ForColor <> -1 Then
      FC = zFormOrPictBox.ForeColor
      zFormOrPictBox.ForeColor = ForColor
    End If
    If dWidth <> -1 Then
      DW = zFormOrPictBox.DrawWidth
      zFormOrPictBox.DrawWidth = dWidth
    End If
    If StartAngle = 0 And EndAngle = 360 Then EndAngle = 0
    SM = zFormOrPictBox.ScaleMode
    zFormOrPictBox.ScaleMode = 3
    zFormOrPictBox.Circle (X, Y), dwRadius, , (StartAngle * PI) / 180, (EndAngle * PI / 180)
    zFormOrPictBox.ScaleMode = SM
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = FC
    If dWidth <> -1 Then zFormOrPictBox.DrawWidth = DW
    DrawAngleCircle = True
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawAngleCircle = False
  End If
End Function

Public Function DrawAngleEllipse(zFormOrPictBox As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1) As Boolean
  Const PI As Single = 3.14159265
  Dim dwRadius As Single, Aspect As Single, SM As Integer, FC As Long, DW As Integer
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If StartAngle < 0 Or EndAngle < 0 Or StartAngle > 360 Or EndAngle > 360 Then
      MsgBox "StartAngle and EndAngle must be between 0 and 360 !", vbExclamation, "SuperDLL"
      DrawAngleEllipse = False
      Exit Function
    End If
    If ForColor <> -1 Then
      FC = zFormOrPictBox.ForeColor
      zFormOrPictBox.ForeColor = ForColor
    End If
    If dWidth <> -1 Then
      DW = zFormOrPictBox.DrawWidth
      zFormOrPictBox.DrawWidth = dWidth
    End If
    If StartAngle = 0 And EndAngle = 360 Then EndAngle = 0
    Aspect = (Y2 - Y1) / (X2 - X1)
    dwRadius = (X2 - X1) / 2
    If (X2 - X1) = (Y2 - Y1) Then
      Aspect = 1
    ElseIf (Y2 - Y1) > (X2 - X1) Then
      dwRadius = (Y2 - Y1) / 2
    End If
    SM = zFormOrPictBox.ScaleMode
    zFormOrPictBox.ScaleMode = 3
    zFormOrPictBox.Circle ((X2 + X1) / 2, (Y2 + Y1) / 2), dwRadius, , (StartAngle * PI) / 180, (EndAngle * PI / 180), Aspect
    zFormOrPictBox.ScaleMode = SM
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = FC
    If dWidth <> -1 Then zFormOrPictBox.DrawWidth = DW
    DrawAngleEllipse = True
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawAngleEllipse = False
  End If
End Function

Public Function DrawCircle(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    DrawCircle = Ellipse(zFormOrPictBox.hdc, X - dwRadius, Y - dwRadius, X + dwRadius + 1, Y + dwRadius + 1)
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawCircle = 0
  End If
End Function

Public Function DrawEllipse(zFormOrPictBox As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    DrawEllipse = Ellipse(zFormOrPictBox.hdc, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0))
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawEllipse = 0
  End If
End Function

Public Function DrawRectangle(zFormOrPictBox As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    DrawRectangle = Rectangle(zFormOrPictBox.hdc, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0))
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawRectangle = 0
  End If
End Function

Public Function DrawRoundRect(zFormOrPictBox As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal pcRoundX As Integer, Optional ByVal pcRoundY As Integer = -1) As Long
  Dim X3 As Long, Y3 As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If pcRoundX > 100 Or pcRoundY > 100 Or pcRoundX < 0 Or pcRoundY < -1 Then
      MsgBox "pcRoundX and pcRoundY must be between 0 and 100 !", vbExclamation, "SuperDLL"
      DrawRoundRect = 0
    Else
      X3 = (pcRoundX * (X2 - X1)) / 100
      If pcRoundY = -1 Then
        Y3 = X3
      Else
        Y3 = (pcRoundY * (Y2 - Y1)) / 100
      End If
      DrawRoundRect = RoundRect(zFormOrPictBox.hdc, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0), X3, Y3)
    End If
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawRoundRect = 0
  End If
End Function

Public Function SetColor(zFormOrPictBox As Object, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1, Optional ByVal tAlign As Align = -1) As Boolean
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = ForColor
    If dWidth <> -1 Then zFormOrPictBox.DrawWidth = dWidth
    If FilColor <> -1 Then zFormOrPictBox.FillColor = FilColor
    If FilStyle <> -1 Then zFormOrPictBox.FillStyle = FilStyle
    If tAlign <> -1 Then SetTextAlign zFormOrPictBox.hdc, tAlign
    SetColor = True
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    SetColor = False
  End If
End Function

Public Function FloodFill(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long, ByVal BorderColor As Long, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1) As Long
  Dim FC As Long, FS As FillStyleConstants
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If FilColor <> -1 Then
      FC = zFormOrPictBox.FillColor
      zFormOrPictBox.FillColor = FilColor
    End If
    If FilStyle <> -1 Then
      FS = zFormOrPictBox.FillStyle
      zFormOrPictBox.FillStyle = FilStyle
    End If
    FloodFill = FloodFill2(zFormOrPictBox.hdc, X, Y, BorderColor)
    If FilColor <> -1 Then zFormOrPictBox.FillColor = FC
    If FilStyle <> -1 Then zFormOrPictBox.FillStyle = FS
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    FloodFill = 0
  End If
End Function

Public Function GetRGB(ByVal cColor As Long) As RGBColor
  GetRGB.cRed = LowByte(LowWord(cColor))
  GetRGB.cGreen = HighByte(LowWord(cColor))
  GetRGB.cBlue = LowByte(HighWord(cColor))
End Function

Public Function GetRed(ByVal cColor As Long) As Byte
  GetRed = LowByte(LowWord(cColor))
End Function

Public Function GetGreen(ByVal cColor As Long) As Byte
  GetGreen = HighByte(LowWord(cColor))
End Function

Public Function GetBlue(ByVal cColor As Long) As Byte
  GetBlue = LowByte(HighWord(cColor))
End Function

Public Function DrawText(zFormOrPictBox As Object, ByVal zString As String, ByVal X As Long, ByVal Y As Long, Optional ByVal tAlign As Align = -1, Optional ByVal ForColor As Long = -1) As Long
  Dim TA As Long, FC As Long
  If (TypeOf zFormOrPictBox Is Form) Or (TypeOf zFormOrPictBox Is PictureBox) Then
    If tAlign <> -1 Then
      TA = GetTextAlign(zFormOrPictBox.hdc)
      SetTextAlign zFormOrPictBox.hdc, tAlign
    End If
    If ForColor <> -1 Then
      FC = zFormOrPictBox.ForeColor
      zFormOrPictBox.ForeColor = ForColor
    End If
    DrawText = TextOut(zFormOrPictBox.hdc, X, Y, zString, Len(zString))
    If tAlign <> -1 Then SetTextAlign zFormOrPictBox.hdc, TA
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = FC
  Else
    MsgBox zFormOrPictBoxStr, vbExclamation, "SuperDLL"
    DrawText = 0
  End If
End Function
