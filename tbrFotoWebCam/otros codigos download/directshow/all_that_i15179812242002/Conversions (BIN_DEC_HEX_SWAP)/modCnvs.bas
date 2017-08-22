Attribute VB_Name = "modCnvs"
Option Explicit

Private Declare Function HighByte Lib "tlbinf32.dll" Alias "hibyte" (ByVal Word As Integer) As Byte
Private Declare Function LowByte Lib "tlbinf32.dll" Alias "lobyte" (ByVal Word As Integer) As Byte
Private Declare Function HighWord Lib "tlbinf32.dll" Alias "hiword" (ByVal DWord As Long) As Integer
Private Declare Function LowWord Lib "tlbinf32.dll" Alias "loword" (ByVal DWord As Long) As Integer
Private Declare Function MakeDWord2 Lib "tlbinf32.dll" Alias "makelong" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Private Declare Function MakeWord2 Lib "tlbinf32.dll" Alias "makeword" (ByVal bLow As Byte, ByVal bHigh As Byte) As Integer

Public Function DecToHex(ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True) As Variant
  Dim var2 As String
  var2 = Hex(var1)
  If AddToNextType Then
    Select Case Len(var2)
      Case 1, 3, 7
        var2 = "0" & var2
      Case 5
        var2 = "000" & var2
      Case 6
        var2 = "00" & var2
    End Select
  End If
  DecToHex = var2
End Function

Public Function HexToDec(ByVal var1 As String) As Long
  On Local Error GoTo ErrCnv
  HexToDec = "&h" & Trim2(var1)
  Exit Function
ErrCnv:
  HexToDec = -1
End Function

Private Function HexBin(ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim t As Long, qaz As String
  qaz = ""
  If Len(var1) = 0 Then
    HexBin = -1
    Exit Function
  Else
    For t = 1 To Len(var1)
      Select Case UCase$(Mid$(var1, t, 1))
        Case "0"
          qaz = qaz & "0000"
        Case "1"
          qaz = qaz & "0001"
        Case "2"
          qaz = qaz & "0010"
        Case "3"
          qaz = qaz & "0011"
        Case "4"
          qaz = qaz & "0100"
        Case "5"
          qaz = qaz & "0101"
        Case "6"
          qaz = qaz & "0110"
        Case "7"
          qaz = qaz & "0111"
        Case "8"
          qaz = qaz & "1000"
        Case "9"
          qaz = qaz & "1001"
        Case "A"
          qaz = qaz & "1010"
        Case "B"
          qaz = qaz & "1011"
        Case "C"
          qaz = qaz & "1100"
        Case "D"
          qaz = qaz & "1101"
        Case "E"
          qaz = qaz & "1110"
        Case "F"
          qaz = qaz & "1111"
        Case Else
          HexBin = -1
          Exit Function
      End Select
    Next t
  End If
  If RemoveLeadingZeros Then
    For t = 1 To Len(qaz)
      If Mid$(qaz, t, 1) <> "0" Then Exit For
    Next t
    qaz = Mid$(qaz, t)
  ElseIf AddToNextType Then
    Select Case Len(qaz)
      Case 4, 12, 28
        qaz = "0000" & qaz
      Case 20
        qaz = "000000000000" & qaz
      Case 24
        qaz = "00000000" & qaz
    End Select
  End If
  HexBin = qaz
End Function

Public Function HexToBin(ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim qwe As Variant
  qwe = HexBin(Trim2(var1), AddToNextType, RemoveLeadingZeros)
  If Len(qwe) = 0 Then qwe = "0"
  HexToBin = qwe
End Function

Private Function BinHex(ByVal var0 As String) As Variant
  Select Case UCase$(var0)
    Case "0000", "000", "00", "0"
      BinHex = "0"
    Case "0001", "001", "01", "1"
      BinHex = "1"
    Case "0010", "010", "10"
      BinHex = "2"
    Case "0011", "011", "11"
      BinHex = "3"
    Case "0100", "100"
      BinHex = "4"
    Case "0101", "101"
      BinHex = "5"
    Case "0110", "110"
      BinHex = "6"
    Case "0111", "111"
      BinHex = "7"
    Case "1000"
      BinHex = "8"
    Case "1001"
      BinHex = "9"
    Case "1010"
      BinHex = "A"
    Case "1011"
      BinHex = "B"
    Case "1100"
      BinHex = "C"
    Case "1101"
      BinHex = "D"
    Case "1110"
      BinHex = "E"
    Case "1111"
      BinHex = "F"
    Case Else
      BinHex = -1
  End Select
End Function

Public Function BinToHex(ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim t As Long, q As Variant, qaz As String, qwe As String
  qwe = Trim2(var1)
  qaz = ""
  If Len(qwe) = 0 Then
    BinToHex = -1
    Exit Function
  Else
    Do
      q = BinHex(Right$(qwe, 4))
      If q = -1 Then
        BinToHex = -1
        Exit Function
      End If
      qaz = q & qaz
      If Len(qwe) <= 4 Then
        qwe = ""
      Else
        qwe = Left$(qwe, Len(qwe) - 4)
      End If
    Loop Until Len(qwe) < 1
  End If
  If RemoveLeadingZeros Then
    For t = 1 To Len(qaz)
      If Mid$(qaz, t, 1) <> "0" Then Exit For
    Next t
    qaz = Mid$(qaz, t)
  ElseIf AddToNextType Then
    Select Case Len(qaz)
      Case 1, 3, 7
        qaz = "0" & qaz
      Case 5
        qaz = "000" & qaz
      Case 6
        qaz = "00" & qaz
    End Select
  End If
  If Len(qaz) = 0 Then qaz = "0"
  BinToHex = qaz
End Function

Public Function BinToDec(ByVal var0 As String) As Long
  Dim qwe As String
  qwe = BinToHex(Trim2(var0), False, False)
  On Local Error GoTo ErrCnv
  BinToDec = "&h" & qwe
  Exit Function
ErrCnv:
  BinToDec = -1
End Function

Public Function DecToBin(ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim qwe As String, qaz As String
  qwe = DecToHex(var1, False)
  qaz = HexBin(qwe, AddToNextType, RemoveLeadingZeros)
  If Len(qaz) = 0 Then qaz = "0"
  DecToBin = qaz
End Function

Public Function HiByte(ByVal Word As Integer) As Byte
  HiByte = HighByte(Word)
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
  LoByte = LowByte(Word)
End Function

Public Function HiWord(ByVal DWord As Long) As Integer
  HiWord = HighWord(DWord)
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
  LoWord = LowWord(DWord)
End Function

Public Function HiByteHiWord(ByVal DWord As Long) As Byte
  HiByteHiWord = HighByte(HighWord(DWord))
End Function

Public Function LoByteHiWord(ByVal DWord As Long) As Byte
  LoByteHiWord = LowByte(HighWord(DWord))
End Function

Public Function HiByteLoWord(ByVal DWord As Long) As Byte
  HiByteLoWord = HighByte(LowWord(DWord))
End Function

Public Function LoByteLoWord(ByVal DWord As Long) As Byte
  LoByteLoWord = LowByte(LowWord(DWord))
End Function

Public Function MakeWord(ByVal HByte As Byte, ByVal LByte As Byte) As Integer
  MakeWord = MakeWord2(LByte, HByte)
End Function

Public Function MakeDWordB(ByVal HByteHWord As Byte, ByVal LByteHWord As Byte, ByVal HByteLWord As Byte, ByVal LByteLWord As Byte) As Long
  MakeDWordB = MakeDWord2(MakeWord2(LByteLWord, HByteLWord), MakeWord2(LByteHWord, HByteHWord))
End Function

Public Function MakeDWordW(ByVal HWord As Integer, LWord As Integer) As Long
  MakeDWordW = MakeDWord2(LWord, HWord)
End Function

Public Sub Swap(var1 As Variant, var2 As Variant)
  Dim var3 As Variant
  var3 = var1: var1 = var2: var2 = var3
End Sub

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
