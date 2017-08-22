VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImgDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' Get the differences between two images
'



Public Function GetDiffPixels(pPicture1 As StdPicture, pPicture2 As StdPicture) As Long
'
' Get the difference between two pictures in pixels
'
On Error GoTo GetDiffPixels_Error
'
Dim iWidth1 As Integer
Dim iHeight1 As Integer
Dim iWidth2 As Integer
Dim iHeight2 As Integer
Dim lDiffPix As Long
Dim rXPos As Single
Dim rYPos As Single
Dim lGlobP As Long

Dim lColor1 As Long
Dim lColor2 As Long

Dim rgbVal1 As RGBpoint
Dim rgbVal2 As RGBpoint
Dim rgbSrc As RGBthingy

GetDiffPixels = -1
Set frmDummy.pbxPic1.Picture = pPicture1
Set frmDummy.pbxPic2.Picture = pPicture2
'
iWidth1 = frmDummy.pbxPic1.ScaleWidth
iHeight1 = frmDummy.pbxPic1.ScaleHeight
iWidth2 = frmDummy.pbxPic1.ScaleWidth
iHeight2 = frmDummy.pbxPic1.ScaleHeight
'
If iWidth2 > iWidth1 Or iHeight2 > iHeight1 Then
 frmDummy.pbxPic2.ScaleWidth = iWidth1
 frmDummy.pbxPic2.ScaleHeight = iHeight1
End If
'
If iWidth1 > iWidth2 Or iHeight1 > iHeight2 Then
 frmDummy.pbxPic1.ScaleWidth = iWidth2
 frmDummy.pbxPic1.ScaleHeight = iHeight2
End If
'
lDiffPix = 0

For rYPos = 0 To iHeight1 - 1
 For rXPos = 0 To iWidth1 - 1
  lColor1 = frmDummy.pbxPic1.Point(rXPos, rYPos)
  lColor2 = frmDummy.pbxPic2.Point(rXPos, rYPos)
  
  If iScanMethod = 1 Or iScanMethod = 2 Then
   rgbSrc.Value = lColor1
   Call CopyMemory(rgbVal1, rgbSrc, 3)
   rgbSrc.Value = lColor2
   Call CopyMemory(rgbVal2, rgbSrc, 3)
   lDiffPix = lDiffPix + CalcToleranceM0(rgbVal1, rgbVal2)
  End If
  If iScanMethod = 3 Or iScanMethod = 4 Then
   lDiffPix = lDiffPix + CalcToleranceM1(lColor1, lColor2)
  End If
  If iScanMethod = 5 Or iScanMethod = 6 Then
   lDiffPix = lDiffPix + CalcToleranceM2(lColor1, lColor2)
  End If
  
  If lDiffPix > 0 And bFastScan Then
   Exit For
  End If
  
 Next rXPos
 If lDiffPix > 0 And bFastScan Then
  Exit For
 End If
Next rYPos
'
If Int(iScanMethod / 2) * 2 = iScanMethod Then
 lGlobP = CLng(iWidth1) * CLng(iHeight1)
 lGlobP = (lGlobP / 100) * iTolerance
 If lDiffPix < lGlobP Then
  lDiffPix = 0
 End If
End If
'
GetDiffPixels = lDiffPix
GoTo GetDiffPixels_End

GetDiffPixels_Error:
 GetDiffPixels = -1
 Exit Function
 
GetDiffPixels_End:
 End Function


Public Function GetMethodList(sMethodDesc() As String) As Integer
'
' Return the name of the available scan methods
'
On Error GoTo GetMethodList_Error
'
GetMethodList = 0
ReDim sMethodDesc(GetMethodNumber)

sMethodDesc(1) = "RGB Scan - Absolute Tolerance"
sMethodDesc(2) = "RGB Scan - Balanced Tolerance"
sMethodDesc(3) = "Color Scan - Absolute Tolerance"
sMethodDesc(4) = "Color Scan - Balanced Tolerance"
sMethodDesc(5) = "4 Shades - Absolute Tolerance"
sMethodDesc(6) = "4 Shades - Balanced Tolerance"
GoTo GetMethodList_End

GetMethodList_Error:
 GetMethodList = -1
 Exit Function
 
GetMethodList_End:
 End Function


Public Function GetMethodNumber() As Integer
'
' Return the number of method available
'
GetMethodNumber = 6
End Function


Private Function CalcToleranceM0(rgbVal1 As RGBpoint, rgbVal2 As RGBpoint) As Long
'
' Return 0 or 1 depending if the two point are identic or quite identic ( depending of tolerance )
' The test consist in comparing the RGB channels of each pixel
'

Dim iValMin As Integer
Dim iValMax As Integer
Dim iTmpVal As Integer
Dim iColT As Integer

CalcToleranceM0 = 1
iColT = 0
'
iValMin = rgbVal1.Blue - (iTolerance * 2.55)
iValMax = rgbVal1.Blue + (iTolerance * 2.55)
iTmpVal = rgbVal2.Blue


If iTmpVal >= iValMin And iTmpVal <= iValMax Then
 iColT = iColT + 1
End If

iValMin = rgbVal1.Green - (iTolerance * 2.55)
iValMax = rgbVal1.Green + (iTolerance * 2.55)
iTmpVal = rgbVal2.Green

If iTmpVal >= iValMin And iTmpVal <= iValMax Then
 iColT = iColT + 1
End If

iValMin = rgbVal1.Red - (iTolerance * 2.55)
iValMax = rgbVal1.Red + (iTolerance * 2.55)
iTmpVal = rgbVal2.Red

If iTmpVal >= iValMin And iTmpVal <= iValMax Then
 iColT = iColT + 1
End If

If iColT = 3 Then
 CalcToleranceM0 = 0
End If
End Function

Private Function CalcToleranceM1(lColor1 As Long, lColor2 As Long) As Long
'
' Return 0 or 1 depending if the two point are identic or quite identic ( depending of tolerance )
' The consist in comparing the weight of each color
'

Dim lValMin As Long
Dim lValMax As Long
Dim dColorP As Double
Dim lTmpVal As Long

CalcToleranceM1 = 1
'
dColorP = 167772.16 * iTolerance
lValMin = lColor1 - dColorP
lValMax = lColor1 + dColorP

If lColor2 >= lValMin And lColor2 <= lValMax Then
 CalcToleranceM1 = 0
End If
End Function

Private Function CalcToleranceM2(lColor1 As Long, lColor2 As Long) As Long
'
' Return 0 or 1 depending if the two point are identic or quite identic ( depending of tolerance )
' The test consist in converting the color in a range of 4 shades
'

Dim lValMin As Long
Dim lValMax As Long
Dim dColorP As Double
Dim lTmpVal As Long
'
CalcToleranceM2 = 1
'
dColorP = (COLOR_BASE / 100) * iTolerance
lValMin = CalcShade(4, lColor1 - dColorP)
lValMax = CalcShade(4, lColor1 + dColorP)
lTmpVal = CalcShade(4, lColor2)
If lTmpVal >= lValMin And lTmpVal <= lValMax Then
 CalcToleranceM2 = 0
End If

End Function


Private Function CalcShade(iNbShades As Integer, lCurrColor As Long) As Long
'
' Calculate the shade number of the color depending of rthe number of shades
'
Dim lShadeRange As Long
'
If iNbShades < 1 Then
 iNbShades = 1
End If
lShadeRange = COLOR_BASE / iNbShades
CalcShade = lCurrColor / lShadeRange
End Function

Public Property Let Tolerance(newvalue As Integer)
'
iTolerance = newvalue
If iTolerance < 0 Then
 iTolerance = 0
End If
If iTolerance > 100 Then
 iTolerance = 100
End If
End Property

Public Property Get Tolerance() As Integer
'
If iTolerance < 0 Then
 iTolerance = 0
End If
If iTolerance > 100 Then
 iTolerance = 100
End If
Tolerance = iTolerance
End Property

Public Property Let FastScan(newvalue As Boolean)
'
bFastScan = newvalue
End Property

Public Property Get FastScan() As Boolean
'
FastScan = bFastScan
End Property

Public Property Let ScanMethod(newvalue As Integer)
'
iScanMethod = newvalue
If iScanMethod < 1 Then
 iScanMethod = 1
End If
If iScanMethod > GetMethodNumber Then
 iScanMethod = GetMethodNumber
End If
If Int(iScanMethod / 2) * 2 = iScanMethod Then
 bFastScan = False
End If
End Property

Public Property Get ScanMethod() As Integer
'
If iScanMethod < 1 Then
 iScanMethod = 1
End If
If iScanMethod > GetMethodNumber Then
 iScanMethod = GetMethodNumber
End If
If Int(iScanMethod / 2) * 2 = iScanMethod Then
 bFastScan = False
End If
ScanMethod = iScanMethod
End Property
