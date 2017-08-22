Attribute VB_Name = "MainModule"
Option Explicit
'
'
'
' Version history
' ===============
'
' 1.1 : - Can select to have a "time stamp" on the captured images
'       - Bug correction : when running as a service and the login procedure takes a long time,
'         the application was not visible on the systray
'
' 2.0 : - Compare the pictures ( simple motion detector )
'
'
' Video capture filter code
'
Public Const sVCFC = "{860BB310-5D01-11d0-BD3B-00A0C911CE86}"
'
' Function to convert bmp to jpg
'
Private Declare Function BMPToJPG Lib "converter.dll" (ByVal InputFilename As String, _
                         ByVal OutputFilename As String, ByVal Quality As Long) As Integer

' GDI functions to draw a DIBSection into a DC
Private Declare Function CreateCompatibleDC Lib "GDI32" _
    (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "GDI32" _
    (ByVal hdc As Long, ByVal hbitmap As Long) As Long
Private Declare Function BitBlt Lib "GDI32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal width As Long, ByVal height As Long, _
    ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal mode As Long) _
    As Long
Private Declare Sub DeleteDC Lib "GDI32" _
    (ByVal hdc As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, src As Any, ByVal count As Long)
    
' User and login functions
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long

'

'
Public clsCapStill() As VBGrabber
Public clsGraph() As IMediaControl
'
Public clsImgDiff As ImageDiff.ImgDiff
'
Public sDeviceName() As String
Public bDeviceEnabled() As Boolean
Public bDeviceSet() As Boolean
Public bDFirstPass() As Boolean
'
Public lBaseMFYPos As Long
'
Public sCfgFileName As String
'
Public iCfgInterval As Integer
Public iCfgMode As Integer
Public sCfgPath As String
Public sCfgFile As String
Public sCfgFileExt As String
Public iCfgFmtQuality As Integer
Public iCfgStart As Integer
Public iCfgService As Integer
Public iCfgTimeStamp As Integer
Public iCfgMDSwitch As Integer
Public iCfgMDMvOnly As Integer
Public iCfgMDTolerance As Integer
Public iCfgMDMethod As Integer
'
Public bTmStart As Boolean
Public bEnableStart As Boolean
Public iServiceStarting As Integer
'
Public lMTimer1 As Long
Public lMTimer2 As Long
Public bTimerRunning As Boolean
'
Public bIDOKPressed As Boolean
'
Public sCmdLine As String


Public Sub ReadCfgFile()
'
' Read the configuration file. If not found create one
'
On Error GoTo ReadCfgFile_Error
'
Dim fnum As Integer
Dim sFBuffer1 As String
Dim sLineSplit() As String
Dim sItem As String
Dim iDrvIdx As Integer
Dim iChkCfg As Integer
Dim iretval As Integer
'
iChkCfg = 0
fnum = FreeFile
Open sCfgFileName For Input Access Read As #fnum
Do Until EOF(fnum)
 Line Input #fnum, sFBuffer1
 sLineSplit = Split(sFBuffer1, ";", -1, vbBinaryCompare)
 sItem = UCase(sLineSplit(0))
 
  If sItem = "DEVICE" Then
   iDrvIdx = FindDevice(sLineSplit(1))
   If iDrvIdx > 0 Then
    If Trim(sLineSplit(2)) = "1" Then
     bDeviceEnabled(iDrvIdx) = True
    End If
   End If
  End If
 
  If sItem = "INTERVAL" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgInterval = Val(Trim(sLineSplit(1)))
    If iCfgInterval < 20 Then
     iCfgInterval = 20
    End If
    If iCfgInterval > 3600 Then
     iCfgInterval = 3600
    End If
   Else
    iCfgInterval = 60
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "MODE" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgMode = Val(Trim(sLineSplit(1)))
    If iCfgMode > 1 Or iCfgMode < 0 Then
     iCfgMode = 0
    End If
   Else
    iCfgMode = 0
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "FILEPATH" Then
   If Dir(Trim(sLineSplit(1)), vbDirectory) <> "" Then
    sCfgPath = Trim(sLineSplit(1))
    If Right(sCfgPath, 1) = "\" Then
     sCfgPath = Left(sCfgPath, Len(sCfgPath) - 1)
    End If
   Else
    sCfgPath = App.Path
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "FILEPREFIX" Then
   sCfgFile = Trim(sLineSplit(1))
   If sCfgFile = "" Then
    sCfgFile = "webcam"
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "FILEFORMAT" Then
   If (UCase(Trim(sLineSplit(1))) = "JPG") Or (UCase(Trim(sLineSplit(1))) = "BMP") Then
    sCfgFileExt = UCase(Trim(sLineSplit(1)))
    If sCfgFileExt = "BMP" Then
     iCfgFmtQuality = 100
    End If
    If sCfgFileExt = "JPG" Then
     iCfgFmtQuality = 75
     If UBound(sLineSplit) >= 3 Then
      If IsNumeric(Trim(sLineSplit(2))) Then
       iCfgFmtQuality = Val(Trim(sLineSplit(2)))
       If iCfgFmtQuality < 10 Then
        iCfgFmtQuality = 10
       End If
       If iCfgFmtQuality > 100 Then
        iCfgFmtQuality = 100
       End If
      End If
     End If
    Else
     sCfgFileExt = "BMP"
     iCfgFmtQuality = 100
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "TIMESTAMP" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgTimeStamp = Val(Trim(sLineSplit(1)))
    If iCfgTimeStamp <> 1 And iCfgTimeStamp <> 0 Then
     iCfgTimeStamp = 0
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  
  If sItem = "AUTOSTART" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgStart = Val(Trim(sLineSplit(1)))
    If iCfgStart <> 1 And iCfgStart <> 0 Then
     iCfgStart = 0
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "RUNASSERVICE" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgService = Val(Trim(sLineSplit(1)))
    If iCfgService <> 1 And iCfgService <> 0 Then
     iCfgService = 0
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "MDSWITCH" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgMDSwitch = Val(Trim(sLineSplit(1)))
    If iCfgMDSwitch <> 1 And iCfgMDSwitch <> 0 Then
     iCfgMDSwitch = 0
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "MDMVONLY" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgMDMvOnly = Val(Trim(sLineSplit(1)))
    If iCfgMDMvOnly <> 1 And iCfgMDMvOnly <> 0 Then
     iCfgMDMvOnly = 0
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "MDTOLERANCE" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgMDTolerance = Val(Trim(sLineSplit(1)))
    If iCfgMDTolerance < 0 Then
     iCfgMDTolerance = 0
    End If
    If iCfgMDTolerance > 100 Then
     iCfgMDTolerance = 100
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
  If sItem = "MDMETHOD" Then
   If IsNumeric(Trim(sLineSplit(1))) Then
    iCfgMDMethod = Val(Trim(sLineSplit(1)))
    If iCfgMDSwitch = 1 Then
     Set clsImgDiff = New ImgDiff
     iretval = clsImgDiff.GetMethodNumber
     Set clsImgDiff = Nothing
     If iCfgMDMethod < 1 Then
      iCfgMDMethod = 1
     End If
     If iCfgMDMethod > iretval Then
      iCfgMDMethod = iretval
     End If
    Else
     iCfgMDMethod = 1
    End If
   End If
   iChkCfg = iChkCfg + 1
  End If
  
Loop

Close #fnum
If iChkCfg <> 12 Then
 Call WriteDefCfgFile
End If
GoTo ReadCfgFile_End

ReadCfgFile_Error:
 Call WriteDefCfgFile
 Exit Sub
 
ReadCfgFile_End:
 End Sub

Public Function FindDevice(sDevice As String) As Integer
'
' Find the device name. Return the index of the driver or 0 if the driver is not found
'
Dim idx As Integer
'
FindDevice = 0
For idx = 1 To UBound(sDeviceName)
 If sDeviceName(idx) = sDevice Then
  FindDevice = idx
  Exit For
 End If
Next idx
End Function


Public Sub WriteDefCfgFile()
'
' Write a default configuration file
'
iCfgInterval = 60
iCfgMode = 0
sCfgPath = App.Path
sCfgFile = "webcam"
sCfgFileExt = "BMP"
iCfgFmtQuality = 100
iCfgStart = 0
iCfgTimeStamp = 0
iCfgService = 0
iCfgMDSwitch = 0
iCfgMDMvOnly = 0
iCfgMDTolerance = 0
iCfgMDMethod = 1
Call WriteCfgFile
End Sub

Public Sub WriteCfgFile()
'
' Write the configuration file
'
On Error GoTo WriteCfgFile_Error
'
Dim fnum As Integer
Dim sFBuffer As String
Dim idx As Integer
'
fnum = FreeFile
Open sCfgFileName For Output Access Write As #fnum
For idx = 1 To UBound(sDeviceName)
 sFBuffer = "DEVICE;" & sDeviceName(idx) & ";"
 If bDeviceEnabled(idx) Then
  sFBuffer = sFBuffer & "1"
 Else
  sFBuffer = sFBuffer & "0"
 End If
 Print #fnum, sFBuffer
Next idx
sFBuffer = "INTERVAL;" & Trim(CStr(iCfgInterval))
Print #fnum, sFBuffer
sFBuffer = "MODE;" & Trim(CStr(iCfgMode))
Print #fnum, sFBuffer
sFBuffer = "FILEPATH;" & Trim(sCfgPath)
Print #fnum, sFBuffer
sFBuffer = "FILEPREFIX;" & Trim(sCfgFile)
Print #fnum, sFBuffer
sFBuffer = "FILEFORMAT;" & UCase(Trim(sCfgFileExt)) & ";" & Trim(CStr(iCfgFmtQuality))
Print #fnum, sFBuffer
sFBuffer = "TIMESTAMP;" & Trim(CStr(iCfgTimeStamp))
Print #fnum, sFBuffer
sFBuffer = "AUTOSTART;" & Trim(CStr(iCfgStart))
Print #fnum, sFBuffer
sFBuffer = "RUNASSERVICE;" & Trim(CStr(iCfgService))
Print #fnum, sFBuffer
sFBuffer = "MDSWITCH;" & Trim(CStr(iCfgMDSwitch))
Print #fnum, sFBuffer
sFBuffer = "MDMVONLY;" & Trim(CStr(iCfgMDMvOnly))
Print #fnum, sFBuffer
sFBuffer = "MDTOLERANCE;" & Trim(CStr(iCfgMDTolerance))
Print #fnum, sFBuffer
sFBuffer = "MDMETHOD;" & Trim(CStr(iCfgMDMethod))
Print #fnum, sFBuffer
Close #fnum

GoTo WriteCfgFile_End

WriteCfgFile_Error:
 If iCfgService <> 1 Then
  MsgBox "Unable to write configuration file", vbCritical + vbOKOnly, "Error"
 End If
 Exit Sub

WriteCfgFile_End:
 End Sub

Public Sub SaveCapImage(idx As Integer, sTmpFile As String, bTmpFile As Boolean)
'
' Save the captured image
'
On Error Resume Next
'
bTmpFile = False
sTmpFile = ""

If sCfgFileExt = "JPG" Or iCfgTimeStamp = 1 Then
 sTmpFile = App.Path & "\" & "tmppct" & Trim(CStr(idx)) & ".bmp"
 bTmpFile = True
Else
 sTmpFile = GetRealFileName(idx)
 bTmpFile = False
End If
clsCapStill(idx).FileName = sTmpFile
clsCapStill(idx).CaptureStill
End Sub


Public Sub HandleCapImage(idx As Integer, sTmpFile As String, bTmpFile As Boolean)
'
' Handle the captured images
'
On Error Resume Next
'
Dim sOldFile As String
Dim sTmpText As String
Dim lQuality As Long
Dim iretval As Integer
Dim lretval As Long
Dim bMDChanges As Boolean
Dim bMDRChanges As Boolean
'
Dim pPicture1 As StdPicture
Dim pPicture2 As StdPicture
'
bMDChanges = True
bMDRChanges = False
'
sOldFile = App.Path & "\" & "oldpct" & Trim(CStr(idx)) & ".bmp"
' if motion detection enabled
If iCfgMDSwitch = 1 And Dir(sOldFile) <> "" Then
 If Not bDFirstPass(idx) Then
  Set clsImgDiff = New ImgDiff
  clsImgDiff.Tolerance = iCfgMDTolerance
  clsImgDiff.FastScan = True
  clsImgDiff.ScanMethod = iCfgMDMethod
  Set pPicture1 = LoadPicture(sOldFile)
  Set pPicture2 = LoadPicture(sTmpFile)
  lretval = clsImgDiff.GetDiffPixels(pPicture1, pPicture2)
  If lretval = 0 And iCfgMDMvOnly = 1 Then
   bMDChanges = False
  End If
  If lretval > 0 Then
   bMDRChanges = True
  End If
  Set clsImgDiff = Nothing
 End If
End If
'
If iCfgMDSwitch = 1 Then
 iretval = GeneralCopy(sTmpFile, sOldFile)
End If
'
If bMDChanges Then
 If iCfgTimeStamp = 1 Then
  MainForm.pctTemp = LoadPicture(sTmpFile)
  MainForm.pctTemp.AutoRedraw = True
  MainForm.pctTemp.CurrentX = 5
  MainForm.pctTemp.CurrentY = 5
  sTmpText = CStr(Format(Now(), "dd/mm/yyyy HH:MM:SS"))
  If bMDRChanges Then
   sTmpText = sTmpText & "*"
  End If
  MainForm.pctTemp.Print sTmpText
  MainForm.pctTemp.Refresh
  SavePicture MainForm.pctTemp.Image, sTmpFile
 End If
 If sCfgFileExt = "JPG" Then
  lQuality = iCfgFmtQuality
  iretval = BMPToJPG(sTmpFile, GetRealFileName(idx), lQuality)
 End If
End If
'
If Dir(sTmpFile) <> "" And bTmpFile Then
 Kill (sTmpFile)
End If
'
bDFirstPass(idx) = False

End Sub

Public Function GetRealFileName(idx As Integer) As String
'
' Return the complete file path for the capture file
'
Dim sCamNum As String
Dim sTimeStamp As String
'
GetRealFileName = ""
sTimeStamp = ""
sCamNum = Format(idx, "000")
If iCfgMode = 1 Then
 sTimeStamp = "_" & Format(Now(), "yyyymmddHHMMSS")
End If
GetRealFileName = sCfgPath & "\" & sCfgFile & sCamNum & sTimeStamp & "." & sCfgFileExt
End Function

Public Function GetDeviceNames() As Boolean
'
' Get the list of existing devices
'
Dim bResult As Boolean
Dim idx As Integer
Dim sDrvName As String
Dim ivcCurrent As IVBCollection
Dim fceCurrent As FilterCatEnumerator
Dim fcCurrent As IFilterClass
'
GetDeviceNames = False
bResult = True
idx = 0
Set fceCurrent = New FilterCatEnumerator
Set ivcCurrent = fceCurrent.Category(sVCFC)
idx = 0
For Each fcCurrent In ivcCurrent
 idx = idx + 1
 ReDim Preserve clsCapStill(idx)
 ReDim Preserve clsGraph(idx)
 ReDim Preserve sDeviceName(idx)
 ReDim Preserve bDeviceEnabled(idx)
 ReDim Preserve bDeviceSet(idx)
 ReDim Preserve bDFirstPass(idx)
 sDeviceName(idx) = Replace(fcCurrent.Name, ";", "_", 1, -1, vbBinaryCompare)
 bDeviceEnabled(idx) = False
 bDeviceSet(idx) = False
 bDFirstPass(idx) = True
 GetDeviceNames = True
Next fcCurrent
End Function
'
Public Sub Sleep(lSec As Long)
'
' Sleep
'
Dim lTimer1 As Long
Dim lTimer2 As Long
Dim lTimer3 As Long
'
lTimer1 = Timer()
lTimer2 = lTimer1 + lSec
lTimer3 = lTimer1
Do While lTimer2 > lTimer1
 DoEvents
 lTimer1 = Timer()
 If lTimer1 < lTimer3 Then
  lTimer1 = lTimer1 + 86400
 End If
Loop
End Sub

Public Function GeneralCopy(sFNamefr As String, sFNameto As String) As Integer
'
' Execute a "copy file"
'
On Error GoTo GeneralCopyError
'
Dim fbuffer() As Byte
Dim fnum As Integer
'
GeneralCopy = 0
fnum = FreeFile
Open sFNamefr For Binary Access Read As #fnum
ReDim fbuffer(LOF(fnum))
Get #fnum, , fbuffer
Close #fnum
fnum = FreeFile
Open sFNameto For Binary Access Write As #fnum
Put #fnum, , fbuffer
Close #fnum
GoTo GeneralCopyEnd

GeneralCopyError:
 GeneralCopy = 1
 Exit Function
 
GeneralCopyEnd:
 End Function


Public Sub Main()
'
' Main routine
'
On Error Resume Next
'
Dim bResult As Boolean
Dim fnum As Integer
'
bTmStart = False
bTimerRunning = False
lMTimer1 = 0
lMTimer2 = 0
sCfgFileName = App.Path & "\" & "wcamcap.cfg"
bResult = GetDeviceNames()
If bResult Then
 Call ReadCfgFile
 If iCfgService = 0 Then
  MainForm.WindowState = 0
  MainForm.Show
 End If
 If iCfgService = 1 Then
  MainForm.WindowState = 1
 End If
Else
 If iCfgService <> 1 Then
  MsgBox "No compatible driver found.", vbCritical + vbOKOnly, "Error"
 End If
End If
End Sub
