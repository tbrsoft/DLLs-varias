VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   Caption         =   "Web Cam Capture"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   540
      Left            =   0
      Picture         =   "MainForm.frx":030A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   2340
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton btCamSet 
      Height          =   380
      Index           =   0
      Left            =   5265
      Picture         =   "MainForm.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Camera settings"
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   2790
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":081E
            Key             =   "AIDefault"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0B38
            Key             =   "AIGreen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0E52
            Key             =   "AIRed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":116C
            Key             =   "AIYellow"
         EndProperty
      EndProperty
   End
   Begin VB.Timer MainTimer 
      Interval        =   10000
      Left            =   315
      Top             =   3420
   End
   Begin VB.TextBox tStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1395
      Width           =   5595
   End
   Begin VB.CommandButton btStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Start the capture"
      Top             =   1755
      Width           =   5595
   End
   Begin VB.CommandButton btSettings 
      Caption         =   "Settings..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Capture settings"
      Top             =   900
      Width           =   5595
   End
   Begin VB.CommandButton btMonitor 
      Height          =   380
      Index           =   0
      Left            =   4860
      Picture         =   "MainForm.frx":1486
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Preview the camera capture"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox tDvcName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Device name"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton btDvcToggle 
      Caption         =   "Disabled"
      Height          =   380
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Enable the capture for the device"
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   45
      X2              =   5670
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Shape shIndicator 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   375
      Index           =   0
      Left            =   1035
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCamSet_Click(Index As Integer)
'
Dim bResult As Boolean
Dim idx As Integer
'
idx = Index + 1
If Not bDeviceSet(idx) Then
 bDeviceSet(idx) = True
 bResult = SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 1)
 bResult = SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 0)
 bDeviceSet(idx) = False
End If
End Sub

Private Sub btDvcToggle_Click(Index As Integer)
'
Dim sCaption As String
'
sCaption = Me.btDvcToggle(Index).Caption
If sCaption = "Enabled" Then
 Me.btDvcToggle(Index).Caption = "Disabled"
 Me.btDvcToggle(Index).ToolTipText = "Enable the capture for the device " & sDeviceName(Index + 1)
 Me.shIndicator(Index).BackColor = &H80FF&
 bDeviceEnabled(Index + 1) = False
End If
If sCaption = "Disabled" Then
 Me.btDvcToggle(Index).Caption = "Enabled"
 Me.btDvcToggle(Index).ToolTipText = "Disable the capture for the device " & sDeviceName(Index + 1)
 Me.shIndicator(Index).BackColor = &HFF00&
 bDeviceEnabled(Index + 1) = True
End If
Call SetStartButton(2)
End Sub

Private Sub btMonitor_Click(Index As Integer)
'
' find grabber output pin and render
'
On Error Resume Next
'
Dim bResult As Boolean
Dim idx As Integer
Dim lDvcState As Long
'
idx = Index + 1
clsGraph(idx).GetState 0, lDvcState
If lDvcState = 0 Then
 bResult = SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 2)
 clsGraph(idx).Run
End If
End Sub

Private Sub btSettings_Click()
'
frmSettings.Show
End Sub

Private Sub btStart_Click()
'
Dim sCaption As String
'
Call SetStartButton(1)
Call WriteCfgFile
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
'
Me.WindowState = 0
Me.Show
End Sub



Private Sub Form_Load()
'
' Load form
'
On Error Resume Next
'
    Dim lSize As Long
    Dim sTmpFName As String
    Dim bResult As Boolean
    Dim iNTSOption As Integer
    Dim iResp As Integer
    Dim sTmpText As String
    
    '
 
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    '
    ' Only if not starting as service ( timeout )
    If iCfgService = 0 Then
     Call SetupWindow
    End If
    '
    ' NT Service installation & start
    ' MUST take place in the "load" form event
    '

    iNTSOption = -1
    bEnableStart = False
'    If sCmdLine = "-install" Then
'     iNTSOption = 1
'     ntService.Interactive = True
'     If ntService.Install Then
'      iCfgService = 1
'      Call WriteCfgFile
'      sTmpText = "Service installed." & vbCrLf & "Close the application and run the service"
'      MsgBox sTmpText, vbInformation + vbOKOnly, "Service"
'      sCmdLine = ""
'     Else
'      sTmpText = "Problem with the service installation."
'      MsgBox sTmpText, vbCritical + vbOKOnly, "Service"
'     End If
'    End If
'    If sCmdLine = "-uninstall" Then
'     iNTSOption = 2
'     If ntService.Uninstall Then
'      sTmpText = "Service uninstalled." & vbCrLf & "Please close the application"
'      iCfgService = 0
'      Call WriteCfgFile
'      MsgBox sTmpText, vbInformation + vbOKOnly, "Service"
'      sCmdLine = ""
'     Else
'      sTmpText = "Can't uninstall the service."
'      MsgBox sTmpText, vbCritical + vbOKOnly, "Service"
'     End If
'    End If

    '
    ' Only if running as service
    '
'    If iCfgService = 1 Then
'     If sCmdLine = "-debug" Then
'      iNTSOption = 3
'      ntService.Debug = True
'      bEnableStart = True
'     End If
'     If sCmdLine = "" Then
'      iNTSOption = 0
'      bEnableStart = True
'     End If
'    End If
    '
    
    '
'    If iCfgService = 1 And bEnableStart Then
'     cSysTray1.InTray = True
'     ntService.ControlsAccepted = svcCtrlPauseContinue
'     ntService.StartService
'    End If
End Sub

Private Function SetupWindow()
'
' Setup the main window buttons and fields depending of the number of devices
'
On Error Resume Next
'
Dim iNbDvc As Integer
Dim idx As Integer
Dim fnum As Integer
'
Call CreateControls
Call SetControlsCfg
Call AlignControls
'
'Set up the filters for the different devices
iNbDvc = UBound(sDeviceName)
If iNbDvc > 0 Then
 For idx = 1 To iNbDvc
  If Not SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 0) Then
   Me.btDvcToggle(idx - 1).Caption = "Disabled"
   Me.btDvcToggle(idx - 1).Enabled = False
   Me.btMonitor(idx - 1).Enabled = False
   Me.shIndicator(idx - 1).BackColor = &HFF&
   bDeviceEnabled(idx) = False
  End If
 Next idx
End If
Call SetStartButton(0)
End Function




Private Function SetupVideoFilters(cCapStill As VBGrabber, cGraph As IMediaControl, _
                                   idx As Integer, iMode As Integer) As Boolean
'
' Setup the video driver depending of the mode
'
' 0 : normal ( no setup window, no render )
' 1 : preview ( render )
' 2 : setup ( setup window, no render )
'
On Error Resume Next
'

Dim bFilterFound As Boolean
Dim fiCurrent As IFilterInfo
Dim fiSource As IFilterInfo
Dim fltCurrent As IRegFilterInfo
Dim piOut As IPinInfo
Dim piIn As IPinInfo
Dim ppiOut As PinPropInfo
Dim scConfig As StreamConfig
Dim rfCurrent As Object
Dim sTmpFName As String
'
SetupVideoFilters = False
Set cCapStill = New VBGrabber
Set cGraph = New FilgraphManager
Set rfCurrent = cGraph.RegFilterCollection
' Add the grabber including vb wrapper and default props
For Each fltCurrent In rfCurrent
 If fltCurrent.Name = "SampleGrabber" Then
  fltCurrent.Filter fiCurrent
  ' wrap this filter in the capstill vb wrapper
  ' also sets rgb-24 media type and other properties
  cCapStill.FilterInfo = fiCurrent
  Exit For
 End If
Next fltCurrent
' add the selected source filter
bFilterFound = False
For Each fltCurrent In rfCurrent
 If fltCurrent.Name = sDeviceName(idx) Then
  fltCurrent.Filter fiSource
  bFilterFound = True
  Exit For
 End If
Next fltCurrent
If bFilterFound Then
 ' check for crossbar and select decoder
 Call CheckCrossBar(fiSource)
 ' find first output on src
 For Each piOut In fiSource.Pins
  If piOut.Direction = 1 Then
   Exit For
  End If
 Next piOut
 'restore previous cfg file ( if exist )
 Set scConfig = New StreamConfig
 scConfig.Pin = piOut
 If scConfig.SupportsConfig Then
  sTmpFName = App.Path & "\" & "dvc" & Trim(CStr(idx)) & "cfg.mt"
  If Dir(sTmpFName) <> "" Then
   scConfig.Restore (sTmpFName)
  End If
 End If
 '
 ' only if setup
 If iMode = 1 Then
 ' show format of output pin
  Set ppiOut = New PinPropInfo
  ppiOut.Pin = piOut
  ppiOut.ShowPropPage 0
  ' save selected format to file
  If scConfig.SupportsConfig Then
   scConfig.SaveCurrentFormat (sTmpFName)
  End If
 End If
 
 ' find first input on grabber and connect
 For Each piIn In fiCurrent.Pins
  If piIn.Direction = 0 Then
   piOut.Connect piIn
   Exit For
  End If
 Next piIn
 '
 ' Only if preview
 If iMode = 2 Then
  For Each piOut In fiCurrent.Pins
   If piOut.Direction = 1 Then
    piOut.Render
    Exit For
   End If
  Next piOut
 End If
End If

SetupVideoFilters = bFilterFound
End Function



Private Sub CheckCrossBar(fiSource As IFilterInfo)
' check for crossbar and select decoder (part1)
'
On Error GoTo CheckCrossBar_Error
'
Dim xbarInfo As CrossbarInfo
Dim lidx As Long
Dim sPin As String
'
Set xbarInfo = New CrossbarInfo
xbarInfo.SetFilter fiSource

For lidx = 0 To xbarInfo.Inputs - 1
 sPin = xbarInfo.Name(True, lidx)
Next lidx

If xbarInfo.Standard <> AnalogVideo_PAL_B Then
 xbarInfo.Standard = AnalogVideo_PAL_B
End If
GoTo CheckCrossBar_End

CheckCrossBar_Error:
 Exit Sub
 
CheckCrossBar_End:
 End Sub



Private Sub CreateControls()
'
' Create the additional controls more than one driver exists
'
On Error Resume Next
'
Dim idx As Integer
Dim iNBElements As Integer
'
iNBElements = UBound(sDeviceName) - 1
If iNBElements > 0 Then
 For idx = 1 To iNBElements
  Load Me.btDvcToggle(idx)
  Load Me.shIndicator(idx)
  Load Me.tDvcName(idx)
  Load Me.btMonitor(idx)
  Load Me.btCamSet(idx)
  Me.btDvcToggle(idx).Top = (idx * 480) + 240
  Me.btDvcToggle(idx).Visible = True
  Me.shIndicator(idx).Top = (idx * 480) + 240
  Me.shIndicator(idx).Visible = True
  Me.tDvcName(idx).Top = (idx * 480) + 240
  Me.tDvcName(idx).Visible = True
  Me.btMonitor(idx).Top = (idx * 480) + 240
  Me.btMonitor(idx).Visible = True
  Me.btCamSet(idx).Top = (idx * 480) + 240
  Me.btCamSet(idx).Visible = True
 Next idx
End If
End Sub

Private Sub SetControlsCfg()
'
' Set the controls config
'
Dim idx As Integer
Dim iNBElements As Integer
'
On Error Resume Next
'
iNBElements = Me.btDvcToggle.UBound
For idx = 0 To iNBElements
 If bDeviceEnabled(idx + 1) Then
  Me.btDvcToggle(idx).Caption = "Enabled"
  Me.shIndicator(idx).BackColor = &HFF00&
  Me.btDvcToggle(idx).ToolTipText = "Disable the capture for the device " & sDeviceName(idx + 1)
 Else
  Me.btDvcToggle(idx).Caption = "Disabled"
  Me.shIndicator(idx).BackColor = &H80FF&
  Me.btDvcToggle(idx).ToolTipText = "Enable the capture for the device " & sDeviceName(idx + 1)
 End If
 Me.tDvcName(idx).Text = sDeviceName(idx + 1)
Next idx
End Sub


Private Sub AlignControls()
'
' Align the controls on the form depending of the number of drivers found
'
On Error Resume Next
'
lBaseMFYPos = (UBound(sDeviceName) * 480) + 330
Me.Line1.Y1 = lBaseMFYPos
Me.btSettings.Top = lBaseMFYPos + 90
Me.tStatus.Top = lBaseMFYPos + 585
Me.btStart.Top = lBaseMFYPos + 945
'
Me.height = lBaseMFYPos + 2055
Me.width = 5835
End Sub

Private Sub SetStartButton(iMode As Integer)
'
' Set the start button and the status
'
On Error Resume Next
'
Dim idx As Integer
Dim bEnableSt As Boolean
Dim sCaption As String
'
sCaption = Me.btStart.Caption

bEnableSt = False
For idx = 1 To UBound(bDeviceEnabled)
 If bDeviceEnabled(idx) Then
  bEnableSt = True
 End If
Next idx
If bEnableSt Then
 If iMode <> 2 Then
  If sCaption = "Start" Then
   If (iMode = 1) Or (iMode = 0 And iCfgStart = 1) Then
    Call setSButtonToStart
    iCfgStart = 1
    bTmStart = True
   Else
    sCaption = "Stop"
   End If
  End If
  If sCaption = "Stop" Then
   Call setSButtonToStop
   iCfgStart = 0
   bTmStart = False
   lMTimer1 = 0
  End If
 Else
  If sCaption = "Start" Then
   Call setSButtonToStop
   iCfgStart = 0
   bTmStart = False
   lMTimer1 = 0
  End If
 End If
Else
 Call setSButtonToDisabled
 iCfgStart = 0
 bTmStart = False
 lMTimer1 = 0
End If

Call WriteCfgFile
End Sub

Private Sub setSButtonToStart()
'
' Set the start button to start
'
On Error Resume Next
'
Me.btStart.Caption = "Stop"
Me.btStart.Enabled = True
Me.tStatus.Text = "Started"
Me.tStatus.BackColor = &HFF00&
Me.btStart.ToolTipText = "Stop the capture"
'If Me.cSysTray1.InTray Then
' Set Me.cSysTray1.TrayIcon = Me.imlMain.ListImages(2).ExtractIcon
' Me.cSysTray1.TrayTip = "Web Cam Capture - RUNNING -"
'End If
End Sub

Private Sub setSButtonToStop()
'
' Set the start button to stop
'
On Error Resume Next
'
Me.btStart.Caption = "Start"
Me.btStart.Enabled = True
Me.btStart.ToolTipText = "Start the capture"
Me.tStatus.BackColor = &HFFFF&
Me.tStatus.Text = "Ready to start"
'If Me.cSysTray1.InTray Then
' Set Me.cSysTray1.TrayIcon = Me.imlMain.ListImages(4).ExtractIcon
' Me.cSysTray1.TrayTip = "Web Cam Capture - STOPPED -"
'End If
End Sub

Private Sub setSButtonToDisabled()
'
' Set the start button to disabled
'
On Error Resume Next
'
Me.btStart.Enabled = False
Me.btStart.Caption = "Start"
Me.tStatus.BackColor = &HFF&
Me.btStart.ToolTipText = ""
Me.tStatus.Text = "Enable camera(s)"
'If Me.cSysTray1.InTray Then
' Set Me.cSysTray1.TrayIcon = Me.imlMain.ListImages(3).ExtractIcon
' Me.cSysTray1.TrayTip = "Web Cam Capture - DISABLED -"
'End If
End Sub

Private Sub EnableUserButtons()
'
' Enable all the preview/cam settings buttons
'
On Error Resume Next
'
Dim idx As Integer
'
For idx = 0 To Me.btMonitor.UBound
 Me.btMonitor(idx).Enabled = True
 Me.btCamSet(idx).Enabled = True
Next idx
End Sub

Private Sub DisableUserButtons()
'
' Disable all the preview/cam settings buttons
'
On Error Resume Next
'
Dim idx As Integer
'
For idx = 0 To Me.btMonitor.UBound
 Me.btMonitor(idx).Enabled = False
 Me.btCamSet(idx).Enabled = False
Next idx
End Sub

Private Sub Form_Resize()
'
On Error Resume Next
'
Dim lMinHeight As Long
Dim lMinWidth As Long
Dim lFHeight As Long
Dim lFWidth As Long
Dim idx As Integer
Dim iDvcNum As Integer
Dim lDvcState As Long
Dim bResult As Boolean
'
Dim frmCurrent As Form
'
' resize/move the buttons/textboxes only if window in normal mode
'
If Me.WindowState = 0 Then
 lMinHeight = lBaseMFYPos + 2055
 lMinWidth = 5835
 If Me.height < lMinHeight Then
  Me.height = lMinHeight
 End If
 If Me.width < lMinWidth Then
  Me.width = lMinWidth
 End If
 lFHeight = Me.height
 lFWidth = Me.width
 '
 iDvcNum = UBound(sDeviceName)
 For idx = 1 To iDvcNum
  Me.btMonitor(idx - 1).Left = lFWidth - 975
  Me.btCamSet(idx - 1).Left = lFWidth - 570
  Me.tDvcName(idx - 1).width = lFWidth - 2340
 Next idx
 '
 Me.Line1.Y1 = lFHeight - 2055
 Me.Line1.Y2 = lFHeight - 2055
 Me.Line1.X2 = lFWidth - 75
 '
 Me.btSettings.Top = lFHeight - 1965
 Me.btSettings.width = lFWidth - 240
 '
 Me.tStatus.Top = lFHeight - 1470
 Me.tStatus.width = lFWidth - 240
 '
 Me.btStart.Top = lFHeight - 1080
 Me.btStart.width = lFWidth - 240
End If
'
' If the window is minimized and running as service :
' - unload linked forms
' - hide current form
'
If Me.WindowState = 1 Then
 If iCfgService = 1 Then
 ' stop the graphs only if the timer is not running
  If Not bTimerRunning Then
   For idx = 1 To UBound(sDeviceName)
    clsGraph(idx).GetState 0, lDvcState
    If lDvcState <> 0 Then
     clsGraph(idx).Stop
     bResult = SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 0)
    End If
   Next idx
  End If
  '
  For Each frmCurrent In Forms
   If frmCurrent.Caption = "Settings" Then
    If frmCurrent.Visible Then
     frmCurrent.Hide
    End If
    Unload frmCurrent
   End If
  Next frmCurrent
  Me.Hide
 End If
End If
End Sub

Private Sub Form_Terminate()
'
Call Form_Unload(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Terminate
'
Dim frmCurrent As Form
Dim idx As Integer
Dim sTmpFile As String
Dim iRep As Integer
Dim bUnload As Boolean
Dim lDvcState As Long
'
' Stop graphs
'
For idx = 1 To UBound(sDeviceName)
 clsGraph(idx).GetState 0, lDvcState
 If lDvcState <> 0 Then
  clsGraph(idx).Stop
 End If
Next idx
'
' Close forms
'
For Each frmCurrent In Forms
 If frmCurrent.Caption = "Settings" Then
  If frmCurrent.Visible Then
   frmCurrent.Hide
  End If
  Unload frmCurrent
 End If
Next frmCurrent
 
End Sub

Private Sub MainTimer_Timer()
'
' control the capture of the images
'
On Error Resume Next
'
Dim idx As Integer
Dim bResult As Boolean
Dim sTmpFile As String
Dim bTmpFile As Boolean
Dim lQuality As Long
Dim lDvcState As Long

'
DoEvents
If bTmStart Then
 If iServiceStarting = 10 Then
  'Me.cSysTray1.InTray = False
  'Me.cSysTray1.InTray = True
  iServiceStarting = 11
 End If
 If iServiceStarting >= 2 And iServiceStarting < 10 Then
  'Me.cSysTray1.InTray = True
 End If
 If iServiceStarting = 1 Then
  iServiceStarting = 2
  Call SetupWindow
 End If
 If Not bTimerRunning Then
  bTimerRunning = True
  lMTimer2 = Timer()
  If lMTimer1 > lMTimer2 Then
   lMTimer2 = lMTimer2 + 86400
  End If
  If lMTimer2 >= lMTimer1 + iCfgInterval Then
   Call DisableUserButtons
   For idx = 1 To UBound(bDeviceEnabled)
    If bDeviceEnabled(idx) Then
     clsGraph(idx).GetState 0, lDvcState
     If lDvcState = 2 Then
      clsGraph(idx).Stop
      bResult = SetupVideoFilters(clsCapStill(idx), clsGraph(idx), idx, 0)
     End If
     clsGraph(idx).Run
     ' Needed because some old cams take time to start...
     Call Sleep(3)
     Call SaveCapImage(idx, sTmpFile, bTmpFile)
     clsGraph(idx).Stop
     Call HandleCapImage(idx, sTmpFile, bTmpFile)
    End If
   Next idx
   lMTimer1 = Timer()
   Call EnableUserButtons
  End If
  bTimerRunning = False
 End If
End If
DoEvents
End Sub



Private Sub ntService_Continue(Success As Boolean)
'
' continue event for the service
End Sub


Private Sub ntService_Pause(Success As Boolean)
' pause event for the service
End Sub

'Private Sub ntService_Start(Success As Boolean)
''
'' This event MUST exist, otherwise the service stops...
''
'On Error GoTo ntService_Start_Error
''
'lMTimer1 = Timer()
'Success = True
'GoTo ntService_Start_End
'
'ntService_Start_Error:
' ntService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
' Exit Sub
'
'ntService_Start_End:
' ' Start the timer
' iServiceStarting = 1
' bTmStart = True
' End Sub

'Private Sub ntService_Stop()
''
'' Service stop
''
'On Error GoTo ntService_Stop_Error
'
'Call Form_Terminate
'GoTo ntService_Stop_End
'
'ntService_Stop_Error:
'ntService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
'Exit Sub
'
'ntService_Stop_End:
'End Sub


