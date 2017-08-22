VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DirectX WAV / MIDI Player"
   ClientHeight    =   1740
   ClientLeft      =   156
   ClientTop       =   432
   ClientWidth     =   4920
   Icon            =   "frmAudio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Master Volume"
      Height          =   675
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   2295
      Begin MSComctlLib.Slider sldVolume 
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   420
         Width           =   1995
         _ExtentX        =   3514
         _ExtentY        =   339
         _Version        =   393216
         LargeChange     =   200
         SmallChange     =   100
         Min             =   -1000
         Max             =   200
         SelStart        =   200
         TickFrequency   =   200
         Value           =   200
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         Height          =   255
         Index           =   3
         Left            =   1860
         TabIndex        =   10
         Top             =   180
         Width           =   315
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.Frame fraTempo 
      Caption         =   "Tempo"
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2295
      Begin MSComctlLib.Slider sldTempo 
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   1995
         _ExtentX        =   3514
         _ExtentY        =   339
         _Version        =   393216
         Max             =   30
         SelStart        =   10
         TickFrequency   =   5
         Value           =   10
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Fast"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   14
         Top             =   180
         Width           =   375
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   540
         TabIndex        =   13
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Slow"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   540
      Width           =   975
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop Audio"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   3675
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Audio File"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   540
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4
   End
End
Attribute VB_Name = "frmAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Implements DirectXEvent8

Private dx As New DirectX8
Private dmp As DirectMusicPerformance8
Private dml As DirectMusicLoader8
Private dmSeg As DirectMusicSegment8
Private dmEvent As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdOpen_Click()
  Static sCurDir As String
  Static lFilter As Long
  cdlOpen.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
  cdlOpen.FilterIndex = lFilter
  cdlOpen.Filter = "Wave Files (*.wav)|*.wav|Music Files (*.mid;*.rmi)|*.mid;*.rmi|Segment Files (*.sgt)|*.sgt|All Audio Files|*.wav;*.mid;*.rmi;*.sgt|All Files (*.*)|*.*"
  cdlOpen.FileName = vbNullString
  If sCurDir = vbNullString Then
    Dim sWindir As String
    sWindir = Space$(255)
    If GetWindowsDirectory(sWindir, 255) = 0 Then
      cdlOpen.InitDir = "C:\"
    Else
      Dim sMedia As String
      sWindir = Left$(sWindir, InStr(sWindir, Chr$(0)) - 1)
      If Right$(sWindir, 1) = "\" Then
        sMedia = sWindir & "Media"
      Else
        sMedia = sWindir & "\Media"
      End If
      If Dir$(sMedia, vbDirectory) <> vbNullString Then
        cdlOpen.InitDir = sMedia
      Else
        cdlOpen.InitDir = sWindir
      End If
    End If
  Else
    cdlOpen.InitDir = sCurDir
  End If
  On Local Error GoTo ClickedCancel
  cdlOpen.CancelError = True
  cdlOpen.ShowOpen
  isOK = False
  frmAudio.Enabled = False
  Dialog.Show
  Do
    DoEvents
  Loop Until isOK
  CleanupAudio
  InitAudio
  frmAudio.Enabled = True
  sCurDir = GetFolder(cdlOpen.FileName)
  dml.SetSearchDirectory sCurDir
  lFilter = cdlOpen.FilterIndex
  On Local Error GoTo NoLoadSegment
  cmdStop_Click
  If FileLen(cdlOpen.FileName) = 0 Then Err.Raise 5
  EnableTempoControl (Right$(cdlOpen.FileName, 4) <> ".wav")
  Set dmSeg = dml.LoadSegment(cdlOpen.FileName)
  If (Right$(cdlOpen.FileName, 4) = ".mid") Or (Right$(cdlOpen.FileName, 4) = ".rmi") Or (Right$(cdlOpen.FileName, 5) = ".midi") Then
    dmSeg.SetStandardMidiFile
  End If
  txtFile.Text = cdlOpen.FileName
  EnablePlayUI True
  sldTempo.Value = 10
  sldTempo_Click
  cmdPlay_Click
  Exit Sub
NoLoadSegment:
  MsgBox "Couldn't load this segment", vbOKOnly Or vbCritical, "Couldn't load"
ClickedCancel:
End Sub

Private Sub cmdPlay_Click()
  If Not (dmSeg Is Nothing) Then
    If chkLoop.Value = vbChecked Then
      dmSeg.SetRepeats -1 'Loop infinitely
    Else
      dmSeg.SetRepeats 0 'Don't loop
    End If
    dmp.PlaySegmentEx dmSeg, DMUS_SEGF_DEFAULT, 0
    EnablePlayUI False
  End If
End Sub

Private Sub cmdStop_Click()
  If Not (dmSeg Is Nothing) Then dmp.StopEx dmSeg, 0, 0
  EnablePlayUI True
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
  Dim dmNotification As DMUS_NOTIFICATION_PMSG
  If Not dmp.GetNotificationPMSG(dmNotification) Then
    MsgBox "Error processing this Notification", vbOKOnly Or vbInformation, "Cannot Process."
    Exit Sub
  Else
    If dmNotification.lNotificationOption = DMUS_NOTIFICATION_SEGEND Then 'The segment has ended
      EnablePlayUI True
    End If
  End If
End Sub

Private Sub Form_Load()
  InitAudio
  EnableTempoControl False
End Sub

Private Sub InitAudio()
  On Error GoTo FailedInit
  Set dmp = dx.DirectMusicPerformanceCreate
  Set dml = dx.DirectMusicLoaderCreate
  Dim dmusAudio As DMUS_AUDIOPARAMS
  Select Case AudioMode
    Case 1
      dmp.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_DYNAMIC_MONO, 128
    Case 2
      dmp.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    Case 3
      dmp.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_DYNAMIC_3D, 128
    Case 4
      dmp.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
  End Select
  dmp.SetMasterAutoDownload True
  dmp.AddNotificationType DMUS_NOTIFY_ON_SEGMENT
  dmEvent = dx.CreateEvent(Me)
  dmp.SetNotificationHandle dmEvent
  Exit Sub
FailedInit:
  MsgBox "Could not initialize DirectMusic." & vbCrLf & "will now exit.", vbOKOnly Or vbInformation, "Exiting..."
  CleanupAudio
  Unload Me
  End
End Sub

Private Sub CleanupAudio()
  On Error Resume Next
  dmp.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
  dx.DestroyEvent dmEvent
  If Not (dmSeg Is Nothing) Then dmp.StopEx dmSeg, 0, 0
  Set dmSeg = Nothing
  Set dml = Nothing
  If Not (dmp Is Nothing) Then dmp.CloseDown
  Set dmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CleanupAudio
  End
End Sub

Private Function GetFolder(ByVal sFile As String) As String
  Dim lCount As Long
  For lCount = Len(sFile) To 1 Step -1
    If Mid$(sFile, lCount, 1) = "\" Then
      GetFolder = Left$(sFile, lCount)
      Exit Function
    End If
  Next
  GetFolder = vbNullString
End Function

Public Sub EnablePlayUI(fEnable As Boolean)
  If fEnable Then
    chkLoop.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = True
    cmdPlay.SetFocus
  Else
    chkLoop.Enabled = False
    cmdStop.Enabled = True
    cmdPlay.Enabled = False
    cmdStop.SetFocus
  End If
End Sub

Private Sub sldTempo_Click()
  dmp.SetMasterTempo (sldTempo.Value / 10)
End Sub

Private Sub sldTempo_Scroll()
  sldTempo_Click
End Sub

Private Sub sldVolume_Click()
  sldVolume_Scroll
End Sub

Private Sub sldVolume_Scroll()
  dmp.SetMasterVolume sldVolume.Value
End Sub

Private Sub EnableTempoControl(ByVal fEnable As Boolean)
  'If this is a wave file, turn off tempo control
  fraTempo.Enabled = fEnable
  sldTempo.Enabled = fEnable
  lbl(4).Enabled = fEnable
  lbl(5).Enabled = fEnable
  lbl(6).Enabled = fEnable
  If Not fEnable Then
    sldTempo.TickStyle = sldNoTicks
  Else
    sldTempo.TickStyle = sldBottomRight
  End If
End Sub

