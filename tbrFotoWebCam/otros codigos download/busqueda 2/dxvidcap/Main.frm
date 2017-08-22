VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "DirectShow 6.0 for VB"
   ClientHeight    =   5115
   ClientLeft      =   705
   ClientTop       =   1620
   ClientWidth     =   4710
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   465
      Left            =   1995
      TabIndex        =   5
      Top             =   825
      Width           =   1155
   End
   Begin VB.CheckBox chkFullScreen 
      Caption         =   "Full Screen Mode (Alt+Tab to restore)"
      Height          =   285
      Left            =   105
      TabIndex        =   4
      Top             =   1395
      Width           =   3030
   End
   Begin VB.ComboBox cbVidCapHW 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   3075
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Connect Preview Pin"
      Height          =   465
      Left            =   75
      TabIndex        =   1
      Top             =   825
      Width           =   1830
   End
   Begin VB.Label lblFilter 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   105
      TabIndex        =   3
      Top             =   15
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Select Video Capture Device:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   3015
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FilterGraph As FilgraphManager
Private m_VidWnd As IVideoWindow

Private Sub chkFullScreen_Click()
If Not m_VidWnd Is Nothing Then
    m_VidWnd.FullScreenMode = CBool(chkFullScreen.Value)
End If
End Sub

Private Sub cmdPause_Click()
If "Pause" = cmdPause.Caption Then
    m_FilterGraph.Pause
    cmdPause.Caption = "Resume"
Else
    m_FilterGraph.Run
    cmdPause.Caption = "Pause"
End If
End Sub

Private Sub cmdPreview_Click()
Dim ret As Boolean

Me.MousePointer = vbHourglass
Me.Caption = "Connecting to " & cbVidCapHW.Text
ret = CapPreviewConnect(cbVidCapHW.Text, m_FilterGraph)
If ret Then
    Me.Caption = "Connected! - Starting Preview..."
    m_VidWnd.Caption = "Live Window (size me)"
    m_VidWnd.Owner = Me.hWnd
    m_VidWnd.Left = 100
    m_VidWnd.Top = 100
    'Remove the caption, border, dialog frame, and scrollbars
    m_VidWnd.WindowStyle = WS_BORDER Or WS_CAPTION Or WS_THICKFRAME Or WS_CHILD Or WS_MAXIMIZEBOX
    m_FilterGraph.Run
    cmdPause.Enabled = True
    chkFullScreen.Enabled = True
Else
    MsgBox "Could not connect to capture driver", vbInformation, App.Title
End If

Me.Caption = App.Title
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim numdevs As Long

On Error Resume Next
Set m_FilterGraph = New FilgraphManager
If Err Then 'assume runtime not installed
    MsgBox "This program requires the DirectX 6.0 Media runtime files.", vbCritical, App.Title
    End 'kill prog
End If
Set m_VidWnd = m_FilterGraph
cmdPreview.Enabled = False
cmdPause.Enabled = False
chkFullScreen.Enabled = False
Me.Caption = "Enumerating Capture Filters..."
Me.Show
Me.Refresh
Me.MousePointer = vbHourglass
Me.Refresh
numdevs = EnumVideoCapHW(cbVidCapHW, lblFilter)
Me.MousePointer = vbDefault
Me.Caption = App.Title
If numdevs > 0 Then
    cmdPreview.Enabled = True
Else
    MsgBox "No video capture hardware detected", vbInformation, App.Title
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_FilterGraph = Nothing
Set m_VidWnd = Nothing
End Sub

