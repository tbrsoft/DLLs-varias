VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WebCam Motion Capture"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      DrawWidth       =   3
      Height          =   5985
      Left            =   900
      ScaleHeight     =   5925
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6000
      Top             =   8760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FOR WEBCAM DECLARATIONS
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

Option Explicit

Private Sub Command1_Click()
    If Timer1.Enabled Then
        STARTCAM
    Else
        STOPCAM
    End If
End Sub

Private Sub Form_Load()
    'set up the visual stuff
    'Picture1.Width = 640 * Screen.TwipsPerPixelX
    'Picture1.Height = 480 * Screen.TwipsPerPixelY
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        STARTCAM
    ElseIf Button = 2 Then
        STOPCAM
    End If
End Sub

Private Sub Timer1_Timer()
    'Get the picture from camera.. the main part
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    Picture1.Picture = Clipboard.GetData
    Clipboard.Clear
End Sub

Sub STOPCAM()
    DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
    Timer1.Enabled = False
End Sub

Sub STARTCAM()
    mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hwnd, 0)
    DoEvents
    SendMessage mCapHwnd, CONNECT, 0, 0
    Timer1.Enabled = True
End Sub
