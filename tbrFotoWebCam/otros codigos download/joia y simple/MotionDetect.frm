VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WebCam Motion Capture"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      DrawWidth       =   3
      Height          =   3195
      Left            =   480
      ScaleHeight     =   3135
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   450
      Width           =   4755
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'FOR WEBCAM DECLARATIONS
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

Public Sub STOPCAM()
    DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
    Timer1.Enabled = False
End Sub

Public Sub STARTCAM(HW As Long, H As Long, W As Long)
    mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, W, H, HW, 0)
    DoEvents
    SendMessage mCapHwnd, CONNECT, 0, 0
    Timer1.Enabled = True
End Sub

Public Sub upFoto()
    'Get the picture from camera.. the main part
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    Picture1.Picture = Clipboard.GetData
    Clipboard.Clear
End Sub

Private Sub Form_Load()
    Picture1.Width = 640 * Screen.TwipsPerPixelX
    Picture1.Height = 480 * Screen.TwipsPerPixelY
    STARTCAM Me.hwnd, 640, 480
    Timer1.Enabled = False
    Timer1.Interval = 300
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    STOPCAM
End Sub

Private Sub Timer1_Timer()
    upFoto
End Sub
