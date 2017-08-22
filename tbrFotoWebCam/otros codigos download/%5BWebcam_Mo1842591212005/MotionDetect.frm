VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WebCam Motion Capture"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      DrawWidth       =   3
      Height          =   6975
      Left            =   150
      ScaleHeight     =   6915
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   30
      Width           =   10095
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6000
      Top             =   8760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "left click to record, right click to stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   8280
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Firstly id like to say that im no expert with webcams, i only know how to capture and stop it
'The major aim of this code, for me, was to make the motion detection algorithm itself

'FOR WEBCAM DECLARATIONS
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

'declarations
Dim P() As Long
Dim POn() As Boolean

Dim inten As Integer

Dim i As Integer, j As Integer

Dim Ri As Long, Wo As Long
Dim RealRi As Long

Dim c As Long, c2 As Long

Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

Dim Tppx As Single, Tppy As Single
Dim Tolerance As Integer

Dim RealMov As Integer

Dim Counter As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long

Option Explicit

Private Sub Form_Load()
'set up the visual stuff
Picture1.Width = 640 * Screen.TwipsPerPixelX
Picture1.Height = 480 * Screen.TwipsPerPixelY

'Inten is the measure of how many pixels are going to be recognized. I highly dont recommend
'setting it lower than this, i have a 3.0 GHz PC and it starts to lag a little. On this setting,
'every 15th pixel is checked
inten = 15
'The tolerance of recognizing the pixel change
Tolerance = 20

Tppx = Screen.TwipsPerPixelX
Tppy = Screen.TwipsPerPixelY

ReDim POn(640 / inten, 480 / inten)
ReDim P(640 / inten, 480 / inten)

STARTCAM
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

Ri = 0 'right
Wo = 0 'wrong

LastTime = GetTickCount

For i = 0 To 640 / inten - 1
    For j = 0 To 480 / inten - 1
    'get a point
    c = Picture1.Point(i * inten * Tppx, j * inten * Tppy)
    'analyze it, Red, Green, Blue
        R = c Mod 256
        G = (c \ 256) Mod 256
        B = (c \ 256 \ 256) Mod 256

    'recall what the point was one step before this
    c2 = P(i, j)
        'analyze it
        R2 = c2 Mod 256
        G2 = (c2 \ 256) Mod 256
        B2 = (c2 \ 256 \ 256) Mod 256

    'main comparison part... if each R, G and B are somewhat same, then it pixel is same still
    'in a perfect camera and software tolerance should theoretically be 1 but this isnt true...
    If Abs(R - R2) < Tolerance And Abs(G - G2) < Tolerance And Abs(B - B2) < Tolerance Then
    'pixel remained same
    Ri = Ri + 1
    'Pon stores a boolean if the pixel changed or didnt, to be used to detect REAL movement
    POn(i, j) = True

    Else
    'Pixel changed
    Wo = Wo + 1
    'make a red dor
    P(i, j) = Picture1.Point(i * inten * Tppx, j * inten * Tppy)
    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbRed
    POn(i, j) = False
    End If

    Next j

Next i

RealRi = 0

For i = 1 To 640 / inten - 2
    For j = 1 To 480 / inten - 2
    If POn(i, j) = False Then
        'Real movement is simply occuring when all 4 pixels around one pixel changed
        'Simply put, If this pixel is changed and all around it changed too, then this is a real
        'movement
        If POn(i, j + 1) = False Then
            If POn(i, j - 1) = False Then
                If POn(i + 1, j) = False Then
                    If POn(i - 1, j) = False Then
                    RealRi = RealRi + 1
                    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbGreen
                    End If
                End If
            End If
        End If

    End If


    Next j
Next i

'state all statistics
Label1.Caption = Int(Wo / (Ri + Wo) * 100) & " % movement" & vbCrLf & "Real Movement: " & RealRi & vbCrLf _
& "Completed in: " & GetTickCount - LastTime

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

'Originally for saving the picture output... i left it out since it has nothing to do with program
'i commented just so you can see how its done in case you dont know
'Private Sub Timer2_Timer()
'SavePicture Picture1.Image, "C:\pics\img" & Counter & ".bmp"
'Counter = Counter + 1
'End Sub
