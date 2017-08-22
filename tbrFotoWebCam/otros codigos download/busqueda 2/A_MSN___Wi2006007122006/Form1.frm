VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN Webcam Capture - Stopped"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Recording"
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   5055
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   30
         Text            =   "1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   28
         Text            =   "15"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "ms"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Minimum time to wait before saving new image:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label9 
         Caption         =   "Compare px steps (Lower is better but takes a lot longer):"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.PictureBox l2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox d 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin VB.FileListBox f 
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Text            =   "1"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Preview"
      Height          =   4095
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "cls"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox p 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3780
         Left            =   120
         ScaleHeight     =   3750
         ScaleWidth      =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   4830
         Begin VB.CommandButton Command3 
            BackColor       =   &H000000FF&
            Caption         =   "Enable this function (required)"
            Height          =   615
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3000
            Width           =   4575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "lbl"
            Height          =   3015
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   4575
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cropping"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Text            =   "252"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Text            =   "322"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "149"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "351"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "height"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "width"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "px from top"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "px from right"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Config"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "IMWindowClass"
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Text            =   " - Conversation"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Filename start"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Window class"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Window title"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim isRecord As Boolean
Dim oldPX As String
Dim lastCap As Double

Private Sub Command1_Click()
If Command1.Caption = "Record" Then
    Command1.Caption = "Stop"
    isRecord = True
    Me.Caption = "MSN Webcam Capture - Recording"
Else
    Command1.Caption = "Record"
    isRecord = False
    Me.Caption = "MSN Webcam Capture - Stopped"
    DoEvents
    DoEvents
    p.Cls
    p.Print "End of record segment."
    SavePicture p.Image, "C:\Webcam\img" & Text9.Text & ".bmp"
    Text9.Text = Text9.Text + 1
End If
End Sub

Private Sub Command2_Click()
p.Cls
End Sub

Private Sub Command3_Click()
Label3.Visible = False
Command3.Visible = False
Command1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = Screen.Height - Me.Height - 480
On Error Resume Next
MkDir "C:\Webcam\"
DoEvents
f.Path = "C:\Webcam\"
DoEvents

Dim tHigh As Integer
For i = 0 To f.ListCount - 1
    a = Mid(f.List(i), 4)
    a = Mid(a, 1, InStr(a, ".") - 1)
    If IsNumeric(Int(a)) Then
        If Int(a) > tHigh Then
            tHigh = Int(a)
        End If
    End If
Next
Text9 = tHigh + 1

Label3.Caption = "Warning: " & vbCrLf & vbCrLf & "This program makes heavy use of the Windows clipboard.  Images are taken using the keyboard API 'keybd_event' from the 'user32.dll' file to simulate the pressing of the 'Print Screen' key." & vbCrLf & vbCrLf & _
"The result is the current screen saved to the clipboard, ERASING ANY DATA YOU CURRENTLY HAVE COPIED." & vbCrLf & "This process takes place every one-hundredth of a second" & vbCrLf & "(10 milliseconds)." & vbCrLf & vbCrLf & _
"Please save any of your data on the Windows clipboard before proceeding." & vbCrLf & "                                                                                         - Lynxy"
Command1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Timer1.Enabled = False
If Text5 = "" Then Text5 = 0
If Text6 = "" Then Text6 = 0
If Text7 = "" Then Text7 = 0
If Text8 = "" Then Text8 = 0
If p.Tag = "" Then p.Tag = 0

Dim tRC As RECT

Handle1 = FindWindow(Text3.Text, Text1.Text & Text2.Text)
If Handle1 = 0 Then
    p.Cls
    p.FontSize = 12
    p.CurrentX = 900
    p.CurrentY = 1300
    p.Print "Could not find the window."
    
    p.CurrentX = 1200
    p.CurrentY = 1600
    p.Print "Check your settings."
    GoTo ExitSub
End If
GetWindowRect Handle1, tRC

Call keybd_event(vbKeySnapshot, 0, 0, 0)
DoEvents
d = Clipboard.GetData(vbCFBitmap)
'p.Cls
DoEvents
p.PaintPicture d, 0, 0, p.Width, p.Height, 15 * (tRC.Right - Text5.Text), 15 * (tRC.Top + Text6.Text), Text7.Text * 15, Text8.Text * 15
DoEvents
If isRecord Then
    For i = 1 To p.Width Step 15 * 15
        For j = 1 To p.Height Step 15 * 15
            tmpPX = tmpPX & p.Point(i, j) & ","
        Next
    Next
    If tmpPX <> oldPX Then
        If Timer - lastCap > Text10 / 1000 Then
            lastCap = Timer
            SavePicture p.Image, "C:\Webcam\img" & Text9.Text & ".bmp"
            Text9.Text = Text9.Text + 1
            oldPX = ""
            For i = 1 To p.Width Step 15 * 15
                For j = 1 To p.Height Step 15 * 15
                    oldPX = oldPX & p.Point(i, j) & ","
                Next
            Next
        End If
    End If
End If
ExitSub:
Timer1.Enabled = True
End Sub
