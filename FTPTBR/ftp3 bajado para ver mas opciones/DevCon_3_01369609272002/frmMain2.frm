VERSION 5.00
Begin VB.Form frmMainSys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DevCon "
   ClientHeight    =   1485
   ClientLeft      =   2640
   ClientTop       =   3555
   ClientWidth     =   4470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer RemClip 
      Interval        =   500
      Left            =   3720
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Picture2 
      Height          =   720
      Left            =   120
      Picture         =   "frmMain2.frx":6246
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Send DevCon to the system tray?"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmMainSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHFindFiles Lib "shell32.dll" Alias "#90" (ByVal pidlRoot As Long, ByVal pidlSavedSearches As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
Picture2.Visible = False
End Sub
Private Sub Command1_Click()
MakeTopMost Me.hwnd
    AddToTray Me, "DevCon Menu", Me.Icon
    SetClipVars 4095, 8295
RemClip.Enabled = True
    SetMenuIcon Me.hwnd, 0, 2, 0, Picture2.Picture, Picture2.Picture
MsgBox "To view DevCon options just left-click the two computers icon.", vbExclamation, "Icon Added"
frmMainSys.Hide
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX

    Select Case Message
        Case 513
            temp = GetY
            If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                CustomMenu.Left = X + CustomMenu.Width ' - (CustomMenu.Width / 2)
                CustomMenu.Top = Screen.Height - CustomMenu.Height - 360 'Y '- CustomMenu.Height
                CustomMenu.Show
            End If
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RemoveClipping
    RemoveFromTray
End Sub
Private Sub RemClip_Timer()
    RemoveClipping
    RemClip.Enabled = False
End Sub



