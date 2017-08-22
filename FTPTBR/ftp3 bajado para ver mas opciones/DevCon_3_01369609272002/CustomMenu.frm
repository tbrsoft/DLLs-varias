VERSION 5.00
Begin VB.Form CustomMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3045
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CustomMenu.frx":0000
   ScaleHeight     =   3045
   ScaleWidth      =   1980
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3.0.3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DevCon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "CustomMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Integer
Dim MyHandleOnThings As Integer
Private Sub Command1_Click()
FrmConnect.Show
End Sub
Private Sub Command2_Click()
frmAbout.Show
End Sub
Private Sub Command4_Click()
End
End Sub
Private Sub Form_Load()
Me.Show
MyHandleOnThings = GetActiveWindow
SetActiveWindow (MyHandleOnThings)
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
If Me.Height >= 3030 Then
Timer1.Enabled = False
Timer2.Enabled = True
Exit Sub
End If
Me.Top = Screen.Height - CustomMenu.Height - 650                'Y '- CustomMenu.Height

Me.Height = Me.Height + 200

End Sub
Private Sub Timer2_Timer()
    If GetActiveWindow() <> MyHandleOnThings Then
    Unload Me
    End If
End Sub

