VERSION 5.00
Begin VB.Form frmButtons 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmButtons"
   ClientHeight    =   5448
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5508
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5448
   ScaleWidth      =   5508
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   5
      Left            =   1920
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   5
      Top             =   1920
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   4
      Left            =   3840
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   4
      Top             =   1920
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   3
      Left            =   2520
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   3
      Top             =   3720
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   2
      Left            =   2640
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   1
      Left            =   240
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   1
      Top             =   3000
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   0
      Left            =   360
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CFCFDF&
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007FFF00&
      Height          =   456
      Left            =   1680
      TabIndex        =   6
      Top             =   1500
      Width           =   1212
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me, , 3
End Sub

Private Sub Form_Load()
  Dim t As Byte
  frmLoading.Show
  frmLoading.Refresh
  Me.Picture = LoadPicture(AppPath(App.Path) & "Image3.jpg")
  ShapeMe Me, RGB(255, 255, 255)
  For t = 0 To Picture1.Count - 1
    Picture1(t).Picture = LoadPicture(AppPath(App.Path) & "Image2.bmp")
    ShapeMe Picture1(t), RGB(0, 0, 0)
  Next t
  SetTrans Me, 0
  Unload frmLoading
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Picture1_Click(Index As Integer)
  If Index = 5 Then
    FadeOut Me, , 3
  Else
    MsgBox Index, vbOKOnly, "CLICK"
  End If
End Sub

Private Sub picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture1(Index).Picture = LoadPicture(AppPath(App.Path) & "Image1.bmp")
End Sub

Private Sub picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture1(Index).Picture = LoadPicture(AppPath(App.Path) & "Image2.bmp")
End Sub
