VERSION 5.00
Begin VB.Form frmMask 
   BorderStyle     =   0  'None
   Caption         =   "frmMask"
   ClientHeight    =   6552
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6552
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   492
      Index           =   1
      Left            =   6240
      TabIndex        =   2
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   612
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1212
      Left            =   120
      ScaleHeight     =   1212
      ScaleWidth      =   1212
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   1212
   End
End
Attribute VB_Name = "frmMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me, , 4
End Sub

Private Sub Command1_Click(Index As Integer)
  FadeOut Me, , 4
End Sub

Private Sub Form_Load()
  frmLoading.Show
  frmLoading.Refresh
  Me.Picture = LoadPicture(AppPath(App.Path) & "shape.jpg")
  Picture1.Picture = LoadPicture(AppPath(App.Path) & "mask.jpg")
  Picture1.AutoSize = True
  ChangeMask Me, Picture1
  SetTrans Me, 0
  Unload frmLoading
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub
