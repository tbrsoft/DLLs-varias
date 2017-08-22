VERSION 5.00
Begin VB.UserControl cltPROG 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ScaleHeight     =   1530
   ScaleWidth      =   4530
   Begin VB.PictureBox fbPorc 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1110
      Width           =   3015
   End
   Begin VB.Label lbINFO 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   2865
   End
End
Attribute VB_Name = "cltPROG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Porc(p As Single, tit As String)
    'fbPorc.Width = (UserControl.Width - 160) * p
    'fbPorc.Caption = CStr(Round(p * 100, 2)) + " %"
    lbINFO.Caption = tit
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    lbINFO.Top = 30
    lbINFO.Left = 30
    lbINFO.Width = UserControl.Width - 60
    lbINFO.Height = UserControl.Height - fbPorc.Height - 260
    
    fbPorc.Top = lbINFO.Top + lbINFO.Height + 30
    'fbPorc.Height = UserControl.Height - lbINFO.Top - lbINFO.Height - 160
    fbPorc.Left = 30
    fbPorc.Width = 120
End Sub

Public Sub Refresh()
    lbINFO.Refresh
    UserControl.Refresh
End Sub

