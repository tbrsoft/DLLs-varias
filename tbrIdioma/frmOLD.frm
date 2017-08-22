VERSION 5.00
Begin VB.Form frmOLD 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      Left            =   270
      TabIndex        =   0
      Top             =   150
      Width           =   4905
   End
End
Attribute VB_Name = "frmOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    List1.Left = 60
    List1.Width = Me.Width - 220
    List1.Top = 60
    List1.Height = Me.Height - 120
End Sub

Private Sub List1_DblClick()
    frmTraductor.txt = List1
    List1.RemoveItem List1.ListIndex
End Sub
