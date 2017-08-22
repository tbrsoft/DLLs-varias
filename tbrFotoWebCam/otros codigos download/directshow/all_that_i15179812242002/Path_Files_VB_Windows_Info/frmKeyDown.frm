VERSION 5.00
Begin VB.Form frmKeyDown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Press Escape Key To Continue ..."
   ClientHeight    =   828
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4392
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   828
   ScaleWidth      =   4392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press Escape Key To Continue ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3924
   End
End
Attribute VB_Name = "frmKeyDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  Do
    DoEvents
  Loop Until isKeyDown(VK_ESCAPE)
  Me.Hide
End Sub

