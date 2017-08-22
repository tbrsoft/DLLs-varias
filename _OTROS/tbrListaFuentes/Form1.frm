VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   360
      TabIndex        =   0
      Top             =   450
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5820
      TabIndex        =   1
      Top             =   1170
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Module1.FillListWithFonts List1
End Sub


Private Sub List1_Click()
    Label1.Font.Name = List1
End Sub
