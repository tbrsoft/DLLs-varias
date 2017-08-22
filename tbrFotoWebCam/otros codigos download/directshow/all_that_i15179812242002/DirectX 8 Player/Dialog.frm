VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audio Mode"
   ClientHeight    =   1944
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   3108
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1944
   ScaleWidth      =   3108
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "SHARED STEREO PLUS REVERB"
      Height          =   312
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Value           =   -1  'True
      Width           =   2892
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DYNAMIC 3D"
      Height          =   312
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DYNAMIC STEREO"
      Height          =   312
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   1812
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DYNAMIC MONO"
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1572
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   1332
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
  isOK = True
End Sub

Private Sub OKButton_Click()
  isOK = True
  Me.Hide
End Sub

Private Sub Option1_Click(Index As Integer)
  AudioMode = Index + 1
End Sub
