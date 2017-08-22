VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Game # 1, Player # 1"
   ClientHeight    =   6084
   ClientLeft      =   2412
   ClientTop       =   1584
   ClientWidth     =   6264
   ClipControls    =   0   'False
   Icon            =   "ttt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6084
   ScaleWidth      =   6264
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1884
      ItemData        =   "ttt.frx":0442
      Left            =   0
      List            =   "ttt.frx":0444
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1692
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   2
      Height          =   252
      Left            =   4440
      Top             =   5760
      Width           =   612
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   252
      Left            =   2160
      Top             =   5760
      Width           =   612
   End
   Begin VB.Label sc2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4560
      TabIndex        =   13
      Top             =   5760
      Width           =   372
   End
   Begin VB.Label sc1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   1
      Top             =   5760
      Width           =   372
   End
   Begin VB.Label Label2 
      Caption         =   "Player #2:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   12
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Player #1:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   11
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   8
      Left            =   4080
      TabIndex        =   10
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   7
      Left            =   2280
      TabIndex        =   9
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   4
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   1692
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5532
      Left            =   360
      Shape           =   1  'Square
      Top             =   120
      Width           =   5532
   End
   Begin VB.Menu mexit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnombre 
      Caption         =   "Number Of Players"
      Begin VB.Menu mnum 
         Caption         =   "&0"
         Index           =   0
      End
      Begin VB.Menu mnum 
         Caption         =   "&1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnum 
         Caption         =   "&2"
         Index           =   2
      End
   End
   Begin VB.Menu mdiff 
      Caption         =   "Difficulty"
      Begin VB.Menu mdif 
         Caption         =   "Very Easy"
         Index           =   0
      End
      Begin VB.Menu mdif 
         Caption         =   "Easy"
         Index           =   1
      End
      Begin VB.Menu mdif 
         Caption         =   "Normal"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mdif 
         Caption         =   "Expert"
         Index           =   3
      End
      Begin VB.Menu mdif 
         Caption         =   "&Impossible"
         Index           =   4
      End
   End
   Begin VB.Menu mcancel 
      Caption         =   "Cancel Game"
   End
   Begin VB.Menu mreset 
      Caption         =   "&Reset Scores"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Label_Click(Index As Integer)
  If c < 9 And d = 0 And Me.Label(Index).Caption = "" And f <> 0 Then
    Form1.List1.AddItem Index
    If a = 1 Then
      Me.Label(Index).Caption = "X"
    ElseIf a = -1 Then
      Me.Label(Index).Caption = "O"
    Else
      Exit Sub
    End If
    sndPlaySound snd, 1
    c = c + 1
    Module1.check
  End If
End Sub

Private Sub mcancel_Click()
  g = g - 1
  Module1.restart
End Sub

Private Sub mdif_Click(Index As Integer)
  Dim mdift As Integer
  h = Index
  For mdift = 0 To Me.mdif.Count - 1
    If mdift = Index Then
      Me.mdif(mdift).Checked = True
    Else
      Me.mdif(mdift).Checked = False
    End If
  Next mdift
End Sub

Private Sub mexit_Click()
  End
End Sub

Private Sub mnum_Click(Index As Integer)
  Dim mnumt As Integer
  f = Index
  For mnumt = 0 To Me.mnum.Count - 1
    If mnumt = Index Then
      Me.mnum(mnumt).Checked = True
    Else
      Me.mnum(mnumt).Checked = False
    End If
  Next mnumt
  If Index = 0 Then
    mdif_Click (4)
    mcancel_Click
    Me.mdiff.Visible = False
  Else
    Me.mdiff.Visible = True
  End If
End Sub

Private Sub mreset_Click()
  d1 = 0
  d2 = 0
  sc1.Caption = 0
  sc2.Caption = 0
  g = 1
  If a = 1 Then
    Form1.Caption = "Game #" & Str$(g) & ", Player # 1"
  Else
    Form1.Caption = "Game #" & Str$(g) & ", Player # 2"
  End If
End Sub

Private Sub Timer1_Timer()
  If c < 9 And d = 0 Then
    If a = -1 Then
      If (f = 0 Or f = 1) Then
        o = "O"
        x = "X"
        Module1.play
      End If
    ElseIf a = 1 Then
      If f = 0 Then
        o = "X"
        x = "O"
        Module1.play
      End If
    End If
  End If
End Sub
