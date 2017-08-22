VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Image Image1 
         Height          =   1830
         Left            =   2520
         Picture         =   "frmAbout.frx":0000
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YZY FTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Verze: 1.0070254 C GŠ 25 Vytvoøil: D. Šmejkal a pan Bù Copyrajt: © 2001, vèechna práva vyhlazena!"
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "davidsmejkal@hellada.cz"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
