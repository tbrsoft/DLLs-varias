VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   1632
   ClientLeft      =   2304
   ClientTop       =   1512
   ClientWidth     =   2928
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1632
   ScaleWidth      =   2928
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   576
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   576
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   216
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   564
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me, , 0.5
End Sub

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Const DI_MASK = &H1
  Const DI_IMAGE = &H2
  Const DI_NORMAL = DI_MASK Or DI_IMAGE
  Dim t As Long
  Label1.Caption = GetAbout(App)
  Me.ScaleMode = vbTwips
  Image1.Left = 240
  Label1.Top = 240
  Label1.Left = (48 * Screen.TwipsPerPixelX) + 480
  Me.Height = Label1.Height + 480
  Me.Width = Label1.Left + Label1.Width + 240
  Me.ScaleMode = vbPixels
  For t = 0 To Me.ScaleHeight
    Me.ForeColor = RGB(0, 0, 255 - ((t * 255) / Me.ScaleHeight))
    DrawLine Me, 0, t, Me.ScaleWidth, t
  Next t
  Image1.Top = (Me.ScaleHeight - 48) / 2
  Image1.Picture = LoadResPicture(101, vbResIcon)
  SetTrans Me, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FadeOut Me, , 0.5
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub
