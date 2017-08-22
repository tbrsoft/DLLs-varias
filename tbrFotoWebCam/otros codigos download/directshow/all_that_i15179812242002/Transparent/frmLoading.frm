VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   0  'None
   Caption         =   "LOADING"
   ClientHeight    =   3612
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8508
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   709
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   2988
      Left            =   24
      Top             =   24
      Width           =   7908
   End
   Begin VB.Label Label1 
      Caption         =   "LOADING"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   48
      TabIndex        =   0
      Top             =   48
      Width           =   7152
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Label1.AutoSize = True
  Me.ScaleMode = vbPixels
  Shape1.Left = Shape1.BorderWidth / 2
  Shape1.Top = Shape1.BorderWidth / 2
  Label1.Left = Label1.Left + (240 / Screen.TwipsPerPixelX)
  Shape1.Width = Label1.Width + (480 / Screen.TwipsPerPixelX) + Shape1.BorderWidth + 1
  Shape1.Height = Label1.Height + Shape1.BorderWidth + 1
  Me.Width = (Label1.Width * Screen.TwipsPerPixelX) + 480 + ((Shape1.BorderWidth * 2) * Screen.TwipsPerPixelX)
  Me.Height = (Label1.Height * Screen.TwipsPerPixelY) + ((Shape1.BorderWidth * 2) * Screen.TwipsPerPixelY)
End Sub
