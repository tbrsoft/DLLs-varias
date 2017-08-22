VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gold Button v1.2 by Night Wolf"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "Text Align && Font"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   22
      Top             =   2040
      Width           =   1815
      Begin pGoldButton.GoldButton GoldButton1 
         Height          =   345
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Left"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pGoldButton.GoldButton GoldButton2 
         Height          =   345
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Center"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pGoldButton.GoldButton GoldButton3 
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Right"
         Alignment       =   1
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
      Begin pGoldButton.GoldButton GoldButton4 
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Font"
         Alignment       =   2
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnHover         =   5
      End
      Begin pGoldButton.GoldButton GoldButton13 
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Hover"
         Alignment       =   2
         HoverColor      =   16711935
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnHover         =   5
      End
      Begin pGoldButton.GoldButton GoldButton15 
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         Caption         =   "Down"
         Alignment       =   2
         HoverColor      =   16711935
         ForeColor       =   -2147483630
         DownColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnHover         =   5
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Styles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.Frame Frame6 
         Caption         =   "On Button Down"
         Height          =   1455
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         Begin pGoldButton.GoldButton GoldButton11 
            Height          =   345
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Default"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin pGoldButton.GoldButton GoldButton12 
            Height          =   345
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Soft"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnDown          =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "On Button Hover"
         Height          =   1455
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   1815
         Begin pGoldButton.GoldButton GoldButton8 
            Height          =   345
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Soft"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin pGoldButton.GoldButton GoldButton9 
            Height          =   345
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "None"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnHover         =   0
         End
         Begin pGoldButton.GoldButton GoldButton10 
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Default"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnHover         =   5
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "On Button Up"
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         Begin pGoldButton.GoldButton GoldButton6 
            Height          =   345
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Soft"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin pGoldButton.GoldButton GoldButton5 
            Height          =   345
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "None"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnUp            =   0
         End
         Begin pGoldButton.GoldButton GoldButton7 
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Default"
            Alignment       =   2
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnUp            =   5
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Skin && Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00937A12&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   1785
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         Begin pGoldButton.GoldButton GoldButton17 
            Height          =   340
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Blue Button"
            Alignment       =   1
            HoverColor      =   16777215
            ForeColor       =   16777215
            SkinDisabledText=   12229376
            SkinHighlight   =   9665042
            PictureBackColor=   9665042
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnUp            =   0
            OnDown          =   2
            Style           =   1
            SkinPicture     =   "frmMain.frx":000C
            Picture         =   "frmMain.frx":0AD0
            PictureHover    =   "frmMain.frx":0C6C
            MaskColor       =   16711935
            UseMaskColor    =   -1  'True
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000E8&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   1785
         TabIndex        =   2
         Top             =   960
         Width           =   1815
         Begin pGoldButton.GoldButton GoldButton16 
            Height          =   340
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            Caption         =   "Red Button"
            HoverColor      =   16777215
            ForeColor       =   16777215
            SkinDisabledText=   12229376
            SkinHighlight   =   9665042
            PictureBackColor=   232
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnHover         =   5
            Style           =   1
            SkinPicture     =   "frmMain.frx":0E08
            Picture         =   "frmMain.frx":1860
            PictureHover    =   "frmMain.frx":19FC
            MaskColor       =   16711935
            UseMaskColor    =   -1  'True
         End
      End
      Begin pGoldButton.GoldButton GoldButton18 
         Height          =   345
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1560
         _ExtentX        =   635
         _ExtentY        =   609
         Caption         =   "Picture"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "frmMain.frx":1B9C
         MaskColor       =   12632256
         UseMaskColor    =   -1  'True
      End
   End
   Begin pGoldButton.GoldButton GoldButton14 
      Height          =   345
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   "Exit"
      Alignment       =   2
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GoldButton14_Click()
Unload Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub GoldButton3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox "Cool"
End Sub
