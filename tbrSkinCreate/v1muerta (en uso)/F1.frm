VERSION 5.00
Begin VB.Form F1 
   Caption         =   "Creacion de skins de tbrSoft"
   ClientHeight    =   9600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "F1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frCrear 
      Caption         =   "Crear un skin"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   30
      TabIndex        =   22
      Top             =   5310
      Visible         =   0   'False
      Width           =   11745
      Begin VB.Frame frmColorSkin 
         Caption         =   "tonos de colores"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2790
         TabIndex        =   74
         Top             =   210
         Width           =   3315
         Begin VB.CommandButton Command4 
            Caption         =   "chg"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7290
            TabIndex        =   75
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblTONOS 
            Alignment       =   2  'Center
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label lblColorSkin 
            BackColor       =   &H00800000&
            Height          =   255
            Left            =   1380
            TabIndex        =   76
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame frIMGSKIN 
         Caption         =   "imagen elegida"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4035
         Left            =   2760
         TabIndex        =   42
         Top             =   180
         Width           =   8895
         Begin VB.CommandButton Command18 
            Caption         =   "chg"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   64
            Top             =   150
            Width           =   435
         End
         Begin VB.PictureBox picCont2 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   4020
            ScaleHeight     =   3735
            ScaleWidth      =   4785
            TabIndex        =   48
            Top             =   210
            Width           =   4785
            Begin VB.Image DEimg2 
               BorderStyle     =   1  'Fixed Single
               Height          =   1155
               Left            =   60
               Top             =   60
               Width           =   1065
            End
         End
         Begin VB.TextBox T2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   840
            TabIndex        =   47
            Top             =   3660
            Width           =   945
         End
         Begin VB.TextBox T2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   2490
            TabIndex        =   46
            Top             =   3690
            Width           =   945
         End
         Begin VB.TextBox T2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   2490
            TabIndex        =   45
            Top             =   3360
            Width           =   945
         End
         Begin VB.TextBox T2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   840
            TabIndex        =   44
            Top             =   3330
            Width           =   945
         End
         Begin VB.TextBox tDesc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   1950
            Width           =   3915
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "q tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   3480
            TabIndex        =   63
            Top             =   3120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "margen der tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   6
            Left            =   90
            TabIndex        =   62
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "margen izq tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   7
            Left            =   90
            TabIndex        =   61
            Top             =   1290
            Width           =   1110
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "margen sup tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   8
            Left            =   90
            TabIndex        =   60
            Top             =   1500
            Width           =   1140
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "margen inf tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   9
            Left            =   90
            TabIndex        =   59
            Top             =   1710
            Width           =   1080
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "transparencia"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   5
            Left            =   75
            TabIndex        =   58
            Top             =   870
            Width           =   1065
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "ancho / alto"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   90
            TabIndex        =   57
            Top             =   660
            Width           =   915
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "alto minimo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   56
            Top             =   450
            Width           =   840
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   "ancho minimo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   55
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "inferior"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   24
            Left            =   1845
            TabIndex        =   54
            Top             =   3720
            Width           =   585
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "superior"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   25
            Left            =   1845
            TabIndex        =   53
            Top             =   3390
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "izquierdo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   27
            Left            =   120
            TabIndex        =   52
            Top             =   3720
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "margen transparencias"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   30
            TabIndex        =   51
            Top             =   3090
            Width           =   1965
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "derecho"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   180
            TabIndex        =   50
            Top             =   3390
            Width           =   645
         End
         Begin VB.Label NFO 
            AutoSize        =   -1  'True
            Caption         =   " izq tra"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2700
            TabIndex        =   49
            Top             =   3090
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.TextBox tNameSKIN 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "nombreSKIN"
         Top             =   240
         Width           =   2565
      End
      Begin VB.ListBox lstImagenes2 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   23
         Top             =   570
         Width           =   2625
      End
   End
   Begin VB.Frame frDEFINE 
      Caption         =   "Definir SKIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame1 
         Caption         =   "etiquetas de texto"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   7650
         TabIndex        =   69
         Top             =   2610
         Visible         =   0   'False
         Width           =   3315
         Begin VB.TextBox Text1 
            Height          =   975
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Text            =   "F1.frx":0CCA
            Top             =   330
            Width           =   2205
         End
         Begin VB.CommandButton Command3 
            Caption         =   "cambiar"
            Height          =   375
            Left            =   2430
            TabIndex        =   72
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "chg"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7290
            TabIndex        =   70
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblTextEtiq 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "texto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   180
            TabIndex        =   71
            Top             =   1350
            Width           =   3015
         End
      End
      Begin VB.Frame frTONOSSKI 
         Caption         =   "tonos de colores"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7650
         TabIndex        =   65
         Top             =   1500
         Width           =   3315
         Begin VB.CommandButton Command1 
            Caption         =   "chg"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7290
            TabIndex        =   66
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblTONOS2 
            BackColor       =   &H00800000&
            Height          =   255
            Left            =   1350
            TabIndex        =   68
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label lblTONOS 
            Alignment       =   2  'Center
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   300
            Width           =   1185
         End
      End
      Begin VB.Frame frIMGSKI 
         Caption         =   "imagen elegida"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   2760
         TabIndex        =   4
         Top             =   630
         Width           =   8955
         Begin VB.CommandButton Command12 
            Caption         =   "chg"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8460
            TabIndex        =   41
            Top             =   180
            Width           =   435
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   2310
            TabIndex        =   33
            Top             =   2640
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1350
            TabIndex        =   31
            Top             =   2640
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   2310
            TabIndex        =   30
            Top             =   3300
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1350
            TabIndex        =   29
            Top             =   3300
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   2310
            TabIndex        =   28
            Top             =   2970
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   1350
            TabIndex        =   27
            Top             =   2970
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2310
            TabIndex        =   25
            Top             =   2310
            Width           =   945
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   1350
            TabIndex        =   20
            Top             =   2310
            Width           =   945
         End
         Begin VB.PictureBox picCont 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   4200
            Left            =   3420
            ScaleHeight     =   4200
            ScaleWidth      =   5460
            TabIndex        =   19
            Top             =   270
            Width           =   5460
            Begin VB.Image DEimg 
               BorderStyle     =   1  'Fixed Single
               Height          =   1155
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   1065
            End
         End
         Begin VB.ComboBox cmbTRA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "F1.frx":0CD0
            Left            =   180
            List            =   "F1.frx":0CDD
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1500
            Width           =   1515
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   5
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   3840
            Width           =   3225
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2190
            TabIndex        =   13
            Top             =   1440
            Width           =   800
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2220
            TabIndex        =   11
            Top             =   870
            Width           =   800
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2220
            TabIndex        =   9
            Top             =   570
            Width           =   800
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   7
            Top             =   870
            Width           =   800
         End
         Begin VB.TextBox T 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   5
            Top             =   570
            Width           =   800
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "margen inferior"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   90
            TabIndex        =   40
            Top             =   3330
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "margen superior"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   39
            Top             =   3000
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "margen izq."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   38
            Top             =   2670
            Width           =   1275
         End
         Begin VB.Label lblINFO 
            Alignment       =   2  'Center
            Caption         =   "transparencia"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   2070
            Width           =   1185
         End
         Begin VB.Label lblINFO 
            Alignment       =   2  'Center
            Caption         =   "ancho minimo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   2190
            TabIndex        =   36
            Top             =   1740
            Width           =   795
         End
         Begin VB.Label lblINFO 
            Caption         =   "ancho: 0000 px"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   35
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label lblINFO 
            Caption         =   "alto: 0000 px"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   34
            Top             =   330
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "margen derecho"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   90
            TabIndex        =   32
            Top             =   2340
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "maximos"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   2340
            TabIndex        =   26
            Top             =   2070
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "minimos"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   1380
            TabIndex        =   21
            Top             =   2070
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "detalles"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   60
            TabIndex        =   18
            Top             =   3630
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "transparencia"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   16
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "ancho / alto"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   2160
            TabIndex        =   14
            Top             =   1230
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "maximo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   1560
            TabIndex        =   12
            Top             =   930
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "minimo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1620
            TabIndex        =   10
            Top             =   600
            Width           =   585
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "maximo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "minimo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   6
            Top             =   600
            Width           =   585
         End
      End
      Begin VB.ListBox lstImagenes 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         IntegralHeight  =   0   'False
         Left            =   90
         TabIndex        =   3
         Top             =   750
         Width           =   2595
      End
      Begin VB.TextBox tNamePAK 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         TabIndex        =   1
         Top             =   210
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "nombre del paquete (sin espacios)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   2985
      End
   End
   Begin VB.Menu mnSistema 
      Caption         =   "Sistema"
      Begin VB.Menu mnGotoSKI_ 
         Caption         =   "ir a Definicion de SKINS (SKI_)"
      End
      Begin VB.Menu mnGoSKIN 
         Caption         =   "ir a SKINS"
      End
      Begin VB.Menu mnQuit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnSKI_ 
      Caption         =   "SKI_"
      Begin VB.Menu mnOpenSKI_ 
         Caption         =   "Abrir SKI_"
      End
      Begin VB.Menu mnSaveSki 
         Caption         =   "Grabar"
      End
      Begin VB.Menu mnSaveAsSKI 
         Caption         =   "Grabar Como"
      End
      Begin VB.Menu mnCloseSKI 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu mnimgSKI_ 
         Caption         =   "imagenes"
         Begin VB.Menu mnVerImgSKI_ 
            Caption         =   "ver"
         End
         Begin VB.Menu mnAddimgSKI_ 
            Caption         =   "Agregar"
         End
         Begin VB.Menu mnKillImgSKI 
            Caption         =   "Quitar"
         End
         Begin VB.Menu mnCHGIMGSKI 
            Caption         =   "Cambiar actual"
         End
      End
      Begin VB.Menu mnColoresSKI_ 
         Caption         =   "colores"
         Begin VB.Menu mnVerTonoSKI 
            Caption         =   "ver"
         End
         Begin VB.Menu mnAddColor 
            Caption         =   "Agregar"
         End
         Begin VB.Menu mnQuitarColor 
            Caption         =   "Quitar"
         End
      End
      Begin VB.Menu mnCoordSKI_ 
         Caption         =   "coordenadas"
      End
   End
   Begin VB.Menu mnSKIN 
      Caption         =   "SKIN"
      Begin VB.Menu mnNewSKIN 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnAbrirSKIN 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mUpDef 
         Caption         =   "UpdateDef"
      End
      Begin VB.Menu mnSaveSKIN 
         Caption         =   "Grabar"
      End
      Begin VB.Menu msSaveAsSKIN 
         Caption         =   "Grabar Como"
      End
      Begin VB.Menu mnCloseSKIN 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu mnimgSKIN 
         Caption         =   "imagenes"
      End
      Begin VB.Menu mnColSKIN 
         Caption         =   "Colores"
      End
      Begin VB.Menu mnSaveSkinComoSki_ 
         Caption         =   "Grabar como SKI_ (probando)"
      End
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PAK As New tbrFullPak02.clsPakageSkin

Dim LastFolderIMG As String
Dim FSO2 As New Scripting.FileSystemObject
Dim NoGrabar As Boolean
Dim PathGrabar As String 'vacio si no se grabo nada aun! (este hace referencia a archivo final SKI_)
Dim PathGrabar2 As String 'vacio si no se grabo nada aun! (este hace referencia a archivo final SKIN)

Dim indexSEL As Long 'elemento elegido en las listas

Private Sub cmbTRA_Click()
    Label2(26).Visible = (cmbTRA.ListIndex <> 2)
    Label2(1).Visible = (cmbTRA.ListIndex <> 2)
    Label2(9).Visible = (cmbTRA.ListIndex <> 2)
    Label2(10).Visible = (cmbTRA.ListIndex <> 2)
    lblINFO(3).Visible = (cmbTRA.ListIndex <> 2)
    Label2(11).Visible = (cmbTRA.ListIndex <> 2)
    Label2(8).Visible = (cmbTRA.ListIndex <> 2)
    T(9).Visible = (cmbTRA.ListIndex <> 2)
    T(12).Visible = (cmbTRA.ListIndex <> 2)
    T(13).Visible = (cmbTRA.ListIndex <> 2)
    T(7).Visible = (cmbTRA.ListIndex <> 2)
    T(8).Visible = (cmbTRA.ListIndex <> 2)
    T(6).Visible = (cmbTRA.ListIndex <> 2)
    T(10).Visible = (cmbTRA.ListIndex <> 2)
    T(11).Visible = (cmbTRA.ListIndex <> 2)
End Sub

Private Sub Command12_Click()
    Dim CM As New CommonDialog
    
    CM.InitDir = LastFolderIMG 'si es "" ta ok en XP
    CM.DialogTitle = "Elija una imagen modelo para este paquete ..."
    CM.Filter = "Imagenes JPG GIF BMP TIFF|*.jpg;*.jpeg;*.gif;*.bmp;*.tiff" 'TODO ver pngs
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'joia para que se acuerde!
    LastFolderIMG = FSO2.GetBaseName(CM.FileName)
    
    PAK.getDef.ChgPathImage lstImagenes.ListIndex + 1, F
    
    ShowImage F, True
End Sub

Private Sub ListarSKIN(Optional what As String = "imagenes")
    tNameSKIN.Text = PAK.NameSKIN
    
    lstImagenes2.Clear
    
    Dim A As Long, hasta As Long
    If what = "imagenes" Then hasta = PAK.getDef.GetCantImgs
    If what = "colores" Then hasta = PAK.getDef.GetCantColores
            
    For A = 1 To hasta
        If what = "imagenes" Then lstImagenes2.AddItem PAK.getDef.GetNameImage(A)
        If what = "colores" Then lstImagenes2.AddItem PAK.getDef.getNameColor(A)
        
    Next A
    
    If lstImagenes2.ListCount > 0 Then lstImagenes2.ListIndex = 0
    
    'mostrar solo lo que corresponde!!
    frIMGSKIN.Visible = (what = "imagenes")
    frmColorSkin.Visible = (what = "colores")
    
End Sub

Private Sub Command18_Click()
    Dim CM As New CommonDialog
    
    CM.InitDir = LastFolderIMG 'si es "" ta ok en XP
    CM.DialogTitle = "Elija una imagen para este paquete ..."
    CM.Filter = "Imagenes JPG GIF BMP TIFF|*.jpg;*.jpeg;*.gif;*.bmp;*.png;*.tiff"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'joia para que se acuerde!
    LastFolderIMG = FSO2.GetBaseName(F)
    
    PAK.getDef.ChgPathImage lstImagenes2.ListIndex + 1, F
    
    ShowImage F, False
End Sub

Private Sub Save()
    If PathGrabar = "" Then
        SaveAs
    Else
        PAK.GrabarPackage PathGrabar, False
        MsgBox "Grabado ok"
    End If
End Sub

Private Sub SaveAs()
    Dim CM As New CommonDialog
    CM.DialogTitle = "Grabar definicion de skins..."
    CM.ShowSave
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    If LCase(Right(F, 5)) <> ".ski_" Then F = F + ".SKI_"
    
    PathGrabar = F
    
    PAK.GrabarPackage F, True
    
    MsgBox "Grabado ok"
End Sub

Private Sub Form_Load()
    mnSKI_.Enabled = False
    mnSKIN.Enabled = False
    chgTamano 0
    PAK.SetLogErr (App.path + "\logSkin.log")
End Sub

Private Sub chgTamano(I As Long)
    Select Case I
        Case 0 'nada elegido
            Me.Width = 3000
            Me.Height = 1550
            frDEFINE.Visible = False
            frCrear.Visible = False
        Case 1
            frDEFINE.Visible = True
            frCrear.Visible = False
            frDEFINE.Left = 60
            frDEFINE.Top = 90
            Me.Width = frDEFINE.Left + frDEFINE.Width + 200
            Me.Height = frDEFINE.Top + frDEFINE.Height + 900
        Case 2
            frCrear.Visible = True
            frDEFINE.Visible = False
            frCrear.Left = 60
            frCrear.Top = 90
            Me.Width = frCrear.Left + frCrear.Width + 200
            Me.Height = frCrear.Top + frCrear.Height + 600
    End Select
    
    CenterMe
End Sub

Private Sub lblColorSkin_Click()

    Dim CM As New CommonDialog
    CM.DialogTitle = "Elegir Color"
    CM.ShowColor
    Dim col As Long
    col = CM.RGBResult
    lblColorSkin.BackColor = col
    'asignarla a la clase que corresponda
    PAK.getDef.DefineColorValue indexSEL, col

End Sub

Private Sub lblTONOS2_Click()

    Dim CM As New CommonDialog
    CM.DialogTitle = "Elegir Color"
    CM.ShowColor
    Dim col As Long
    col = CM.RGBResult
    lblTONOS2.BackColor = col
    'asignarla a la clase que corresponda
    PAK.getDef.DefineColorValue indexSEL, col

End Sub

Private Sub lstImagenes_Click()
    NoGrabar = True
    
    Dim OB As Long
    OB = lstImagenes.ListIndex + 1
    indexSEL = OB
    
    'puedo estar en imagenes ....
    If frIMGSKI.Visible Then
        ShowImage PAK.getDef.GetpathImage(OB), True
        
        'mostrar todo lo demas!
        T(0).Text = PAK.getDef.GetMinWidth(OB)
        T(1).Text = PAK.getDef.GetMaxWidth(OB)
        T(2).Text = PAK.getDef.GetMinHeight(OB)
        T(3).Text = PAK.getDef.GetMaxHeight(OB)
        T(4).Text = PAK.getDef.GetCoef(OB)
        cmbTRA.ListIndex = PAK.getDef.GetTranspType(OB) - 1
        T(5).Text = PAK.getDef.GetTranspDescripcion(OB)
        T(6).Text = PAK.getDef.GetMaxMargenDerechoTrans(OB)
        T(7).Text = PAK.getDef.GetMinMargenSuperiorTrans(OB)
        T(8).Text = PAK.getDef.GetMaxMargenSuperiorTrans(OB)
        T(9).Text = PAK.getDef.GetMinMargenDerechoTrans(OB)
        T(10).Text = PAK.getDef.GetMinMargenInferiorTrans(OB)
        T(11).Text = PAK.getDef.GetMaxMargenInferiorTrans(OB)
        T(12).Text = PAK.getDef.GetMinMargenIzquierdoTrans(OB)
        T(13).Text = PAK.getDef.GetMaxMargenIzquierdoTrans(OB)
    End If
    
    'o en colores ...
    If frTONOSSKI.Visible Then
        lblTONOS2.BackColor = PAK.getDef.getColorById(OB)
    End If
    
    NoGrabar = False
End Sub

Private Sub ShowImage(F As String, ToSKI_ As Boolean)

    Dim IMG As Image
    Dim PIC As PictureBox
    
    If ToSKI_ Then
        Set IMG = DEimg
        Set PIC = picCont
    Else
        Set IMG = DEimg2
        Set PIC = picCont2
    End If
    
    IMG.Stretch = False 'para saber sus medidas verdaderas y poderlo en proporcion
    IMG.Picture = LoadPicture(F)
    
    If ToSKI_ Then
        lblINFO(0) = "ancho: " + CStr(CLng(IMG.Width / 15)) + " px"
        lblINFO(1) = "alto: " + CStr(CLng(IMG.Height / 15)) + " px"
        lblINFO(2) = Round(CSng(IMG.Width / IMG.Height), 2)
    Else
        NFO(1) = CLng(IMG.Width / 15)
        NFO(3) = CLng(IMG.Height / 15)
    End If
    
    Dim isMayWi As Boolean
    isMayWi = (IMG.Width > IMG.Height)
    Dim CoefX As Single
    If isMayWi Then
        CoefX = IMG.Width / (PIC.Width * 0.95)
    Else
        CoefX = IMG.Height / (PIC.Height * 0.95)
    End If
    'loacomodo
    IMG.Stretch = True
    IMG.Width = IMG.Width / CoefX
    IMG.Height = IMG.Height / CoefX
    IMG.Left = PIC.Width / 2 - IMG.Width / 2
    IMG.Top = PIC.Height / 2 - IMG.Height / 2
End Sub

Private Sub GrabarFields(ToSKI_ As Boolean)
    If NoGrabar Then Exit Sub
    
    Dim OB As Long
    
    
    If ToSKI_ Then
        OB = lstImagenes.ListIndex + 1
        If IsNumeric(T(0).Text) Then PAK.getDef.DefineMinWidth OB, CLng(T(0).Text) 'es el minimo ancho
        If IsNumeric(T(1).Text) Then PAK.getDef.DefineMaxWidth OB, CLng(T(1).Text) 'es el maximo ancho
        If IsNumeric(T(2).Text) Then PAK.getDef.DefineMinHeight OB, CLng(T(2).Text) 'es el minimo alto
        If IsNumeric(T(3).Text) Then PAK.getDef.DefineMaxHeight OB, CLng(T(3).Text) 'es el maximo alto
        If IsNumeric(T(4).Text) Then PAK.getDef.DefineCoef OB, CSng(T(4).Text) 'ancho / alto
        If cmbTRA.Text = "" Then
            PAK.getDef.DefineTranspType OB, Opcional
        Else
            PAK.getDef.DefineTranspType OB, CLng(Left(cmbTRA.Text, 1))
        End If
        PAK.getDef.DefineTranspDescripcion OB, T(5).Text
        If IsNumeric(T(6).Text) Then PAK.getDef.DefineMaxMargenDerechoTrans OB, CSng(T(6).Text) 'maximo margen derecho
        If IsNumeric(T(7).Text) Then PAK.getDef.DefineMinMargenSuperiorTrans OB, CSng(T(7).Text) 'minimo mar sup
        If IsNumeric(T(8).Text) Then PAK.getDef.DefineMaxMargenSuperiorTrans OB, CSng(T(8).Text) 'maximo margen sup
        If IsNumeric(T(9).Text) Then PAK.getDef.DefineMinMargenDerechoTrans OB, CSng(T(9).Text) 'min margen derecho
        If IsNumeric(T(10).Text) Then PAK.getDef.DefineMinMargenInferiorTrans OB, CSng(T(10).Text) 'min mar inf
        If IsNumeric(T(11).Text) Then PAK.getDef.DefineMaxMargenInferiorTrans OB, CSng(T(11).Text) 'max mar inf
        If IsNumeric(T(12).Text) Then PAK.getDef.DefineMinMargenIzquierdoTrans OB, CSng(T(12).Text) 'min mar izq
        If IsNumeric(T(13).Text) Then PAK.getDef.DefineMaxMargenIzquierdoTrans OB, CSng(T(13).Text) 'max mar izq
    Else
        OB = lstImagenes2.ListIndex + 1
        If IsNumeric(T2(17).Text) Then PAK.getDef.DefineFinalMargenDerechoTra OB, CLng(T2(17).Text)
        If IsNumeric(T2(14).Text) Then PAK.getDef.DefineFinalMargenIzquierdoTra OB, CLng(T2(14).Text)
        If IsNumeric(T2(16).Text) Then PAK.getDef.DefineFinalMargenSuperiorTra OB, CLng(T2(16).Text)
        If IsNumeric(T2(15).Text) Then PAK.getDef.DefineFinalMargenInferiorTra OB, CLng(T2(15).Text)
    End If
End Sub

Private Sub ClearFields()
    NoGrabar = True
    Dim J As Long
    For J = 0 To 13
        T(J).Text = ""
    Next J
    cmbTRA.ListIndex = 1
    NoGrabar = False
End Sub

Private Sub lstImagenes2_Click()
    NoGrabar = True
    
    Dim OB As Long
    OB = lstImagenes2.ListIndex + 1
    indexSEL = OB
    
    'puedo estar en imagenes ...
    If frIMGSKIN.Visible Then
        ShowImage PAK.getDef.GetpathImage(OB), False
        
        'mostrar todo lo demas!
        NFO(0).Caption = "ancho (" + NFO(1) + ") minimo:" + CStr(PAK.getDef.GetMinWidth(OB)) + " maximo: " + CStr(PAK.getDef.GetMaxWidth(OB))
        NFO(2).Caption = "alto (" + NFO(3) + ") maximo minimo:" + CStr(PAK.getDef.GetMinHeight(OB)) + " maximo: " + CStr(PAK.getDef.GetMaxHeight(OB))
        NFO(4).Caption = "ancho / alto: " + CStr(PAK.getDef.GetCoef(OB))
        
        Select Case PAK.getDef.GetTranspType(OB)
            Case 1: NFO(5).Caption = "transparencia: 1-Obligatoria"
            Case 2: NFO(5).Caption = "transparencia: 2-Opcional"
            Case 3: NFO(5).Caption = "transparencia: 3-Prohibida"
        End Select
        
        NFO(6).Caption = "margen derecho minimo:" + CStr(PAK.getDef.GetMinMargenDerechoTrans(OB)) + _
            " maximo: " + CStr(PAK.getDef.GetMaxMargenDerechoTrans(OB))
        
        NFO(7).Caption = "margen izquierdo minimo:" + CStr(PAK.getDef.GetMinMargenIzquierdoTrans(OB)) + _
            " maximo: " + CStr(PAK.getDef.GetMaxMargenIzquierdoTrans(OB))
            
        NFO(8).Caption = "margen superior minimo:" + CStr(PAK.getDef.GetMinMargenSuperiorTrans(OB)) + _
            " maximo: " + CStr(PAK.getDef.GetMaxMargenSuperiorTrans(OB))
            
        NFO(9).Caption = "margen inferior minimo:" + CStr(PAK.getDef.GetMinMargenInferiorTrans(OB)) + _
            " maximo: " + CStr(PAK.getDef.GetMaxMargenInferiorTrans(OB))
        
        tDesc.Text = "Observaciones:" + PAK.getDef.GetTranspDescripcion(OB)
        
        T2(17).Text = PAK.getDef.GetFinalMargenDerechoTra(OB)
        T2(14).Text = PAK.getDef.GetFinalMargenIzquierdoTra(OB)
        T2(16).Text = PAK.getDef.GetFinalMargenSuperiorTra(OB)
        T2(15).Text = PAK.getDef.GetFinalMargenInferiorTra(OB)
    End If
    
    ' o en colores ...
    If frmColorSkin.Visible Then
        lblColorSkin.BackColor = PAK.getDef.getColorById(OB)
    End If
    
    NoGrabar = False
End Sub

Private Sub mnAbrirSKIN_Click()
    Dim CM As New CommonDialog
    CM.DialogTitle = "Abrir SKIN"
    CM.Filter = "SKINs (*.SKIN)|*.SKIN"
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    PathGrabar2 = F ' FSO2.GetParentFolderName(F) + "\" + FSO2.GetBaseName(F) + "\"
    
    PAK.AbrirSKIN F
    
    'ahora debo cargarlo en pantalla
    ListarSKIN "imagenes"
End Sub

Private Sub mnAddColor_Click()
    Dim nI As String
    nI = InputBox("Defina el nombre con que se llamara al color")
    If nI = "" Then Exit Sub
    
    'ver que no exista el nombre!
    If lstImagenes.ListCount > 0 Then
        Dim J As Long
        For J = 0 To lstImagenes.ListCount - 1
            If nI = lstImagenes.List(J) Then
                MsgBox "Ya existe ese nombre de imagen definido!"
                Exit Sub
            End If
        Next J
    End If
    
    Dim CM As New CommonDialog
    CM.DialogTitle = "Elegir Color"
    CM.ShowColor
    
    Dim col As Long
    col = CM.RGBResult
    
    PAK.getDef.AddColor col, nI
    
    lstImagenes.AddItem nI
    
    lstImagenes.ListIndex = lstImagenes.ListCount - 1
End Sub

Private Sub mnAddimgSKI__Click()
    Dim CM As New CommonDialog
    
    CM.InitDir = LastFolderIMG 'si es "" ta ok en XP
    CM.DialogTitle = "Elija una imagen modelo para este paquete ..."
    CM.Filter = "Imagenes JPG GIF BMP TIFF|*.jpg;*.jpeg;*.gif;*.bmp;*.png;*.tiff"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'joia para que se acuerde!
    LastFolderIMG = FSO2.GetBaseName(CM.FileName)
    
    Dim nI As String
defName:
    nI = InputBox("Defina el nombre con que se llamara a la imagen")
    If nI = "" Then Exit Sub
    
    'ver que no exista el nombre!
    If lstImagenes.ListCount > 0 Then
        Dim J As Long
        For J = 0 To lstImagenes.ListCount - 1
            If nI = lstImagenes.List(J) Then
                MsgBox "Ya existe ese nombre de imagen definido!"
                GoTo defName
            End If
        Next J
    End If
    ClearFields
    
    PAK.getDef.AddImage F
    
    lstImagenes.AddItem nI
    PAK.getDef.DefineNameImage lstImagenes.ListCount, nI
    
    ShowImage F, True
    
    lstImagenes.ListIndex = lstImagenes.ListCount - 1
End Sub

Private Sub mnCHGIMGSKI_Click()
    Dim nI As String
defName:
    nI = InputBox("Defina el NUEVO nombre con que se llamara a la imagen. " + vbCrLf + _
        "El actual es " + lstImagenes.List(lstImagenes.ListIndex) + vbCrLf + _
        "Deje en blanco para no cambiar", "Cambia id de imagen", lstImagenes.List(lstImagenes.ListIndex))
    If nI = "" Then Exit Sub
    
    'ver que no exista el nombre!
    If lstImagenes.ListCount > 0 Then
        Dim J As Long
        For J = 0 To lstImagenes.ListCount - 1
            If nI = lstImagenes.List(J) Then
                MsgBox "Ya existe ese nombre de imagen definido!"
                GoTo defName
            End If
        Next J
    End If
    
    PAK.getDef.DefineNameImage lstImagenes.ListIndex + 1, nI
    lstImagenes.List(lstImagenes.ListIndex) = nI
    

End Sub

Private Sub mnCloseSKI_Click()
    lstImagenes.Clear
    PAK.getDef.Clean
    chgTamano 0
End Sub

Private Sub mnCloseSKIN_Click()
    lstImagenes2.Clear
    PAK.getDef.Clean
    chgTamano 0
End Sub

Private Sub mnColSKIN_Click()
    ListarSKIN "colores"
End Sub

Private Sub mnGoSKIN_Click()
    chgTamano 2
    mnSKIN.Enabled = True
End Sub

Private Sub mnGotoSKI__Click()
    chgTamano 1
    ShowFR frIMGSKI
    mnSKI_.Enabled = True
End Sub

Private Sub mnimgSKIN_Click()
    ListarSKIN "imagenes"
End Sub

Private Sub mnKillImgSKI_Click()
    If lstImagenes.ListIndex = -1 Then Exit Sub
    PAK.getDef.RemoveImage lstImagenes.ListIndex + 1
    lstImagenes.RemoveItem lstImagenes.ListIndex
End Sub

Private Sub mnNewSKIN_Click()
    'si o si tiene que abrir un pquete donde esta definido el skin
    'ademas tomas sus imagenes originales para cambiarlas y poder probarlas antes de terminar
    
    Dim CM As New CommonDialog
    CM.DialogTitle = "Indique la definicion de SKIN (*.SKI_) a usar"
    CM.Filter = "Definiciones de SKIN (*.SKI_)|*.SKI_"
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    CM.DialogTitle = "Ahora indique con que nombre desea grabarlo"
    CM.Filter = "Archivos SKIN|*.SKIN"
    CM.ShowSave
    
    Dim F2 As String
    F2 = CM.FileName
    
    If F2 = "" Then Exit Sub
    
    PAK.GrabarSKINFromSKI_ F, F2, True
    PathGrabar2 = F2 'FSO2.GetParentFolderName(F2) + "\" + FSO2.GetBaseName(F2) + "\"
    
    'ahora debo cargarlo en pantalla
    ListarSKIN
End Sub

Private Sub mnOpenSKI__Click()
    Dim CM As New CommonDialog
    CM.DialogTitle = "Abrir definicion de SKIN"
    CM.Filter = "Definicion de SKIN (*.SKI_)|*.ski_"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    PAK.AbrirSKI_ F
    PathGrabar = F 'FSO2.GetParentFolderName(F) + "\" + FSO2.GetBaseName(F) + "\"
    'para que no pregunte al grabar
    
    tNamePAK.Text = PAK.NamePackage
    Dim A As Long
    lstImagenes.Clear
    For A = 1 To PAK.getDef.GetCantImgs
        lstImagenes.AddItem PAK.getDef.GetNameImage(A)
    Next A
    
    lstImagenes.ListIndex = 0
    
End Sub

Private Sub mnQUIT_Click()
    Unload Me
End Sub

Private Sub mnQuitarColor_Click()
    If lstImagenes.ListIndex = -1 Then Exit Sub
    PAK.getDef.RemoveColor lstImagenes.ListIndex + 1
    lstImagenes.RemoveItem lstImagenes.ListIndex
End Sub

Private Sub mnSaveAsSKI_Click()
    SaveAs
End Sub

Private Sub mnSaveSKI_Click()
    Save
End Sub

Private Sub mnSaveSKIN_Click()
    If PathGrabar2 = "" Then
        msSaveAsSKIN_Click 'guardar como
    Else
        PAK.GrabarSKIN PathGrabar2, False
    End If
End Sub

Private Sub mnSaveSkinComoSki__Click()
    If PathGrabar = "" Then
        SaveAs
    Else
        PAK.GrabarPackage PathGrabar, False
        MsgBox "Grabado ok"
    End If
End Sub

Private Sub mnVerImgSKI__Click()
    ShowFR frIMGSKI
    
    Dim A As Long
    lstImagenes.Clear
    For A = 1 To PAK.getDef.GetCantImgs
        lstImagenes.AddItem PAK.getDef.GetNameImage(A)
    Next A
End Sub

Private Sub ShowFR(FR As Frame)
    Select Case LCase(FR.Name)
        Case "frimgski", "frtonosski" 'esta en ski_
            frIMGSKI.Visible = False
            frTONOSSKI.Visible = False
            
            FR.Visible = True
            FR.Left = lstImagenes.Left + lstImagenes.Width + 60
            FR.Top = tNamePAK.Top + tNamePAK.Height
    End Select
End Sub

Private Sub mnVerTonoSKI_Click()
    ShowFR frTONOSSKI
    
    Dim A As Long
    lstImagenes.Clear
    For A = 1 To PAK.getDef.GetCantColores
        lstImagenes.AddItem PAK.getDef.getNameColor(A)
    Next A
End Sub

Private Sub msSaveAsSKIN_Click()
    Dim CM As New CommonDialog
    CM.DialogTitle = "Grabar SKIN"
    CM.Filter = "SKINs (*.SKIN)|*.SKIN"
    
    CM.ShowSave
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    If LCase(Right(F, 5)) <> ".skin" Then F = F + ".SKIN"
    PAK.GrabarSKIN F, True
    PathGrabar2 = F 'FSO2.GetParentFolderName(F) + "\" + FSO2.GetBaseName(F) + "\"
    
    MsgBox "Grabado OK"
End Sub

Private Sub mUpDef_Click()
    'Cuado el .SKI_ cambia todos los .SKIN quedan regalados
    If PathGrabar2 = "" Then
        MsgBox "No tiene SKIN abierto"
        Exit Sub
    End If
    
    Dim CM As New CommonDialog
    CM.DialogTitle = "Definir SKI_para actualizar"
    CM.Filter = "Definicion de SKIN (*.SKI_)|*.SKI_"
    
    CM.ShowOpen
        
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    Dim EncuentraCambios As String 'para saber si graba o no
    EncuentraCambios = ""
    
    'abro aparte el SKI
    Dim sk As New tbrFullPak02.clsPakageSkin
    Dim df As New tbrFullPak02.clsDef
    sk.AbrirSKI_ F
    Set df = sk.getDef
    
    Dim dfViejo As tbrFullPak02.clsDef
    Set dfViejo = PAK.getDef
    'ahora los puedo comparar
    'primero me intersa ver que todos las imagees del ski_ este en el skin
    Dim H As Long
    For H = 1 To df.GetCantImgs
        Dim sNam As String
        sNam = df.GetNameImage(H)
        'buscarlo en el viejo
        Dim J As Long, Existe As Boolean
        Existe = False
        For J = 1 To dfViejo.GetCantImgs
            Dim sNam2 As String
            sNam2 = dfViejo.GetNameImage(J)
            If LCase(sNam) = LCase(sNam2) Then
                Existe = True
            End If
        Next J
        'ver si no esta y agregarlo
        If Existe = False Then
            EncuentraCambios = EncuentraCambios + "Nueva imagen: " + sNam + vbCrLf
            dfViejo.AddImageFull df.getDefImage(H) 'agrego enterita la imagen
        End If
    Next H
    
    For H = 1 To df.GetCantColores
        
        sNam = df.getNameColor(H)
        'buscarlo en el viejo
        
        Existe = False
        For J = 1 To dfViejo.GetCantColores
            sNam2 = dfViejo.getNameColor(J)
            If LCase(sNam) = LCase(sNam2) Then
                Existe = True
            End If
        Next J
        'ver si no esta y agregarlo
        If Existe = False Then
            EncuentraCambios = EncuentraCambios + "Nueva imagen: " + sNam + vbCrLf
            dfViejo.AddColor df.getColorById(H), df.getNameColor(H)    'agrego enterita la imagen
        End If
    Next H
    
    'grabo los cambios si hay
    If EncuentraCambios <> "" Then
        ListarSKIN 'los lista solamente, si quiere grabar lo hace despues
        'PAK.GrabarSKIN PathGrabar2, False
        MsgBox "Encontro cambios:" + vbCrLf + EncuentraCambios
    Else
        MsgBox "No hay cambios!"
    End If
    
    
    
End Sub

Private Sub T_Change(Index As Integer)
    GrabarFields True
End Sub

Private Sub T2_Change(Index As Integer)
    GrabarFields False
End Sub

Private Sub tNamePAK_Change()
    PAK.NamePackage = Replace(tNamePAK, " ", "")
End Sub

Private Sub CenterMe()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub
