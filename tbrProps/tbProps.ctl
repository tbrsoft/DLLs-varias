VERSION 5.00
Begin VB.UserControl tbrProps 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2475
   ScaleWidth      =   10260
   Begin VB.PictureBox pProps 
      BackColor       =   &H00808080&
      Height          =   2085
      Left            =   150
      ScaleHeight     =   2025
      ScaleWidth      =   9855
      TabIndex        =   9
      Top             =   150
      Width           =   9915
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   1950
         Left            =   15
         ScaleHeight     =   1890
         ScaleWidth      =   2100
         TabIndex        =   30
         Top             =   30
         Width           =   2160
         Begin VB.TextBox txAlphaC 
            Height          =   390
            Left            =   60
            TabIndex        =   7
            Text            =   "0"
            Top             =   720
            Width           =   1110
         End
         Begin VB.CheckBox chAlpha 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Alpha Habilitado"
            Height          =   270
            Left            =   135
            TabIndex        =   6
            Top             =   60
            Width           =   1935
         End
         Begin VB.Line Line2 
            X1              =   -30
            X2              =   2130
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            X1              =   -60
            X2              =   2100
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha Cantidad (0-254)"
            Height          =   285
            Left            =   45
            TabIndex        =   33
            Top             =   420
            Width           =   2100
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "AlphaColor"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label lblAlphaC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            TabIndex        =   31
            Top             =   1560
            Width           =   810
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   8700
         ScaleHeight     =   585
         ScaleWidth      =   915
         TabIndex        =   27
         Top             =   60
         Width           =   975
         Begin VB.Label lblColorSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "ColorSel"
            Height          =   285
            Left            =   45
            TabIndex        =   28
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   1620
         Left            =   2250
         ScaleHeight     =   1560
         ScaleWidth      =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   4380
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00E0E0E0&
            Height          =   630
            Left            =   2730
            ScaleHeight     =   570
            ScaleWidth      =   1485
            TabIndex        =   22
            Top             =   255
            Width           =   1545
            Begin VB.Label lblColorUnSel 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   30
               TabIndex        =   24
               Top             =   270
               Width           =   1410
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "FontColorUnSel "
               Height          =   285
               Left            =   75
               TabIndex        =   23
               Top             =   15
               Width           =   1365
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00E0E0E0&
            Height          =   585
            Left            =   1560
            ScaleHeight     =   525
            ScaleWidth      =   1755
            TabIndex        =   19
            Top             =   930
            Width           =   1815
            Begin VB.Label lblFontColCont 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   45
               TabIndex        =   21
               Top             =   210
               Width           =   1680
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "FontColorContorno"
               Height          =   285
               Left            =   60
               TabIndex        =   20
               Top             =   -30
               Width           =   1785
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00E0E0E0&
            Height          =   630
            Left            =   1560
            ScaleHeight     =   570
            ScaleWidth      =   1005
            TabIndex        =   16
            Top             =   270
            Width           =   1065
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "FontColor "
               Height          =   285
               Left            =   75
               TabIndex        =   18
               Top             =   15
               Width           =   855
            End
            Begin VB.Label lblFontCol 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   30
               TabIndex        =   17
               Top             =   270
               Width           =   960
            End
         End
         Begin VB.TextBox txtTama 
            Height          =   390
            Left            =   15
            TabIndex        =   5
            Text            =   "0"
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtFuente 
            Height          =   390
            Left            =   30
            TabIndex        =   4
            Text            =   "Arial"
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label lblx 
            BackStyle       =   0  'Transparent
            Caption         =   "FontSize "
            Height          =   285
            Left            =   15
            TabIndex        =   26
            Top             =   825
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "FontName"
            Height          =   285
            Left            =   30
            TabIndex        =   25
            Top             =   30
            Width           =   1155
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aplicar Cambios"
         Height          =   405
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1590
         Width           =   2940
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00E0E0E0&
         Height          =   780
         Left            =   6750
         ScaleHeight     =   720
         ScaleWidth      =   2865
         TabIndex        =   10
         Top             =   765
         Width           =   2925
         Begin VB.TextBox txtX 
            Height          =   390
            Left            =   75
            TabIndex        =   0
            Text            =   "0"
            Top             =   255
            Width           =   660
         End
         Begin VB.TextBox txtY 
            Height          =   390
            Left            =   780
            TabIndex        =   1
            Text            =   "0"
            Top             =   255
            Width           =   660
         End
         Begin VB.TextBox txtAncho 
            Height          =   390
            Left            =   1470
            TabIndex        =   2
            Text            =   "0"
            Top             =   270
            Width           =   660
         End
         Begin VB.TextBox txtAlto 
            Height          =   390
            Left            =   2160
            TabIndex        =   3
            Text            =   "0"
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   15
            Width           =   450
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "y"
            Height          =   285
            Left            =   825
            TabIndex        =   13
            Top             =   15
            Width           =   465
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho"
            Height          =   285
            Left            =   1485
            TabIndex        =   12
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Alto"
            Height          =   285
            Left            =   2175
            TabIndex        =   11
            Top             =   0
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "tbrProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim CD As New CommonDialog
Public Event Cambios()

Private Sub Command1_Click()
    RaiseEvent Cambios
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < pProps.Width Then UserControl.Width = pProps.Width
    If UserControl.Height < pProps.Height Then UserControl.Height = pProps.Height
    pProps.Left = (UserControl.Width / 2) - (pProps.Width / 2)
    pProps.Top = (UserControl.Height / 2) - (pProps.Height / 2)
End Sub

'=========================================================================0
'X,Y,Alto,Ancho
'=========================================================================0
Public Property Get X() As Long
    X = Val(txtX.Text)
End Property

Public Property Let X(ByVal vNewValue As Long)
    txtX.Text = CStr(vNewValue)
End Property

Public Property Get Y() As Long
    Y = Val(txtY.Text)
End Property

Public Property Let Y(ByVal vNewValue As Long)
    txtY.Text = CStr(vNewValue)
End Property

Public Property Get Ancho() As Long
    Ancho = Val(txtAncho.Text)
End Property

Public Property Let Ancho(ByVal vNewValue As Long)
    txtAncho.Text = CStr(vNewValue)
End Property

Public Property Get Alto() As Long
    Alto = Val(txtAlto.Text)
End Property

Public Property Let Alto(ByVal vNewValue As Long)
    txtAlto.Text = CStr(vNewValue)
End Property
'=========================================================================0
'Fuente
'=========================================================================0
Public Property Get Font() As String
    Font = txtFuente.Text
End Property

Public Property Let Font(ByVal vNewValue As String)
    txtFuente.Text = vNewValue
End Property
 
Public Property Get FontSize() As Long
    FontSize = txtTama.Text
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
    txtTama.Text = vNewValue
End Property
 
Public Property Get FontColor() As Long
    FontColor = lblFontCol.BackColor
End Property

Public Property Let FontColor(ByVal vNewValue As Long)
    lblFontCol.BackColor = vNewValue
End Property

Public Property Get FontColorUnSel() As Long
    FontColorUnSel = lblColorUnSel.BackColor
End Property

Public Property Let FontColorUnSel(ByVal vNewValue As Long)
    lblColorUnSel.BackColor = vNewValue
End Property

Public Property Get FontColorContorno() As Long
    FontColorContorno = lblFontColCont.BackColor
End Property

Public Property Let FontColorContorno(ByVal vNewValue As Long)
    lblFontColCont.BackColor = vNewValue
End Property

'=========================================================================0
'Alpha
'=========================================================================0
Public Property Get AlphaHabilitado() As Boolean
    AlphaHabilitado = chAlpha.Value
End Property

Public Property Let AlphaHabilitado(ByVal vNewValue As Boolean)
    chAlpha.Value = Val(vNewValue)
End Property

Public Property Get AlphaCantidad() As Long
    AlphaCantidad = Val(txAlphaC.Text)
End Property

Public Property Let AlphaCantidad(ByVal vNewValue As Long)
    txAlphaC.Text = CStr(vNewValue)
End Property

Public Property Get AlphaColor() As Long
    AlphaColor = lblAlphaC.BackColor
End Property

Public Property Let AlphaColor(ByVal vNewValue As Long)
    lblAlphaC.BackColor = vNewValue
End Property

'=========================================================================0
'Otros
'=========================================================================0
Public Property Get ColorSel() As Long
    ColorSel = lblColorSel.BackColor
End Property

Public Property Let ColorSel(ByVal vNewValue As Long)
    lblColorSel.BackColor = vNewValue
End Property

'=========================================================================0
'Codigo
'=========================================================================0
Private Sub lblAlphaC_Click()
    GetColor lblAlphaC
End Sub

Private Sub lblColorSel_Click()
    GetColor lblColorSel
End Sub

Private Sub lblColorUnSel_Click()
    GetColor lblColorUnSel
End Sub

Private Sub lblFontCol_Click()
    GetColor lblFontCol
End Sub

Private Sub lblFontColCont_Click()
    GetColor lblFontColCont
End Sub

Private Sub GetColor(lblAux As Label)
    CD.ShowColor
    lblAux.BackColor = CD.RGBResult
End Sub

'=======================================
'FocoBotones
'=======================================
Private Sub txtX_GotFocus()
    txtX.SelStart = 0
    txtX.SelLength = Len(txtX.Text)
End Sub
Private Sub txtY_GotFocus()
    txtY.SelStart = 0
    txtY.SelLength = Len(txtY.Text)
End Sub
Private Sub txtAncho_GotFocus()
    txtAncho.SelStart = 0
    txtAncho.SelLength = Len(txtAncho.Text)
End Sub
Private Sub txtAlto_GotFocus()
    txtAlto.SelStart = 0
    txtAlto.SelLength = Len(txtAlto.Text)
End Sub
Private Sub txAlphaC_GotFocus()
    txAlphaC.SelStart = 0
    txAlphaC.SelLength = Len(txAlphaC.Text)
End Sub
Private Sub txtResAncho_GotFocus()
    txtResAncho.SelStart = 0
    txtResAncho.SelLength = Len(txtResAncho.Text)
End Sub
Private Sub txtResAlto_GotFocus()
    txtResAlto.SelStart = 0
    txtResAlto.SelLength = Len(txtResAlto.Text)
End Sub


