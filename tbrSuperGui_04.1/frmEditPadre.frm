VERSION 5.00
Begin VB.Form frmEditPadre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "img Fondo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3300
      Width           =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "guardar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3810
      Width           =   1050
   End
   Begin VB.TextBox txName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   9
      Text            =   "label"
      Top             =   2820
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "como"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3810
      Width           =   1050
   End
   Begin VB.ComboBox cmbAlignV 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditPadre.frx":0000
      Left            =   150
      List            =   "frmEditPadre.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1770
      Width           =   1965
   End
   Begin VB.ComboBox cmbAlignH 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditPadre.frx":0038
      Left            =   150
      List            =   "frmEditPadre.frx":0045
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   1965
   End
   Begin VB.CheckBox chkEstirable 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estirable"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   2190
      Width           =   1665
   End
   Begin VB.Frame frRect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rect"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2025
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Text            =   "0"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Text            =   "0"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Text            =   "0"
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Text            =   "0"
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2550
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "frmEditPadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O As ObjFullPadre  'padre QueRepresenta

Private Sub chkEstirable_Click()
    O.GetSgoInterno.Estirable = CBool(chkEstirable.Value)
End Sub

Private Sub cmbAlignH_Click()
    O.GetSgoInterno.AlignementH = cmbAlignH.ListIndex
End Sub

Private Sub cmbAlignV_Click()
    O.GetSgoInterno.AlignementV = cmbAlignV.ListIndex
End Sub

'cargar la imagen de fondo!
Private Sub cmdAddFile_Click(Index As Integer)
    
    Dim CM As New CommonDialog
    CM.DialogTitle = "Cargar imagen de fondo de formulario"
    CM.Filter = "Imagenes NO png|*.jpg; *.jpeg; *.gif"
    CM.ShowOpen
    Dim F As String
    
    F = CM.FileName
    If F = "" Then Exit Sub
    
    O.pathImgFondo = F
    'dibujarla!!
End Sub

Private Sub Command1_Click()
    Dim c As New CommonDialog
    c.ShowSave
    Dim F As String
    F = c.FileName
    
    If F <> "" Then
        Dim J As Long
        J = O.Save(F)
        
        If J = 0 Then
            O.PintarFondo
            MsgBox "Se grabo ok"
        Else
            MsgBox "Error al grabar: " + CStr(J)
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    O.Save
    O.PintarFondo
End Sub

Private Sub txName_Change()
    If txName.Text <> "" Then O.sName = txName.Text
End Sub

Private Sub txtRect_Validate(Index As Integer, Cancel As Boolean)

    If IsNumeric(txtRect(Index).Text) = False Then
        Cancel = True
        Exit Sub
    End If

    Select Case Index
        Case 0
            O.GetSgoInterno.X = CLng(txtRect(0).Text)
        Case 1
            O.GetSgoInterno.Y = CLng(txtRect(1).Text)
        Case 2
            O.GetSgoInterno.W = CLng(txtRect(2).Text)
        Case 3
            O.GetSgoInterno.H = CLng(txtRect(3).Text)
    
    End Select
    
End Sub

Public Sub SetObjPadre(obj As ObjFullPadre)
    
    Set O = obj
        
    txtRect(0).Text = O.GetSgoInterno.X
    txtRect(1).Text = O.GetSgoInterno.Y
    txtRect(2).Text = O.GetSgoInterno.W
    txtRect(3).Text = O.GetSgoInterno.H
    
    cmbAlignH.ListIndex = O.GetSgoInterno.AlignementH
    cmbAlignV.ListIndex = O.GetSgoInterno.AlignementV
    
    chkEstirable.Value = Abs(CLng(O.GetSgoInterno.Estirable))
    
    txName.Text = O.sName

End Sub
