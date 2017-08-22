VERSION 5.00
Begin VB.Form FCfg 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Configuracion de teclas"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   435
      Left            =   1650
      TabIndex        =   9
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   435
      Left            =   870
      TabIndex        =   8
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   435
      Left            =   90
      TabIndex        =   7
      Top             =   30
      Width           =   735
   End
   Begin VB.PictureBox PicCONT 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   60
      ScaleHeight     =   2655
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   1080
      Width           =   4875
      Begin VB.PictureBox picLIST 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   60
         ScaleHeight     =   795
         ScaleWidth      =   3135
         TabIndex        =   3
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtDescr 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   0
            Left            =   2220
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   60
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox cmbKey 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   30
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   330
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label lblKey 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Visible         =   0   'False
            Width           =   2070
         End
      End
   End
   Begin VB.VScrollBar vsT 
      Height          =   4755
      Left            =   7170
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de teclas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   6840
   End
End
Attribute VB_Name = "FCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tK As cl_tbrSoftSelectKey

Public Sub ShowCFG(my_TK As cl_tbrSoftSelectKey)
    Set tK = my_TK
    
    'cargar la lista de teclas en el primer combo
    Dim H As Long
    For H = 1 To 144
        If tK.GetKDDesc(H) <> "" Then
            cmbKey(0).AddItem tK.GetKDDesc(H)
        End If
    Next H
    
    UnloadAll

    For H = 1 To tK.GetMaxCfg
        Load lblKey(H)
        Load cmbKey(H)
        Load txtDescr(H)
        
        CopyCombo cmbKey(0), cmbKey(H)
        
        lblKey(H).Caption = tK.GetName(H)
        cmbKey(H).ListIndex = GetListIndex(tK.GetKeyCode(H))
        txtDescr(H).Text = tK.GetDescr(H)
        
        lblKey(H).Top = txtDescr(H - 1).Top + txtDescr(H - 1).Height + 30
        cmbKey(H).Top = lblKey(H).Top + lblKey(H).Height
        txtDescr(H).Top = lblKey(H).Top
        
        lblKey(H).Visible = True
        cmbKey(H).Visible = True
        txtDescr(H).Visible = True
        
        picLIST.Height = txtDescr(H).Top + txtDescr(H).Height
    Next H
    
    AcomodarVST
    
    vsT.Value = 0
    
    Me.Show
End Sub

Private Sub CopyCombo(cmbOrig As ComboBox, cmbDest As ComboBox)
    cmbDest.Clear
    Dim H As Long
    For H = 0 To cmbOrig.ListCount - 1
        cmbDest.AddItem cmbOrig.List(H)
    Next H
End Sub

Private Sub AcomodarVST()
    Dim lMax As Long
    lMax = picLIST.Height - PicCONT.Height
    
    If lMax > 0 Then
        vsT.Max = picLIST.Height - PicCONT.Height
    Else
        vsT.Max = 0
    End If
    
    vsT.Min = 0
    
    Dim SCh As Long, LCh As Long
    SCh = CLng((vsT.Max - vsT.Min) / 20)
    LCh = CLng((vsT.Max - vsT.Min) / 5)
    If SCh > 0 And LCh > 0 Then
        vsT.SmallChange = SCh
        vsT.LargeChange = LCh
        vsT.Enabled = True
    Else
        vsT.Enabled = False
    End If
    
End Sub

Private Sub UnloadAll()
    Dim H As Long
    For H = 1 To lblKey.Count - 1
        Unload lblKey(H)
        Unload cmbKey(H)
        Unload txtDescr(H)
    Next H
End Sub

Private Sub Command1_Click()
    tK.SaveCfg "" 'ya tengo un archivo supestamente aqui
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtDescr(0).Locked = True
    
    txtDescr(0).Top = -txtDescr(0).Height
    lblKey(0).Top = -txtDescr(0).Height
    cmbKey(0).Top = -txtDescr(0).Height + lblKey(0).Height
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    Label1.Left = 0
    Label1.Width = Me.Width
    Label1.Top = Command1.Top + Command1.Height + 60
    
    PicCONT.Top = Label1.Top + Label1.Height + 60
    PicCONT.Left = 0
    PicCONT.Width = Me.Width - vsT.Width - 120
    PicCONT.Height = Me.Height - PicCONT.Top - 500
        
    picLIST.Width = PicCONT.Width
    picLIST.Left = 0
    
    lblKey(0).Width = (PicCONT.Width / 2) - 90
    cmbKey(0).Width = lblKey(0).Width
    txtDescr(0).Width = lblKey(0).Width
    
    Dim H As Long
    For H = 1 To lblKey.Count - 1
        lblKey(H).Width = lblKey(0).Width
        cmbKey(H).Width = lblKey(0).Width
        txtDescr(H).Width = lblKey(0).Width
        
        lblKey(H).Left = picLIST.Left + 30
        cmbKey(H).Left = lblKey(H).Left
        txtDescr(H).Left = lblKey(H).Left + lblKey(H).Width + 30
    Next H
    
    vsT.Top = PicCONT.Top
    vsT.Height = PicCONT.Height
    vsT.Left = PicCONT.Left + PicCONT.Width
    
    'acomodar maximos y minimos del vsT
    AcomodarVST
End Sub

Private Function GetListIndex(lKC As Long) As Long
    'ver que listIndex del combo tiene el keycode que queremos mostrar
    Dim H As Long, SP() As String
    For H = 0 To cmbKey(0).ListCount - 1
        SP = Split(cmbKey(0).List(H))
        If CLng(SP(0)) = lKC Then
            GetListIndex = H
        End If
    Next H
End Function

Private Sub vsT_Change()
    picLIST.Top = -vsT.Value
End Sub
