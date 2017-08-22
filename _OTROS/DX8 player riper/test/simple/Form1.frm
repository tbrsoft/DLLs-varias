VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1035
      Index           =   7
      Left            =   9090
      ScaleHeight     =   975
      ScaleWidth      =   285
      TabIndex        =   49
      Top             =   7140
      Width           =   345
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1305
      Index           =   6
      Left            =   7950
      ScaleHeight     =   1245
      ScaleWidth      =   1035
      TabIndex        =   48
      Top             =   7020
      Width           =   1095
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   5
      Left            =   6960
      ScaleHeight     =   795
      ScaleWidth      =   165
      TabIndex        =   47
      Top             =   7110
      Width           =   225
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1065
      Index           =   4
      Left            =   6360
      ScaleHeight     =   1005
      ScaleWidth      =   375
      TabIndex        =   46
      Top             =   7080
      Width           =   435
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1425
      Index           =   3
      Left            =   4650
      ScaleHeight     =   1365
      ScaleWidth      =   675
      TabIndex        =   45
      Top             =   7080
      Width           =   735
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1065
      Index           =   2
      Left            =   3210
      ScaleHeight     =   1005
      ScaleWidth      =   1365
      TabIndex        =   44
      Top             =   7080
      Width           =   1425
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1455
      Index           =   1
      Left            =   690
      ScaleHeight     =   1395
      ScaleWidth      =   2145
      TabIndex        =   43
      Top             =   6990
      Width           =   2205
   End
   Begin VB.PictureBox picVis 
      BackColor       =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   60
      ScaleHeight     =   1395
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   6990
      Width           =   615
   End
   Begin VB.HScrollBar hsPAN 
      Height          =   345
      Index           =   3
      Left            =   7350
      Max             =   1000
      Min             =   -1000
      TabIndex        =   41
      Top             =   6570
      Width           =   2265
   End
   Begin VB.HScrollBar hsVOL 
      Height          =   345
      Index           =   3
      Left            =   7350
      Max             =   1000
      TabIndex        =   40
      Top             =   6090
      Width           =   2265
   End
   Begin VB.HScrollBar hsPAN 
      Height          =   345
      Index           =   2
      Left            =   4980
      Max             =   1000
      Min             =   -1000
      TabIndex        =   39
      Top             =   6570
      Width           =   2265
   End
   Begin VB.HScrollBar hsVOL 
      Height          =   345
      Index           =   2
      Left            =   4980
      Max             =   1000
      TabIndex        =   38
      Top             =   6090
      Width           =   2265
   End
   Begin VB.HScrollBar hsPAN 
      Height          =   345
      Index           =   1
      Left            =   2550
      Max             =   1000
      Min             =   -1000
      TabIndex        =   37
      Top             =   6570
      Width           =   2265
   End
   Begin VB.HScrollBar hsVOL 
      Height          =   345
      Index           =   1
      Left            =   2550
      Max             =   1000
      TabIndex        =   36
      Top             =   6090
      Width           =   2265
   End
   Begin VB.HScrollBar hsPAN 
      Height          =   345
      Index           =   0
      Left            =   30
      Max             =   1000
      Min             =   -1000
      TabIndex        =   35
      Top             =   6570
      Width           =   2265
   End
   Begin VB.HScrollBar hsVOL 
      Height          =   345
      Index           =   0
      Left            =   30
      Max             =   1000
      TabIndex        =   34
      Top             =   6090
      Width           =   2265
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop"
      Height          =   435
      Index           =   3
      Left            =   9120
      TabIndex        =   29
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPausa 
      Caption         =   "Paus"
      Height          =   435
      Index           =   3
      Left            =   8610
      TabIndex        =   28
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPLAY 
      Caption         =   "Play"
      Height          =   435
      Index           =   3
      Left            =   8100
      TabIndex        =   27
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdOPEN 
      Caption         =   "Abrir MP3"
      Height          =   435
      Index           =   3
      Left            =   7290
      TabIndex        =   25
      Top             =   1650
      Width           =   795
   End
   Begin VB.ListBox lstINFO 
      Height          =   3180
      Index           =   3
      Left            =   7290
      TabIndex        =   24
      Top             =   2100
      Width           =   2355
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "close"
      Height          =   345
      Index           =   3
      Left            =   7980
      TabIndex        =   23
      Top             =   5640
      Width           =   795
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop"
      Height          =   435
      Index           =   2
      Left            =   6720
      TabIndex        =   22
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPausa 
      Caption         =   "Paus"
      Height          =   435
      Index           =   2
      Left            =   6210
      TabIndex        =   21
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPLAY 
      Caption         =   "Play"
      Height          =   435
      Index           =   2
      Left            =   5700
      TabIndex        =   20
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdOPEN 
      Caption         =   "Abrir MP3"
      Height          =   435
      Index           =   2
      Left            =   4890
      TabIndex        =   18
      Top             =   1650
      Width           =   795
   End
   Begin VB.ListBox lstINFO 
      Height          =   3180
      Index           =   2
      Left            =   4890
      TabIndex        =   17
      Top             =   2100
      Width           =   2355
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "close"
      Height          =   345
      Index           =   2
      Left            =   5580
      TabIndex        =   16
      Top             =   5640
      Width           =   795
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop"
      Height          =   435
      Index           =   1
      Left            =   4290
      TabIndex        =   15
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPausa 
      Caption         =   "Paus"
      Height          =   435
      Index           =   1
      Left            =   3780
      TabIndex        =   14
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPLAY 
      Caption         =   "Play"
      Height          =   435
      Index           =   1
      Left            =   3270
      TabIndex        =   13
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdOPEN 
      Caption         =   "Abrir MP3"
      Height          =   435
      Index           =   1
      Left            =   2460
      TabIndex        =   11
      Top             =   1650
      Width           =   795
   End
   Begin VB.ListBox lstINFO 
      Height          =   3180
      Index           =   1
      Left            =   2460
      TabIndex        =   10
      Top             =   2100
      Width           =   2355
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "close"
      Height          =   345
      Index           =   1
      Left            =   3150
      TabIndex        =   9
      Top             =   5640
      Width           =   795
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop"
      Height          =   435
      Index           =   0
      Left            =   1860
      TabIndex        =   8
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPausa 
      Caption         =   "Paus"
      Height          =   435
      Index           =   0
      Left            =   1350
      TabIndex        =   7
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton cmdPLAY 
      Caption         =   "Play"
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   1650
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INI elegida"
      Height          =   495
      Left            =   8070
      TabIndex        =   5
      Top             =   660
      Width           =   1575
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "close"
      Height          =   345
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   5640
      Width           =   795
   End
   Begin VB.ListBox lstINFO 
      Height          =   3180
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   2100
      Width           =   2355
   End
   Begin VB.CommandButton cmdOPEN 
      Caption         =   "Abrir MP3"
      Height          =   435
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   1650
      Width           =   795
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8580
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTIME 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   8340
      TabIndex        =   33
      Top             =   5310
      Width           =   900
   End
   Begin VB.Label lblTIME 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5910
      TabIndex        =   32
      Top             =   5310
      Width           =   900
   End
   Begin VB.Label lblTIME 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3510
      TabIndex        =   31
      Top             =   5310
      Width           =   900
   End
   Begin VB.Label lblTIME 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1110
      TabIndex        =   30
      Top             =   5310
      Width           =   900
   End
   Begin VB.Label lblSTATUS 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS:"
      Height          =   285
      Index           =   3
      Left            =   7290
      TabIndex        =   26
      Top             =   5310
      Width           =   1005
   End
   Begin VB.Label lblSTATUS 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS:"
      Height          =   285
      Index           =   2
      Left            =   4890
      TabIndex        =   19
      Top             =   5310
      Width           =   1005
   End
   Begin VB.Label lblSTATUS 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS:"
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   12
      Top             =   5310
      Width           =   1005
   End
   Begin VB.Label lblSTATUS 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS:"
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   5310
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents T8 As clsTbrDx8ESP
Attribute T8.VB_VarHelpID = -1

Option Explicit

Private Sub cmdCLOSE_Click(Index As Integer)
    T8.Player_ClearMemory CLng(Index)
End Sub

Private Sub cmdOPEN_Click(Index As Integer)
    Dim lI As Long: lI = CLng(Index)
'    cualquiera sea el STREAM usado tira la lista de las extenciones compatibles!
    dlg.FileName = vbNullString
    dlg.Filter = "(*.MP3)|*.mp3"
    dlg.ShowOpen
    
    If dlg.FileName = vbNullString Then Exit Sub
    
    T8.Player_AbrirMP3 dlg.FileName, lI
    
    'mostrar los datos!
    
    With lstINFO(Index)
        .Clear
        .AddItem T8.Player_File
        .AddItem "BitRate: " + CStr(T8.Get_Info_BitRate(lI))
        .AddItem "BitPorSample: " + CStr(T8.Get_Info_BitsPorSample(lI))
        .AddItem "Canales: " + CStr(T8.Get_Info_Canales(lI))
        .AddItem "DuracionMS: " + CStr(T8.Get_Info_DuracionMiliSegundos(lI))
        .AddItem "Duracion: " + CStr(T8.Get_Info_DuracionTexto(lI))
        .AddItem "ID: " + CStr(T8.Get_Info_ID(lI))
        .AddItem "SampleRate: " + CStr(T8.Get_Info_SampleRate(lI))
        
        'poner los tags tambien
        Dim Ts() As String
        Ts = T8.Get_Tags(lI)
        Dim U As Long
        For U = 1 To UBound(Ts)
            .AddItem Ts(U)
        Next U
    End With
   
    
End Sub

Private Sub cmdPausa_Click(Index As Integer)
    Dim lI As Long: lI = CLng(Index)
    T8.Player_PausaMP3 lI
End Sub

Private Sub cmdPLAY_Click(Index As Integer)
    Dim lI As Long: lI = CLng(Index)
    T8.Player_PlayMP3 lI
End Sub

Private Sub cmdSTOP_Click(Index As Integer)
    Dim lI As Long: lI = CLng(Index)
    T8.Player_StopMP3 lI
End Sub

Private Sub Command3_Click()
    T8.InicializarPlaca 44100, 2, 16, List1
End Sub

Private Sub Form_Load()
    Set T8 = New clsTbrDx8ESP
    T8.LeerPlacas
    Dim U As Long
    
    For U = 1 To T8.TotalPlacas
        List1.AddItem T8.Get_Placa_ID(U)
        List1.AddItem "   " + T8.Get_Placa_Nombre_ByIndex(U)
        List1.AddItem "   " + T8.Get_Placa_Descripcion_ByIndex(U)
    Next U
End Sub

Private Sub hsPAN_Change(Index As Integer)
    T8.Player_Set_Pan CLng(Index), hsPAN(Index)
End Sub

Private Sub hsVOL_Change(Index As Integer)
    T8.Player_Set_Volumen CLng(Index), hsVOL(Index)
End Sub

Private Sub T8_CancionCambiaEstado(Index As Long, IdSonido As Long, NuevoEstado As String)
    lblSTATUS(Index) = NuevoEstado
End Sub

Private Sub T8_CancionCorriendo(Index As Long, IdSonido As Long, strTime As String, MiliSecPlayed As Long, MiliSecToEnd As Long)
      lblTIME(Index) = MiliSecPlayed
      lblTIME(Index).Refresh
End Sub
