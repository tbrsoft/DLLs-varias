VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   4665
      Left            =   5790
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   1845
      Left            =   7860
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0076
      Top             =   0
      Width           =   2325
   End
   Begin VB.CommandButton Command3 
      Caption         =   "crg quick"
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Top             =   210
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "crg"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   555
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gou"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   525
      Index           =   0
      Left            =   10380
      ScaleHeight     =   465
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   30
      Width           =   795
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Index           =   0
      Left            =   11220
      Stretch         =   -1  'True
      Top             =   30
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOP As New tbrAlotOfPictures.clsALotOfPictures
Dim IMGs() As String
Dim LastTime As Single
'Dim M As New tbrMEM

Private Sub Command1_Click()
    Picture1(0).AutoRedraw = True
    Log "Carga Imagenes iniciada..."
    Dim H As Long, H2 As Long
    For H = 1 To 10
        Randomize
            H2 = Int(Rnd * UBound(IMGs))
            Set Picture1(H).Picture = LOP.GetPicture(IMGs(H2), "sdsd")
            'LOP.Paint Picture1(H), IMGs(H2), "sdsd"
        Randomize
            H2 = Int(Rnd * UBound(IMGs))
            Image1(H).Picture = LOP.GetPicture(IMGs(H2), "sdsd")
            'LOP.Paint Image1(H), IMGs(H2), "sdsd"
    Next H
    Log "Cargo 10 imagenes"
End Sub

Private Sub Command2_Click()
    Cargar True
End Sub

Private Sub Cargar(bLoad As Boolean)
    Log "Carga Iniciada"
    ReDim IMGs(0)
    'cargar todas las imagenes
    Dim sJPG As String
    sJPG = Dir(App.Path + "\lap\*.jpg")
    LOP.ClearAll
    Do While sJPG <> ""
        If LOP.AddImage(App.Path + "\lap\" + sJPG, bLoad) = 0 Then
            ReDim Preserve IMGs(UBound(IMGs) + 1)
            IMGs(UBound(IMGs)) = App.Path + "\lap\" + sJPG
            'M.SetMomento CStr(FileLen(App.Path + "\lap\" + sJPG))
        Else
            Log "NO SE CARGO !!" + CStr(UBound(IMGs))
        End If
        sJPG = Dir
    Loop
    
    Log "Se cargaron " + CStr(UBound(IMGs)) + " imagenes"
    
    'Text2.Text = M.GetFullDetalles
    
End Sub

Private Sub Command3_Click()
    Cargar False
End Sub

Private Sub Form_Load()
    Dim H As Long
    For H = 1 To 10
        Load Picture1(H)
        Load Image1(H)
        Picture1(H).Top = Picture1(H - 1).Top + Picture1(H - 1).Height
        Image1(H).Top = Image1(H - 1).Top + Image1(H - 1).Height
        
        Picture1(H).Visible = True
        Image1(H).Visible = True
    Next H
    
    ReDim IMGs(0)
End Sub

Private Sub Log(T As String)
    List1.AddItem CS(CStr(Round(LastTime, 4)), 12) + " " + CS(CStr(Round(Timer - LastTime, 4)), 12) + " " + T, 0
    LastTime = Timer
End Sub

Private Function CS(Te As String, Total As Long) As String
    Dim T As String
    T = Space(Total - Len(Te)) + Te
    CS = T
End Function

