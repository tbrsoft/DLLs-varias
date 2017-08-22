VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConnect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spojení"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      ScaleHeight     =   255
      ScaleWidth      =   2655
      TabIndex        =   24
      Top             =   400
      Width           =   2655
      Begin VB.OptionButton Option7 
         Caption         =   "GOPHER"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "FTP"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "HTTP"
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Text            =   "21"
      Top             =   1740
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   22
      Text            =   "zlg90"
      Top             =   1380
      Width           =   2655
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   2880
      TabIndex        =   21
      Text            =   "zlg"
      Top             =   1020
      Width           =   2655
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   2880
      TabIndex        =   20
      Text            =   "ftp.psiap.com"
      Top             =   660
      Width           =   2655
   End
   Begin VB.TextBox txtSiteName 
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   16
      Top             =   2100
      Width           =   1695
      Begin VB.OptionButton Option4 
         Caption         =   "Active tr."
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   320
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Passive tr."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   20
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   480
      Pattern         =   "*.ftp"
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0E5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   2700
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Save password"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ASCII"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2700
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Binary"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2700
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Annonymous"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2100
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.TreeView TView1 
      Height          =   3180
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2140
      _ExtentX        =   3784
      _ExtentY        =   5609
      _Version        =   393217
      Indentation     =   706
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   1785
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Passwrd:"
      Height          =   195
      Left            =   2200
      TabIndex        =   4
      Top             =   1425
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User ID:"
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   1035
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "FTP://"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   705
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   165
      Width           =   465
   End
End
Attribute VB_Name = "FrmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tvNode As Node
Dim strFileName
Private FtpService As Integer

Private Sub Command1_Click()
If txtSiteName.Text = "" Or txtServer.Text = "" Then
    MsgBox "Fill something in the first two text boxes!", vbExclamation
    Exit Sub
End If
strFileName = App.Path & "\" & txtSiteName.Text & ".ftp"
writeprivateprofilestring "PROFIL", "sname", txtSiteName.Text, strFileName
writeprivateprofilestring "PROFIL", "adresa", txtServer.Text, strFileName
writeprivateprofilestring "PROFIL", "ID", txtUser.Text, strFileName
If Check2.Value = 1 And txtPass.Text <> "" Then
        writeprivateprofilestring "PROFIL", "pass", txtPass.Text, strFileName
Else: writeprivateprofilestring "PROFIL", "pass", "", strFileName
End If
If txtPort.Text <> "" Then
        writeprivateprofilestring "PROFIL", "port", txtPort.Text, strFileName
Else: writeprivateprofilestring "PROFIL", "port", "", strFileName
End If
Nodes
Command5.Enabled = False
Nodes
MsgBox "Profil Saved", vbInformation
End Sub

Private Sub Command2_Click()
txtSiteName.Text = ""
txtServer.Text = ""
If Check1.Value = 1 Then
    txtUser.Text = "annonymous"
    txtPass.Text = "user@unknown.com"
    Else: txtPass.Text = ""
          txtUser.Text = ""
End If
txtPort.Text = "21"
Command5.Enabled = False
End Sub

Private Sub Command3_Click()
    Dim Service As Long
    'On Error GoTo Err
         If Len(txtServer.Text) <= 6 Then
              MsgBox "Wrong adress!"
              Exit Sub
         End If
    frmmain.TView1.Nodes.Clear
    frmmain.txtInfo.Text = ""
    Adresa = txtServer.Text
    ID = txtUser.Text
    Pass = txtPass.Text
    Port = txtPort.Text
    Klic = ""
    If txtSiteName.Text = "" Then
        txtSiteName.Text = App.Title
    End If
    If Option1.Value = 1 Then
        Transfer = FTP_TRANSFER_TYPE_BINARY
        frmmain.zBinary.Checked = True
        frmmain.zAscii.Checked = False
      Else
        Transfer = FTP_TRANSFER_TYPE_ASCII
        frmmain.zBinary.Checked = False
        frmmain.zAscii.Checked = True
    End If
    If Option3.Value = 1 Then
        Service = INTERNET_FLAG_PASSIVE
        frmmain.zPassive.Checked = True
      Else
        Service = INTERNET_FLAG_EXISTING_CONNECT
        frmmain.zPassive.Checked = False
    End If
    session = InternetOpen(txtSiteName.Text, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    Me.Hide
    If session <> 0 Then
        frmmain.txtInfo.SelText = Date & ", " & Time & " *** " & UCase(txtSiteName.Text) & " ***" & vbCrLf & Time & " > Connecting to: " & Adresa & "..." & vbCrLf
        server = InternetConnect(session, Adresa, Port, ID, Pass, INTERNET_SERVICE_FTP, Service, &H0)
        If server = 0 Then
            MsgBox "Falló la conexion al servidor!", vbExclamation
            frmmain.txtInfo.SelText = Time & " > Connection to server failed." & vbCrLf
            InternetCloseHandle session
           Exit Sub
        Else
            frmmain.txtInfo.SelText = Time & " > Connected to servise, looking for host." & vbCrLf
            frmmain.StatusBar1.Panels(3).Text = "Onlain"
            adr = Space(260)
            FtpGetCurrentDirectory server, adr, Len(adr)
            Label3.Caption = adr
    '        adr = "/*.*"
            adr = Left(adr, InStr(1, adr, Chr(0)) - 1)
            adr = adr & IIf((Right(adr, 1) = "\"), "*.*", "\*.*")
            frmmain.txtInfo.SelText = Time & " > Connected to server." & vbCrLf
            Set tvNode = frmmain.TView1.Nodes.Add(, , "\", "..", 11)
            Klic = "\"
            frmmain.List
        End If
    Else
    MsgBox "Connection to service failed!", vbExclamation
    frmmain.txtInfo.SelText = "Connection to service failed." & vbCrLf
    InternetCloseHandle session
    Exit Sub
    End If
    Unload Me
    'Err: MsgBox Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Dim Result As Integer
Result = MsgBox("Are you sure you want to delete this profile: " & TView1.SelectedItem.Text & "?", vbYesNo)
If Result = vbYes Then
Kill App.Path & "\" & TView1.SelectedItem.Text
Nodes
End If
End Sub


Private Sub Form_Load()
File1.Path = App.Path & "\"
FtpService = 1
Nodes
End Sub

Private Sub Nodes()
Dim i As Integer

File1.Refresh
TView1.Nodes.Clear
Set tvNode = TView1.Nodes.Add(, , "r", "Profily", 1)
For i = 0 To File1.ListCount - 1
    Set tvNode = TView1.Nodes.Add("r", tvwChild, , File1.List(i), 2)
Next i
TView1.Nodes(1).Expanded = True
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    txtUser.Text = "annonymous"
    txtPass.Text = "user@unknown.com"
Else: txtUser.Text = ""
    txtPass.Text = ""
End If
Command5.Enabled = False
End Sub

Private Sub Check2_Click()
Command5.Enabled = False
End Sub

Private Sub Option1_Click()
Command5.Enabled = False
End Sub

Private Sub Option2_Click()
Command5.Enabled = False
End Sub

Private Sub Text1_Click(Index As Integer)
Command5.Enabled = False
End Sub

Private Sub Option5_Click()
Dim ad As String
If Option5.Value = True Then
    FtpService = 3
    txtPort.Text = "80"
    Label2.Left = 2240
    Label2.Caption = "HTTP://"
End If
End Sub

Private Sub Option6_Click()
Dim ad As String
If Option6.Value = True Then
    FtpService = 1
    txtPort.Text = "21"
    Label2.Left = 2160
    Label2.Caption = "FTP://"
End If
End Sub

Private Sub TView1_KeyUp(KeyCode As Integer, Shift As Integer)
Set tvNode = TView1.SelectedItem
If KeyCode = vbKeyDelete Then
    Command5_Click
End If
End Sub

Private Sub TView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim sname As String
Dim Adresa As String
Dim ID As String
Dim Pass As String
Dim Port As String
Dim szReturn As String, X As Integer

sname = Space(50)
Adresa = Space(75)
ID = Space(50)
Pass = Space(50)
Port = Space(50)
If Node.Key <> "r" Then
strFileName = App.Path & "\" & Node.Text
getprivateprofilestring "PROFIL", "sname", szReturn, sname, Len(sname), strFileName
getprivateprofilestring "PROFIL", "adresa", szReturn, Adresa, Len(Adresa), strFileName
getprivateprofilestring "PROFIL", "ID", szReturn, ID, Len(ID), strFileName
getprivateprofilestring "PROFIL", "pass", szReturn, Pass, Len(Pass), strFileName
getprivateprofilestring "PROFIL", "port", szReturn, Port, Len(Port), strFileName
txtSiteName.Text = sname
txtServer.Text = Adresa
txtUser.Text = ID
txtPass.Text = Pass
txtPort.Text = Port
If txtUser.Text = "annonymous" Then
    Check1.Value = 1
Else: Check1.Value = 0
End If
If txtPass.Text = "" Then
    Check2.Value = 0
Else: Check2.Value = 1
End If
Command5.Enabled = True
Else: Command5.Enabled = False
End If
End Sub
