VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "convert64"
      Height          =   375
      Left            =   4710
      TabIndex        =   6
      Top             =   1560
      Width           =   1365
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   3120
      TabIndex        =   5
      Text            =   "clave"
      Top             =   1560
      Width           =   1425
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   2025
   End
   Begin VB.TextBox Text3 
      Height          =   1275
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3630
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Height          =   1275
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   5985
   End
   Begin VB.CommandButton Command1 
      Caption         =   "gen"
      Height          =   405
      Left            =   180
      TabIndex        =   1
      Top             =   1530
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   1275
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   150
      Width           =   5955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New tbrCrypto.Crypt

Private Sub Command1_Click()
    
    If cmbTipo.ListIndex = -1 Then Exit Sub
    
    Select Case cmbTipo.ListIndex
        Case 0
            Text2.Text = C.Base64String(Text1.Text, eB64_Encode)
            Text3.Text = C.Base64String(Text2.Text, eB64_Decode)
        Case 1
            Text2.Text = C.EncryptString(eMC_Blowfish, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_Blowfish, Text2.Text, Text4.Text, Check1)
        Case 2
            Text2.Text = C.EncryptString(eMC_CryptAPI, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_CryptAPI, Text2.Text, Text4.Text, Check1)
    
        Case 3
            Text2.Text = C.EncryptString(eMC_DES, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_DES, Text2.Text, Text4.Text, Check1)
        
        Case 4
            Text2.Text = C.EncryptString(eMC_Gost, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_Gost, Text2.Text, Text4.Text, Check1)
            
        Case 5
            Text2.Text = C.EncryptString(eMC_RC4, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_RC4, Text2.Text, Text4.Text, Check1)
    
        Case 6
            Text2.Text = C.EncryptString(eMC_Skipjack, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_Skipjack, Text2.Text, Text4.Text, Check1)
            
        Case 7
            Text2.Text = C.EncryptString(eMC_TEA, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_TEA, Text2.Text, Text4.Text, Check1)
            
        Case 8
            Text2.Text = C.EncryptString(eMC_Twofish, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_Twofish, Text2.Text, Text4.Text, Check1)
            
        Case 9
            Text2.Text = C.EncryptString(eMC_XOR, Text1.Text, Text4.Text, Check1)
            Text3.Text = C.DecryptString(eMC_XOR, Text2.Text, Text4.Text, Check1)
        
    End Select
End Sub

Private Sub Form_Load()
    cmbTipo.Clear
    cmbTipo.AddItem "Base 64"
    cmbTipo.AddItem "BlowFish"
    
    cmbTipo.AddItem "eMC_CryptAPI"
    cmbTipo.AddItem "eMC_DES"
    cmbTipo.AddItem "eMC_Gost"
    cmbTipo.AddItem "eMC_RC4"
    cmbTipo.AddItem "eMC_Skipjack"
    cmbTipo.AddItem "eMC_TEA"
    cmbTipo.AddItem "eMC_Twofish"
    cmbTipo.AddItem "eMC_XOR"
    
    cmbTipo.ListIndex = 0
    
End Sub

Private Sub cmbTipo_Click()
    Select Case cmbTipo.ListIndex
        Case 0
            Text4.Visible = False 'no lleva clave
            Check1.Visible = False
        Case Else
            Text4.Visible = True
            Check1.Visible = True
    End Select
End Sub

