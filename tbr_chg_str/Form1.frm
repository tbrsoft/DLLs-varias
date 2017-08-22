VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "un-crypt"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   1530
      Width           =   1005
   End
   Begin VB.ComboBox cmbUsados 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1980
      Width           =   10785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "convert64"
      Height          =   375
      Left            =   6180
      TabIndex        =   6
      Top             =   1500
      Width           =   1365
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   4650
      TabIndex        =   5
      Text            =   "clave"
      Top             =   1500
      Width           =   1425
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   2490
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1500
      Width           =   2025
   End
   Begin VB.TextBox Text3 
      Height          =   3195
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4890
      Width           =   11025
   End
   Begin VB.TextBox Text2 
      Height          =   2475
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2370
      Width           =   11085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "gen"
      Height          =   405
      Left            =   1110
      TabIndex        =   1
      Top             =   1470
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1275
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   30
      Width           =   10995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New tbrCrypto.Crypt
Dim FSO As New Scripting.FileSystemObject, fileUS As String
Private Sub cmbUsados_Click()
    On Local Error Resume Next
    
    Dim sp() As String
    sp = Split(cmbUsados, "|")
    
    Text1.Text = Trim(sp(0))
    cmbTipo.ListIndex = CLng(Trim(sp(1)))
    Text4.Text = Trim(sp(2))
    If CLng(Trim(sp(3))) = 0 Then
        Check1.Value = 0
    Else
        Check1.Value = 1
    End If
End Sub

Private Sub Command1_Click()
    
    If cmbTipo.ListIndex = -1 Then Exit Sub
    
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(Text1.Text) Then
                
        Dim tmp As String
        Dim filCRYP As String
        filCRYP = Text1.Text + ".ENCRYPT"
        
        'es un archivo, seguramente lo uso mucho, guardarlo
        Dim TE As TextStream, newR As String
        newR = Text1.Text + " | " + CStr(cmbTipo.ListIndex) + " | " + Text4.Text + " | " + CStr(Check1.Value)
        'ver si ya esta!
        Dim j As Long, existe As Boolean
        For j = 0 To cmbUsados.ListCount - 1
            If LCase(cmbUsados.List(j)) = LCase(newR) Then
                existe = True
                Exit For
            End If
        Next j
        If existe = False Then
            Set TE = FSO.OpenTextFile(fileUS, ForAppending, True)
                TE.WriteLine newR
            TE.Close
        End If
        
        Select Case cmbTipo.ListIndex
            Case 0
                If Check2.Value = 0 Then 'si es encriptar lo hago y despues desencripto
                    tmp = fileToStr(Text1.Text)
                    Text2.Text = C.Base64String(tmp, eB64_Encode)
                    strToFile Text2.Text, filCRYP
                Else 'tomo como que el archivo que se abre esta encriptado
                    Text2.Text = ""
                    FSO.CopyFile Text1.Text, filCRYP, True
                End If
                
                Text3.Text = C.Base64String(Text2.Text, eB64_Decode)
            Case 1
            
                If Check2.Value = 0 Then 'si es encriptar lo hago y despues desencripto
                    C.EncryptFile eMC_Blowfish, Text1.Text, filCRYP, Text4.Text, Check1
                    Text2.Text = fileToStr(filCRYP)
                Else 'tomo como que el archivo que se abre esta encriptado
                    Text2.Text = ""
                    FSO.CopyFile Text1.Text, filCRYP, True
                End If
                
                C.DecryptFile eMC_Blowfish, filCRYP, filCRYP + ".decr", Text4.Text, Check1
                Text3.Text = fileToStr(filCRYP + ".decr")
            Case 2
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_CryptAPI, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_CryptAPI, Text2.Text, Text4.Text, Check1)
            Case 3
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_DES, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_DES, Text2.Text, Text4.Text, Check1)
            Case 4
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_Gost, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_Gost, Text2.Text, Text4.Text, Check1)
            Case 5
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_RC4, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_RC4, Text2.Text, Text4.Text, Check1)
            Case 6
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_Skipjack, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_Skipjack, Text2.Text, Text4.Text, Check1)
            Case 7
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_TEA, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_TEA, Text2.Text, Text4.Text, Check1)
            Case 8
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_Twofish, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_Twofish, Text2.Text, Text4.Text, Check1)
            Case 9
                MsgBox "En desarrollo"
                'Text2.Text = C.EncryptString(eMC_XOR, Text1.Text, Text4.Text, Check1)
                'Text3.Text = C.DecryptString(eMC_XOR, Text2.Text, Text4.Text, Check1)
            
        End Select
    
    
    Else
    
        Select Case cmbTipo.ListIndex
            Case 0
                If Check2.Value = 0 Then
                    Text2.Text = C.Base64String(Text1.Text, eB64_Encode)
                    Text3.Text = C.Base64String(Text2.Text, eB64_Decode)
                Else
                    Text2.Text = C.Base64String(Text1.Text, eB64_Decode)
                    Text3.Text = C.Base64String(Text2.Text, eB64_Encode)
                End If
            Case 1
                If Check2.Value = 0 Then
                    Text2.Text = C.EncryptString(eMC_Blowfish, Text1.Text, Text4.Text, Check1)
                    Text3.Text = C.DecryptString(eMC_Blowfish, Text2.Text, Text4.Text, Check1)
                Else
                    Text2.Text = C.DecryptString(eMC_Blowfish, Text1.Text, Text4.Text, Check1)
                    Text3.Text = C.EncryptString(eMC_Blowfish, Text2.Text, Text4.Text, Check1)
                End If
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
    End If
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
    
    'cargar los ya usados que fueron archivos!
    fileUS = FSO.BuildPath(App.path, "usados.txt")
    
    If FSO.FileExists(fileUS) Then
        Dim TE As TextStream, r As String
        Set TE = FSO.OpenTextFile(fileUS)
            r = TE.ReadAll
        TE.Close
        
        cmbUsados.Clear
        Dim sp() As String
        sp = Split(r, vbCrLf)
        Dim j As Long
        For j = 0 To UBound(sp)
            cmbUsados.AddItem sp(j)
        Next j
        
    End If
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

Private Sub Text1_DblClick()
    Dim cmd As New CommonDialog
    cmd.ShowOpen
    
    Text1.Text = cmd.FileName
End Sub

Private Function fileToStr(fil As String) As String

    Dim res As String
    
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fil) = False Then
        fileToStr = ""
        Exit Function
    End If
        
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(fil, ForReading)
        fileToStr = TE.ReadAll
    TE.Close
    
End Function

Private Sub strToFile(s As String, f As String)
    Dim FSO As New Scripting.FileSystemObject
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(f, True)
        TE.Write s
    TE.Close
End Sub
