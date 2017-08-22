VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvert 
   Caption         =   "Convert BMP to JPG"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadFile 
      Caption         =   "Carica Bitmap..."
      Height          =   615
      Left            =   2910
      Picture         =   "frmConvert.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   5580
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboQuality 
      Height          =   315
      Left            =   870
      TabIndex        =   2
      Top             =   90
      Width           =   1500
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Converti in JPG..."
      Height          =   615
      Left            =   4500
      Picture         =   "frmConvert.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.PictureBox picBitmap 
      AutoSize        =   -1  'True
      Height          =   2265
      Left            =   180
      Picture         =   "frmConvert.frx":0C5E
      ScaleHeight     =   2205
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   870
      Width           =   2400
   End
   Begin VB.Label lblBitmap 
      AutoSize        =   -1  'True
      Caption         =   "Immagine sorgente:"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label lblQuality 
      AutoSize        =   -1  'True
      Caption         =   "Qualità:"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/ LEGGERE IL FILE LEGGIMI.TXT
'/ Per convertire un file BMP in JPG, indicando la qualità di compressione.
'/
'/ Nota: il file Bitmap deve essere presente sul disco!
'/ Quindi non è possibile salvare un'immagine che sia solo in memoria,
'/ in questo caso occorre quindi prima salvare in formato Bitmap e poi
'/ eseguire la conversione.
'/ -------------------------------------------------------------------
'/ gibra@amc2000.it
'/ -------------------------------------------------------------------
Private Declare Function BMPToJPG Lib "converter.dll" (ByVal InputFilename As String, ByVal OutputFilename As String, ByVal Quality As Long) As Integer

Dim sSrcFile As String   '/ File BMP sorgente
Dim sDestFile As String  '/ file JPG di destinazione
Dim lQuality As Long     '/ qualita' JPG


Private Sub cboQuality_Click()
    lQuality = cboQuality.ListIndex * 10
End Sub


Private Sub cmdConvert_Click()
    
    '/ if bitmap is loaded from disk use same name
    If sSrcFile = "" Then
        sSrcFile = App.Path & "\bmp2jpg.bmp"
    End If
    
    On Error Resume Next
    With cdl1
        .CancelError = True
        .DefaultExt = ".jpg"
        .Filter = "Jpeg files (*.jpg)|*.jpg"
        .FileName = Left(sSrcFile, Len(sSrcFile) - 4)
        .Flags = cdlOFNNoChangeDir Or cdlOFNOverwritePrompt
        .ShowSave
        If Err.Number = cdlCancel Then Exit Sub
        sDestFile = .FileName
    End With
    
    If MsgBox("Convertire il file:" & vbCr & UCase(sSrcFile) & vbCr & " in " & vbCr & UCase(sDestFile) & "?", vbInformation, "Convert BMP to JPG") = vbCancel Then Exit Sub
    If Dir(sSrcFile) = "" Then
        '/ source file not exists, first must to save as bitmap
        SavePicture picBitmap, sSrcFile
    End If
    
    '/ convert Bitmap to Jpeg
    If BMPToJPG(sSrcFile, sDestFile, lQuality) <> 0 Then
        MsgBox "Errore " & Err.Number & vbCr & Err.Description, vbCritical, "Operazione fallita"
    Else
        '/ tutto bene!
        MsgBox "Il file " & UCase(sSrcFile) & " è stato convertito.", vbInformation, "Operazione riuscita"
    End If

End Sub


Private Sub cmdLoadFile_Click()

    On Error Resume Next
    With cdl1
        .CancelError = True
        .Filter = "Bitmap *.BMP|*.bmp"
        .ShowOpen
        If Err.Number = cdlCancel Then Exit Sub
        
        '/ Load bitmap
        sSrcFile = .FileName
        picBitmap.Picture = LoadPicture(.FileName)
    End With
    
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 10 To 100 Step 10
        cboQuality.AddItem CStr(i)
    Next i
    cboQuality.ListIndex = 9
    
End Sub


