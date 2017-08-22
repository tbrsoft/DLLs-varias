VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   1320
   ClientTop       =   1395
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   6090
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      Filter          =   "*.png|*.png"
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   7155
      Left            =   3180
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7095
      ScaleWidth      =   7665
      TabIndex        =   4
      Top             =   540
      Width           =   7725
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   1980
         ScaleHeight     =   1605
         ScaleWidth      =   1785
         TabIndex        =   5
         Top             =   480
         Width           =   1785
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pngClass As New LoadPNG

Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub File1_Click()
    
    If File1.filename <> "" Then
        
        Picture1.Picture = LoadPicture("")
        pngClass.PicBox = Picture1
        pngClass.SetToBkgrnd False, Picture1.Left / 15, Picture1.Top / 15, 0, 0 '   'set to Background (True or false), x and y
        'x e y son las coordenadas dentro del picturebox que uso
        pngClass.BackgroundPicture = Picture2 'de donde se leen los pixeles que son transparentes (la imagen que este atras).
        pngClass.SetAlpha = True 'when Alpha then alpha
        pngClass.SetTrans = True 'when transparent Color then transparent Color
        
        ' Visualiza el Archivo en Picture1
        pngClass.OpenPNG File1.Path & "\" & File1.filename
    
    End If
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.png"
End Sub

