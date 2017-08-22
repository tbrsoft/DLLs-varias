VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMART CAMERA"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   FillColor       =   &H00FFFFFF&
   FillStyle       =   3  'Vertical Line
   ForeColor       =   &H00000000&
   Icon            =   "merger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   5280
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "By :                               Handri Toar Pangkerego Nim : 32102018     Elektronika Industri Politeknik Batam"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   1545
         Left            =   360
         Picture         =   "merger.frx":4F0A
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Panel"
      Height          =   1695
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Settings 
         Caption         =   "Settings"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Format 
         Caption         =   "Format"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer timer1 
      Interval        =   33
      Left            =   1800
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Video"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox Picture1 
         Height          =   3600
         Left            =   120
         ScaleHeight     =   236
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   316
         TabIndex        =   1
         Top             =   360
         Width           =   4800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "http://students.polibatam.ac.id/~m20218"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   840
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   4320
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'deklarasi web cam
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
'deklarasi delay
Private Declare Function GetTickCount Lib "kernel32" () As Long

'deklarasi parallel
Private Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

'constanta web cam
Private mCapHwnd As Long
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const COPY As Long = 1054
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Private Const GET_FRAME As Long = 1084

'variabel program
Dim P(40, 30) As Long                           ' penyimpanan nilai rgb setiap titik

Dim C As Long, Ca As Long                       ' nilai rgb dari suatu titik
Dim R As Integer, G As Integer, B As Integer    ' pemisahan RGB menjadi R, G, B gambar sekarang
Dim Ra As Integer, Ga As Integer, Ba As Integer ' pemisahan RGB menjadi R, G, B gamba sebelum

Dim Toleransi As Integer                        ' toleransi perbedaan warna yang dianggap pergerakan
Dim Gerak As Integer                            ' banyaknya titik yang berubah
Dim Fokus_X, Fokus_Y As Integer                 ' titik tengah pergerakan
Dim Jarak(2) As Long                            ' jumlah koordinat titik fokus pergerakan 0=x 1=y

Dim Motor_X, Motor_Y As Integer                 ' koordinat pergerakan motor
Dim Output, Out_X, Out_Y, Step As Integer       ' output to parallel
Dim Arah1 As String                             ' arah motor 1 (bawah)
Dim Arah2 As String                             ' arah motor 2 (atas)

'proses inisial awal
Private Sub Form_Load()
Toleransi = 30

Out_X = 1
Out_Y = 16

mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hwnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0
End Sub

'program ditutup
Private Sub Form_Unload(Cancel As Integer)
SendMessage mCapHwnd, DISCONNECT, 0, 0
End Sub

'tombol format diklik
Private Sub Format_Click()
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

'tombol setting diklik
Private Sub Settings_Click()
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
End Sub

'sub rutin pemangilan camera setiap 33ms
Private Sub Timer1_Timer()

'proses pengambilan gambar
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0
Picture1.Picture = Clipboard.GetData
Clipboard.Clear

Jarak(0) = "0"
Jarak(1) = "0"
Gerak = 0

'pendeteksian terhadap 1200 titik dalam gambar (40*30)
For i = 0 To 40
    For j = 0 To 30
    
        'proses pengambilan nilai RGB dan pemisahannya dari gambar sekarang
        C = Picture1.Point(i * 8, j * 8)                    ' jarak antar titik = 8 pixel, gambar sekarang
        R = C Mod 256                                       ' pemisahan Red dalam RGB
        G = (C \ 256) Mod 256                               ' pemisahan Green dalam RGB
        B = (C \ 256 \ 256) Mod 256                         ' pemisahan Blue dalam RGB
                            
        'proses pengambilan nilai RGB dan pemisahannya dari gambar sebelumnya
        Ca = P(i, j)
        Ra = Ca Mod 256
        Ga = (Ca \ 256) Mod 256
        Ba = (Ca \ 256 \ 256) Mod 256
        
        If P(i, j) = 0 Then
            P(i, j) = C
        Else
            
            'proses pembandingan kedua gambar
            If Abs(R - Ra) > Toleransi And Abs(G - Ga) > Toleransi And Abs(B - Ba) > Toleransi Then     'pendeteksian perubahan warna jika melebihi dari batas toleransi
                P(i, j) = Picture1.Point(i * 8, j * 8)
                'Picture1.Circle (i * 8, j * 8), 2, vbRed    ' terjadi perubahan warna
                Jarak(0) = Jarak(0) + (i * 8)               ' jumlah koordinat x seluruh gerakan
                Jarak(1) = Jarak(1) + (j * 8)               ' jumlah koordinat y seluruh gerakan
                Gerak = Gerak + 1                           ' banyaknya titik yang berubah
            End If
        End If
        
    Next j
Next i

'proses pencarian titik fokus gerakan
If Gerak > 10 And Gerak < 500 Then                          ' jika ada lebih dari 10 tiik yang berubah maka kamera bergerak

    Fokus_X = Int(Jarak(0)) \ Gerak                         ' koordinat X
    Fokus_Y = Int(Jarak(1)) \ Gerak                         ' koordinat Y
    'Picture1.Circle (Fokus_X, Fokus_Y), 25, vbBlue          ' fokus gerakan kamera pada lingkaran biru
    
    For s = 0 To 40                                         ' setelah motor bergerak maka
        For t = 0 To 30                                     ' penyimpanan nilai RGB untuk gambar
            P(s, t) = 0                                     ' sebelum menjadi 0 (kosong)
        Next t
    Next s
    
    'proses konversi titik fokus gerakan menjadi penggerak motor
    Motor_X = (160 - Fokus_X) \ 12
    Motor_Y = (120 - Fokus_Y) \ 12
    If Motor_X < 0 Then Arah1 = "KANAN" Else Arah1 = "KIRI"
    If Motor_Y < 0 Then Arah2 = "BAWAH" Else Arah2 = "ATAS"
    Motor Abs(Motor_Y), Arah2, Abs(Motor_X), Arah1          ' mengirimkan perintah ke motor untuk bergerak
    
End If
End Sub

' proses delay (dalam satuan milidetik
Sub Delay(wait As Long)
Dim LastTick As Long, CurrentTick As Long
LastTick = GetTickCount

Do
  CurrentTick = GetTickCount
  DoEvents
Loop Until (CurrentTick - LastTick) > wait
End Sub

'proses penggerak motor stepper
Sub Motor(Step_X As Integer, Arah_X As String, Step_Y As Integer, Arah_Y As String)
If Step_X > Step_Y Then Step = Step_X Else Step = Step_Y

For A = 1 To Step
    If A <= Step_X Then
    
        ' kamera bergerak kebawah
        If Arah_X = "BAWAH" Then
            Select Case Out_X
            Case 1:  Out_X = 3
            Case 3:  Out_X = 2
            Case 2:  Out_X = 6
            Case 6:  Out_X = 4
            Case 4:  Out_X = 12
            Case 12: Out_X = 8
            Case 8:  Out_X = 9
            Case 9:  Out_X = 1
            End Select
        End If
        
        'kamera bergerak keatas
        If Arah_X = "ATAS" Then
            Select Case Out_X
            Case 1:  Out_X = 9
            Case 3:  Out_X = 1
            Case 2:  Out_X = 3
            Case 6:  Out_X = 2
            Case 4:  Out_X = 6
            Case 12: Out_X = 4
            Case 8:  Out_X = 12
            Case 9:   Out_X = 8
            End Select
        End If
    End If
    
    If A <= Step_Y Then
    
        'kamera bergerak kekiri
        If Arah_Y = "KIRI" Then
            Select Case Out_Y
            Case 16:  Out_Y = 48
            Case 48:  Out_Y = 32
            Case 32:  Out_Y = 96
            Case 96:  Out_Y = 64
            Case 64:  Out_Y = 192
            Case 192: Out_Y = 128
            Case 128: Out_Y = 144
            Case 144: Out_Y = 16
            End Select
        End If
        
        'kamera bergerak kekanan
        If Arah_Y = "KANAN" Then
            Select Case Out_Y
            Case 16:  Out_Y = 144
            Case 48:  Out_Y = 16
            Case 32:  Out_Y = 48
            Case 96:  Out_Y = 32
            Case 64:  Out_Y = 96
            Case 192: Out_Y = 64
            Case 128: Out_Y = 192
            Case 144: Out_Y = 128
            End Select
        End If
    End If
    
    Out 888, (Out_X + Out_Y)                ' keluar keparalel
    Delay 1
Next A
End Sub
