VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Demo de karaoke en formato MNx"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
   Icon            =   "F1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9960
      ScaleHeight     =   375
      ScaleWidth      =   645
      TabIndex        =   6
      Top             =   300
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picKAR 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   60
      ScaleHeight     =   1875
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      Begin VB.Shape shKAR 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         Height          =   345
         Left            =   1980
         Shape           =   3  'Circle
         Top             =   210
         Width           =   405
      End
      Begin VB.Label LF1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait 99"
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
         Height          =   300
         Left            =   30
         TabIndex        =   2
         Top             =   1110
         Width           =   1700
      End
      Begin VB.Label lblTimeK 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   30
         TabIndex        =   1
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label lblTimeK2 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LF2 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait 99"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   60
         TabIndex        =   3
         Top             =   1140
         Width           =   1700
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   4605
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   9465
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AP As String
Dim WithEvents MP3 As tbrPlayer02.MainPlayer
Attribute MP3.VB_VarHelpID = -1
Dim ListaMN0() As String
Dim ListaMN1() As String
Dim indActual As Long
Dim FSO As New Scripting.FileSystemObject
Dim SF As String
'****************************************************************************************
'clave de los karaokes base de prueba
'****************************************************************************************
'kar de tbrSoft marcaregistrada creados por rava ale y manu 14578965412339782100548165164
'****************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Local Error Resume Next 'pueden poner una letra que no existe!

    If KeyCode > 64 And KeyCode <= vbKeyZ Then
        MP3.DoClose 2
        Dim uSel As Long
        uSel = KeyCode - 64
        If uSel < 1 Then uSel = 1
        If uSel > UBound(ListaMN1) Then uSel = UBound(ListaMN1)
        If uSel = 0 Then Exit Sub
        Play uSel
    End If
    
    If KeyCode = vbKeyEscape Then
        'borrar el temporal anterior!
        If FSO.FileExists(ListaMN0(indActual)) Then FSO.DeleteFile ListaMN0(indActual)
    
        Unload Me
    End If
    If KeyCode = vbKey0 Then
        picKAR.Visible = Not picKAR.Visible
        Label1.Visible = Not picKAR.Visible
    End If
        
End Sub

Private Sub Form_Load()
    picKAR.Visible = False
    
    Set MP3 = New tbrPlayer02.MainPlayer
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    SF = FSO.GetSpecialFolder(SystemFolder)
    If Right(SF, 1) <> "\" Then SF = SF + "\"
    
    'ver cuantos MN0 hay en la carpeta Path
    ReDim ListaMN0(0): ReDim ListaMN1(0)
    Dim A As Long, tmpSong As String
    'ponerlos todos en una matriz a partir de uno
    tmpSong = Dir(AP + "*.mn1")
    Label1 = "Haga esta lista visible/invisible presionando 0 (cero)" + vbCrLf + _
        "Solo presione la letra de la cancion que desea ejecutar. Puede " + _
        "interrumpir cualquier canción en cualquier momento. Presione ESCAPE para cerrar" + _
        vbCrLf + vbCrLf
    Do While tmpSong <> ""
        A = A + 1
        Label1 = Label1 + Chr(64 + A) + " = " + tmpSong + vbCrLf
        ReDim Preserve ListaMN0(A): ReDim Preserve ListaMN1(A)
        ListaMN0(A) = SF + CStr(CLng(Timer)) + CStr(A): ListaMN1(A) = AP + tmpSong
        tmpSong = Dir
    Loop
    
    If A = 0 Then
        Label1 = "NO HAY KARAOKES DISPONIBLES"
        Exit Sub
    End If
    
End Sub

Private Sub Play(i As Long) 'indice del archivo
    'borrar el temporal anterior!
    If FSO.FileExists(ListaMN0(indActual)) Then FSO.DeleteFile ListaMN0(indActual)
    
    indActual = i
    
    'desencriptar de mn1 a mn0
    Dim T(10) As String
    T(0) = "ma"
    T(1) = "r "
    T(2) = "ad"
    MP3.doTem True, "ka" + T(1) + "de tbrSoft " + T(0) + _
        "rcaregistr" + T(2) + "a cre" + T(2) + "os po" + T(1) + "rava ale y " + T(0) + _
        "nu 14578965412339782100548165164", ListaMN1(i), ListaMN0(i)

    Dim R As Long
    R = MP3.DoOpenKar(ListaMN0(i), picKAR, shKAR)
    
    If R = 1 Then 'hay algun error!
        Exit Sub
    End If
    
    MP3.DoPlayKar
    picKAR.Visible = True
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MP3.DoClose 2
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    Me.WindowState = 2

    Label1.Top = (Me.Height / 2 - Label1.Height / 2) - 1000
    Label1.Left = 100
    Label1.Left = Me.Width / 2 - Label1.Width / 2

    picKAR.Left = 0
    picKAR.Top = 0
    picKAR.Width = Me.Width '- (picKAR.Left * 3)
    
    picKAR.Height = Me.Height ' - picKAR.Top - 1500
    Me.AutoRedraw = True
    Picture1.Picture = LoadPicture(App.Path + "\fondo.jpg")
    Me.PaintPicture Picture1.Picture, 0, 0, Me.Width, Me.Height
    Me.Picture = Me.Image
End Sub

Private Sub MP3_EndPlay(iAlias As Long)
    'borrar el temporal anterior!
    If FSO.FileExists(ListaMN0(indActual)) Then FSO.DeleteFile ListaMN0(indActual)
    picKAR.Visible = False
    Label1.Visible = True
End Sub

Private Sub MP3_FaltaNextEvKAR(dMiliSec As Double)
    If (dMiliSec > 0) Then
        LF1 = Format(dMiliSec, "00")
    End If
    LF1.Visible = (dMiliSec > 0)
    
    LF1 = "wait " + LF1
    LF2 = LF1
    LF2.Visible = LF1.Visible
End Sub

Private Sub MP3_Played(SecondsPlayed As Long, iAlias As Long, MS As Long)
    lblTimeK = MP3.Falta(2)
    lblTimeK2 = lblTimeK
End Sub
