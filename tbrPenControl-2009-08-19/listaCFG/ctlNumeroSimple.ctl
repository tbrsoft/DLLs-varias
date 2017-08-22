VERSION 5.00
Begin VB.UserControl ctlNumeroSimple 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton btOk 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   930
      TabIndex        =   5
      Top             =   2430
      Width           =   1000
   End
   Begin VB.CommandButton btCa 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   1785
      TabIndex        =   4
      Top             =   2415
      Width           =   1000
   End
   Begin VB.TextBox lbTITULO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   330
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   330
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.TextBox txtSimple 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1020
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lbDer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2700
      TabIndex        =   3
      Top             =   1485
      Width           =   150
   End
   Begin VB.Label lbIzq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   855
      TabIndex        =   2
      Top             =   1470
      Width           =   150
   End
End
Attribute VB_Name = "ctlNumeroSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'----------------------------------------------

Private manageNUM As New clsVERumeros

Public Event ClickOK()
Public Event ClickCancel()

Private Sub btCa_Click()
    RaiseEvent ClickCancel
End Sub

Private Sub btOk_Click()
    RaiseEvent ClickOK
End Sub

Public Function setManager(mng As clsVERumeros)
    Set manageNUM = mng
    
    If manageNUM.Valor = -9999 Then  'al incializarse queda asi para saber si ya se eligio algo o va el predeterminado
        txtSimple.Text = manageNUM.Predeterminado
    Else
        txtSimple.Text = manageNUM.Valor
    End If
    
    txtSimple.Visible = True
    
End Function

Public Function getManager() As clsVERumeros
    Set getManager = manageNUM
End Function

Public Sub SetTitulo(t As String)
    lbTITULO.Text = t
End Sub

Private Sub lbDer_Click()
    SelNext
End Sub

Private Sub lbIzq_Click()
    SelPrev
End Sub

Private Sub txtSimple_Change()
    manageNUM.ValTMP = txtSimple.Text
End Sub

Private Sub UserControl_Initialize()
    lbTITULO.Font = "Verdana"
    lbTITULO.FontSize = 8
    lbTITULO.FontBold = False
    lbTITULO.ForeColor = RGB(80, 80, 80)
    lbTITULO.BackColor = vbWhite
    
    txtSimple.Visible = False
    
    btOk.Font = "Verdana"
    btOk.FontSize = 8
    btOk.FontBold = False
    btCa.Font = "Verdana"
    btCa.FontSize = 8
    btCa.FontBold = False
End Sub

Private Sub UserControl_Resize()
    lbTITULO.Left = 45
    lbTITULO.Width = UserControl.Width - 8
    lbTITULO.Top = 0
    lbTITULO.Height = 1200
    
    'txtSimple.Top = (UserControl.Height - (lbTITULO.Top + lbTITULO.Height)) / 2 - txtSimple.Height / 2
    txtSimple.Top = 10
    txtSimple.Left = UserControl.Width / 2 - txtSimple.Width / 2
    txtSimple.Locked = True 'use texto y no label por que quizas me guste que lo editen a futuro
    
    lbIzq.Top = txtSimple.Top + 30
    lbIzq.Left = txtSimple.Left - lbIzq.Width
    lbDer.Top = lbIzq.Top
    lbDer.Left = txtSimple.Left + txtSimple.Width
    
    btOk.Top = UserControl.Height - btOk.Height - 30
    btOk.Left = 60
    btCa.Top = btOk.Top
    btCa.Left = btOk.Left + btOk.Width + 30
    
End Sub

Public Sub SelNext()
    terr.Anotar "qdn4"
    
    If manageNUM.ValTMP + manageNUM.Step > manageNUM.MaxVal Then
        manageNUM.ValTMP = manageNUM.MinVal
    Else
        manageNUM.ValTMP = manageNUM.ValTMP + manageNUM.Step
    End If
    
    terr.Anotar "qdo4", manageNUM.ValTMP
    txtSimple.Text = CStr(manageNUM.ValTMP)
End Sub

Public Sub SelPrev()
    terr.Anotar "qdn6"
    
    If manageNUM.ValTMP - manageNUM.Step < manageNUM.MinVal Then
        manageNUM.ValTMP = manageNUM.MaxVal
    Else
        manageNUM.ValTMP = manageNUM.ValTMP - manageNUM.Step
    End If
    
    terr.Anotar "qdo6", manageNUM.ValTMP
    txtSimple.Text = CStr(manageNUM.ValTMP)
End Sub

Public Sub ImitarFondo(hdcPadre As Long, cX As Long, cY As Long)
    BitBlt UserControl.HDC, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, cX, cY, SRCCOPY
    'BitBlt UserControl.hdc, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, UserControl.CurrentX, UserControl.CurrentY, SRCCOPY
    UserControl.Refresh
End Sub

