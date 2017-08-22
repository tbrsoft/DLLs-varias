VERSION 5.00
Begin VB.UserControl ctlTextoSimple 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ScaleHeight     =   4695
   ScaleWidth      =   5265
   Begin VB.CommandButton btCa 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   3870
      Width           =   1000
   End
   Begin VB.CommandButton btOk 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   1125
      TabIndex        =   5
      Top             =   3870
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
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   330
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.CommandButton cmdPred 
      Caption         =   "Limpiar / predeterminado"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   3045
      Width           =   2895
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "elegir"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3225
      TabIndex        =   2
      Top             =   3045
      Width           =   825
   End
   Begin VB.TextBox txtMulti 
      Appearance      =   0  'Flat
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
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2130
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.TextBox txtSimple 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   4005
   End
End
Attribute VB_Name = "ctlTextoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'----------------------------------------------

Private manageTXT As New clsVerTextoSimple
Private CM As New CommonDialog
Public Event ClickOK()
Public Event ClickCancel()

Private Sub btCa_Click()
    RaiseEvent ClickCancel
End Sub

Private Sub btOk_Click()
    RaiseEvent ClickOK
End Sub

Public Function setManager(mng As clsVerTextoSimple)
    Set manageTXT = mng
    
    txtMulti.Visible = (manageTXT.Multiline = 1)
    txtSimple.Visible = (manageTXT.Multiline = 0)
    
    If manageTXT.Valor = "NULL" Then 'al incializarse queda asi para saber si ya se eligio algo o va el predeterminado
        txtMulti.Text = manageTXT.Predeterminado
        txtSimple.Text = manageTXT.Predeterminado
    Else
        txtMulti.Text = manageTXT.Valor
        txtSimple.Text = manageTXT.Valor
    End If
    
    cmdShow.Visible = (manageTXT.ShowFile Or manageTXT.ShowFolder)
    cmdPred.Visible = (manageTXT.ShowFile Or manageTXT.ShowFolder)
    
    txtSimple.Locked = (manageTXT.ShowFile Or manageTXT.ShowFolder)
    txtMulti.Locked = (manageTXT.ShowFile Or manageTXT.ShowFolder)
    
End Function

Public Sub SetText(newT As String) 'lo agregue para poder traducir desde afuera
    txtMulti.Text = newT
    txtSimple.Text = newT
End Sub

Public Function getManager() As clsVerTextoSimple
    Set getManager = manageTXT
End Function

Public Sub SetTitulo(t As String)
    lbTITULO.Text = t
End Sub

Private Sub cmdPred_Click()
    txtMulti.Text = manageTXT.Predeterminado
End Sub

Private Sub cmdPred_GotFocus()
    'If manageTXT.Multiline Then txtMulti.SetFocus
    'If manageTXT.Multiline = 0 Then txtSimple.SetFocus
End Sub

Private Sub cmdShow_Click()
    
    Dim res As String
    
    CM.InitDir = manageTXT.InitialPath
    If manageTXT.ShowFile Then
        CM.Filter = manageTXT.Filter
        CM.ShowOpen
        
        res = CM.FileName
    End If
    
    If manageTXT.ShowFolder Then
        CM.ShowFolder
        
        res = CM.InitDir
    End If
    
    txtMulti.Text = res
    
End Sub

Private Sub cmdShow_GotFocus()
    'If manageTXT.Multiline Then txtMulti.SetFocus
    'If manageTXT.Multiline = 0 Then txtSimple.SetFocus
End Sub

Private Sub txtMulti_Change()
    manageTXT.ValorTMP = txtMulti
End Sub

Private Sub txtSimple_Change()
    manageTXT.ValorTMP = txtSimple
End Sub

Private Sub UserControl_Initialize()
    lbTITULO.Font = "Verdana"
    lbTITULO.FontSize = 8
    lbTITULO.FontBold = False
    lbTITULO.ForeColor = RGB(80, 80, 80)
    lbTITULO.BackColor = vbWhite
    
    
    txtSimple.Font = "Verdana"
    txtSimple.FontSize = 8
    txtSimple.FontBold = False
    txtSimple.ForeColor = RGB(80, 80, 80)
    txtSimple.BackColor = vbWhite
    
    txtMulti.Font = "Verdana"
    txtMulti.FontSize = 8
    txtMulti.FontBold = False
    txtMulti.ForeColor = RGB(80, 80, 80)
    txtMulti.BackColor = vbWhite
    
    cmdPred.Font = "Verdana"
    cmdPred.FontSize = 8
    cmdPred.FontBold = False
    cmdShow.Font = "Verdana"
    cmdShow.FontSize = 8
    cmdShow.FontBold = False

    btOk.Font = "Verdana"
    btOk.FontSize = 8
    btOk.FontBold = False
    btCa.Font = "Verdana"
    btCa.FontSize = 8
    btCa.FontBold = False

    
    txtMulti.Visible = False
    txtSimple.Visible = False
End Sub

Private Sub UserControl_Resize()

    lbTITULO.Left = 45
    lbTITULO.Width = UserControl.Width - (120 / 15)
    lbTITULO.Top = 0
    lbTITULO.Height = (1200 / 15)
    
    'txtSimple.Top = (UserControl.Height - (lbTITULO.Top + lbTITULO.Height)) / 2 - txtSimple.Height / 2
    txtSimple.Top = 60
    txtSimple.Left = UserControl.Width / 2 - txtSimple.Width / 2
    
    'txtMulti.Top = (UserControl.Height - (lbTITULO.Top + lbTITULO.Height)) / 2 - txtMulti.Height / 2
    txtMulti.Top = 60
    txtMulti.Left = UserControl.Width / 2 - txtMulti.Width / 2
    
    cmdPred.Top = txtMulti.Top + txtMulti.Height + 30
    cmdPred.Left = txtMulti.Left
    
    'cmdShow.Top = txtMulti.Top + txtMulti.Height + 30
    cmdShow.Top = cmdPred.Top + cmdShow.Height + 60
    'cmdShow.Left = cmdPred.Left + cmdPred.Width + 60
    cmdShow.Left = cmdPred.Left
    
    btOk.Top = UserControl.Height - btOk.Height - 30
    btOk.Left = 60
    btCa.Top = btOk.Top
    btCa.Left = btOk.Left + btOk.Width + 30
    
    txtSimple.Width = UserControl.Width - 120
    txtMulti.Width = txtSimple.Width
End Sub

Public Sub ImitarFondo(hdcPadre As Long, cX As Long, cY As Long)
'Public Sub ImitarFondo(hdcPadre As Long)
    BitBlt UserControl.HDC, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, cX, cY, SRCCOPY
    'BitBlt UserControl.hdc, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, UserControl.CurrentX, UserControl.CurrentY, SRCCOPY
    UserControl.Refresh
End Sub

