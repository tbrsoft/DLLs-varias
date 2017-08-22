VERSION 5.00
Object = "*\A..\ObjetoXControl.vbp"
Begin VB.Form f1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debuj tbrObjetoX"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   679
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   832
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "SetPropiedad"
      Height          =   300
      Left            =   2130
      TabIndex        =   3
      Top             =   45
      Width           =   1860
   End
   Begin tbrObjetoX_ocx.tbrObjetoX OX 
      Height          =   9000
      Left            =   90
      TabIndex        =   2
      Top             =   420
      Width           =   12000
      _ExtentX        =   13679
      _ExtentY        =   6641
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar"
      Height          =   300
      Left            =   1095
      TabIndex        =   1
      Top             =   45
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   1020
   End
   Begin VB.Label lblProp 
      Caption         =   "ObjetoX"
      Height          =   660
      Left            =   105
      TabIndex        =   4
      Top             =   9480
      Width           =   12000
   End
End
Attribute VB_Name = "f1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpLbl As Label

Dim rX As Long
Dim rY As Long
Dim rAncho As Long
Dim rAlto As Long

Private Sub Command1_Click()
    OX.AgregarObjetoX
End Sub

Private Sub Command2_Click()
    Dim ix As Long
    ix = OX.GetObjetoXSeleccionadoIndex
    OX.QuitarObjetoX ix
End Sub

Private Sub Command3_Click()
    Dim ix As Long
    ix = OX.GetObjetoXSeleccionadoIndex
    OX.SetPropiedadesObjetoX ix, InputBox("Pone alguna propiedad", "Debug ObjetoX")
End Sub

Private Sub OX_SeleccionaItem(IndexItem As Long)
    lblProp.Caption = OX.GetPropiedadesObjetoX(IndexItem)
    GetRectaObjetoX IndexItem
    lblProp.Caption = lblProp.Caption + vbCrLf
    lblProp.Caption = lblProp.Caption + "X:" + CStr(rX) + " "
    lblProp.Caption = lblProp.Caption + "Y:" + CStr(rY) + " "
    lblProp.Caption = lblProp.Caption + "Ancho:" + CStr(rAncho) + " "
    lblProp.Caption = lblProp.Caption + "Alto:" + CStr(rAlto)
End Sub

Sub GetRectaObjetoX(Index As Long)
    Set tmpLbl = OX.GetObjetoX(Index)
    
    rX = tmpLbl.Left
    rY = tmpLbl.Top
    rAncho = tmpLbl.Width
    rAlto = tmpLbl.Height
End Sub
