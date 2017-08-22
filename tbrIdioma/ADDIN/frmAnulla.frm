VERSION 5.00
Begin VB.Form frmAnulla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Descartar todo lo que tenga los prefijos elegidos"
      Height          =   405
      Left            =   1050
      TabIndex        =   2
      Top             =   4950
      Width           =   3975
   End
   Begin VB.ListBox lstAnulla 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   570
      Width           =   5835
   End
   Begin VB.Label Label1 
      Caption         =   "Se encontraron los siguientes prefijos. Indique cuales con seguridad no se necesitan traducir"
      Height          =   525
      Left            =   210
      TabIndex        =   1
      Top             =   90
      Width           =   5745
   End
End
Attribute VB_Name = "frmAnulla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadPX(PXs() As String)
    'cargar los prefijos para elegir cual matar
    Dim H As Long
    For H = 1 To UBound(PXs)
        If Len(PXs(H)) > 2 Then lstAnulla.AddItem PXs(H)
    Next H
    
    Me.Show 1
End Sub

Private Sub Command1_Click()
    Dim tmpPX() As String
    ReDim tmpPX(0)
    
    'marco los prefijos que no se van a usar
    Dim H As Long, I As Long
    I = 0
    For H = 0 To lstAnulla.ListCount - 1
        If lstAnulla.Selected(H) Then
            I = I + 1
            ReDim Preserve tmpPX(I)
            tmpPX(I) = lstAnulla.List(H)
        End If
    Next H
    
    frmPrincipal.SetNewPrefijos tmpPX
    
    Unload Me
End Sub
