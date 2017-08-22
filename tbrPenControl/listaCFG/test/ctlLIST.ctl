VERSION 5.00
Begin VB.UserControl ctlLIST 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lbTITULO 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1170
      TabIndex        =   1
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label lbList 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "lista 0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1290
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ctlLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mMargen As Long 'margen arriba abajo der izq
Private mMargenRenglones As Long

Private mForeColorNormal As Long
Private mForeColorSel As Long
Private mBackColorSel As Long
Private mBackColorFondo As Long

Private PosicionElegida As Long 'posicion del elemento elegido

Private mOpciones() As String
Private manageList As New tbrListaConfig.clsVERListaSimple

Public Event Change(NewSel As String)

Public Function setManager(mng As tbrListaConfig.clsVERListaSimple)
    Set manageList = mng
End Function

Public Function getManager() As tbrListaConfig.clsVERListaSimple
    Set getManager = manageList
End Function

Public Function LoadList()
    
    UnLoadLST 'limpiar
    PosicionElegida = -1
    
    Dim a As Long
    For a = 0 To UBound(manageList.GetStringListVisible)
        If a > 0 Then Load lbList(a)
        lbList(a).Caption = manageList.GetOpVisible(a)
    Next a
    
    ReAcomodar
    UpdateSel 'que despinte a todos
End Function

Public Sub SelNext()
    Select Case PosicionElegida
        Case -1
            PosicionElegida = 0
        Case lbList.Count - 1
            PosicionElegida = 0
        Case Else
            PosicionElegida = PosicionElegida + 1
    End Select
    UpdateSel 'mostrar el elegido
End Sub

Public Sub SelPrev()
    Select Case PosicionElegida
        Case -1
            PosicionElegida = lbList.Count - 1
        Case 0
            PosicionElegida = lbList.Count - 1
        Case Else
            PosicionElegida = PosicionElegida - 1
    End Select
    UpdateSel 'mostrar el elegido
End Sub

Public Sub SelElegida() 'marcar como elegida la que corresponde
    Dim a As Long
    PosicionElegida = 0 'valor predeterminado
    For a = 0 To lbList.Count - 1
        If lbList(a).Caption = manageList.GetSelectOp Then
            PosicionElegida = a
            Exit For
        End If
    Next a
    
    UpdateSel
End Sub

Private Sub UpdateSel()
    Dim a As Long
    For a = 0 To lbList.Count - 1
        If a = PosicionElegida Then
            lbList(a).BackStyle = 1 'opaco
            lbList(a).BackColor = mBackColorSel
            lbList(a).ForeColor = mForeColorSel
        Else
            lbList(a).BackStyle = 0 'transp
            lbList(a).ForeColor = mForeColorNormal
        End If
    Next a
    
    'que la clase se acuerde del elegido
    If PosicionElegida >= 0 Then
        manageList.TryToSelectFromVisibleOptions lbList(PosicionElegida)
        RaiseEvent Change(lbList(PosicionElegida))
    End If
    
End Sub

Private Sub UserControl_Initialize()
    mMargen = 60
    mMargenRenglones = 60
    
    Me.Alignment = vbCenter
    
    mBackColorFondo = 0
    mBackColorSel = &HFF0000
    mForeColorNormal = vbWhite
    mForeColorSel = &HFFFFC0
    
    lbTITULO.Top = 0
    lbTITULO.Left = 0
    lbTITULO.Width = UserControl.Width
    lbTITULO.Height = 630 'mas o menos 2 renglones
    
    lbList(0).Top = lbTITULO.Top + lbTITULO.Height + mMargen
    
End Sub

Private Sub UserControl_Resize()
    ReAcomodar
End Sub

Public Property Get Alignment() As AlignmentConstants
    Alignement = lbList(0).Alignment
End Property

Public Property Let Alignment(val As AlignmentConstants)
    lbList(0).Alignment = val
    ReAcomodar
End Property

Public Property Get Font() As StdFont
    Set Font = lbList(0).Font
End Property

Public Property Set Font(val As StdFont)
    Dim a As Long
    For a = 0 To lbList.Count - 1
        Set lbList(a).Font = val
    Next a
    
    Set lbTITULO.Font = Font
    
    ReAcomodar
End Property

Private Sub ReAcomodar()  'cuando cambia la fuente o algo se debe reacomodar

    lbTITULO.Width = UserControl.Width - margen * 2
    lbTITULO.Left = margen
    
    lbList(0).Top = lbTITULO.Top + lbTITULO.Height + mMargen
    
    Dim a As Long
    For a = 0 To lbList.Count - 1
        
        'alineacion
        If lbList(0).Alignment = 0 Then 'izq
            lbList(a).Left = mMargen
        End If
        
        If lbList(0).Alignment = 1 Then 'der
            lbList(a).Left = UserControl.Width - lbList(a).Width - mMargen
        End If
        
        If lbList(a).Alignment = 2 Then 'cent
            lbList(a).Left = UserControl.Width / 2 - lbList(a).Width / 2
        End If
        
        'lista
        If a > 0 Then
            lbList(a).Top = lbList(a - 1).Top + lbList(a).Height + mMargenRenglones
        End If
    
        lbList(a).Visible = True
    Next a

End Sub

Private Sub UnLoadLST()
    Dim a As Long
    For a = 1 To lbList.Count - 1
        Unload lbList(a)
    Next a
End Sub

Public Sub SetTitulo(t As String)
    lbTITULO.Caption = t
End Sub
