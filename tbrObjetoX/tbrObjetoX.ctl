VERSION 5.00
Begin VB.UserControl tbrObjetoX 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   135
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   165
      Width           =   6975
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "ObjetoX"
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   705
         Visible         =   0   'False
         Width           =   6285
      End
      Begin VB.Shape Shap 
         Height          =   300
         Left            =   465
         Top             =   15
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H0009FFDA&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "tbrObjetoX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event SeleccionaItem(IndexItem As Long)


Dim SeHizoDrag As Boolean
Private oldRecta As RECT
Private newRecta As RECT

Dim TipoResize As Long
Dim lblXSeleccionado As Long

Dim rX As Long
Dim rY As Long
    
Private Sub AcomodarShap(qRecta As RECT)
    Shap.Left = qRecta.qLeft
    Shap.Top = qRecta.qTop
    Shap.Width = qRecta.qRight
    Shap.Height = qRecta.qBottom
End Sub

Private Sub RestaurarRectas()
    newRecta.qLeft = oldRecta.qLeft
    newRecta.qTop = oldRecta.qTop
    newRecta.qRight = oldRecta.qRight
    newRecta.qBottom = oldRecta.qBottom
End Sub

Private Sub RestaurarLabelsDesaparecidos()
    'por si un objeto se arrastra muy lejos y desaparece!
    If SeHizoDrag = True Then
        lblXVisibles
    End If
End Sub

Private Sub lblX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RestaurarLabelsDesaparecidos
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RestaurarLabelsDesaparecidos
End Sub

Private Sub lblX_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblXSeleccionado = Index
    LabelsVisibles False
    rX = X / 15
    rY = Y / 15
    SeHizoDrag = True
    lblX(Index).Visible = False
    lblX(Index).Drag 1
End Sub

Private Sub Picture1_Click()
    LabelsVisibles False
    lblXSeleccionado = -1
End Sub

'=============================================================
'CABMIAR MEDIDAS
'=============================================================
Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    Dim lblXix As Integer
    
    SeHizoDrag = False
    Select Case Source.Name
        Case "lblX"
            Source.Left = X - rX
            Source.Top = Y - rY
            Source.Visible = True
            Source.Drag 0
            AcomodarLabels Source
            LabelsVisibles True
            
            RaiseEvent SeleccionaItem(Source.Index)
        Case "lbl"
            Shap.Visible = False
            lblXix = CInt(lbl(Index).Tag)
            
            lblX(lblXix).Width = 200
            
            lblX(lblXix).Visible = True
            
            lblX(lblXix).Left = newRecta.qLeft
            lblX(lblXix).Top = newRecta.qTop
            lblX(lblXix).Width = newRecta.qRight
            lblX(lblXix).Height = newRecta.qBottom
            
            AcomodarLabels lblX(lblXix)
            HabilitarLblXTodos
            LabelsVisibles True
            
            RaiseEvent SeleccionaItem(CLng(lblXix))
    End Select
    
End Sub
Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim minMedida As Long
    Dim lblXix As Integer
    
    minMedida = 10
    
    Select Case Source.Name
        Case "lbl"
            lblXix = CInt(lbl(Source.Index).Tag)
            
            Select Case TipoResize
                Case 0
                    newRecta.qLeft = X
                    newRecta.qTop = Y
                    newRecta.qRight = oldRecta.qRight - (X - oldRecta.qLeft)
                    newRecta.qBottom = oldRecta.qBottom - (Y - oldRecta.qTop)
                Case 1
                    newRecta.qLeft = oldRecta.qLeft
                    newRecta.qTop = Y
                    newRecta.qRight = oldRecta.qRight
                    newRecta.qBottom = oldRecta.qBottom - (Y - oldRecta.qTop)
                Case 2
                    newRecta.qLeft = oldRecta.qLeft
                    newRecta.qTop = Y
                    newRecta.qRight = X - oldRecta.qLeft
                    newRecta.qBottom = oldRecta.qBottom - (Y - oldRecta.qTop)
                Case 3
                    newRecta.qLeft = X
                    newRecta.qTop = oldRecta.qTop
                    newRecta.qRight = oldRecta.qRight - (X - oldRecta.qLeft)
                    newRecta.qBottom = oldRecta.qBottom
                Case 4
                    'NO SE USA
                Case 5
                    newRecta.qLeft = oldRecta.qLeft
                    newRecta.qTop = oldRecta.qTop
                    newRecta.qRight = X - oldRecta.qLeft
                    newRecta.qBottom = oldRecta.qBottom
                Case 6
                    newRecta.qLeft = X
                    newRecta.qTop = oldRecta.qTop
                    newRecta.qRight = oldRecta.qRight - (X - oldRecta.qLeft)
                    newRecta.qBottom = Y - oldRecta.qTop
                Case 7
                    newRecta.qLeft = oldRecta.qLeft
                    newRecta.qTop = oldRecta.qTop
                    newRecta.qRight = oldRecta.qRight
                    newRecta.qBottom = Y - oldRecta.qTop
                Case 8
                    newRecta.qLeft = oldRecta.qLeft
                    newRecta.qTop = oldRecta.qTop
                    newRecta.qRight = X - oldRecta.qLeft
                    newRecta.qBottom = Y - oldRecta.qTop
            End Select
            
            If newRecta.qRight < minMedida Then
                RestaurarRectas
            End If
            If newRecta.qBottom < minMedida Then
                RestaurarRectas
            End If
            AcomodarShap newRecta
            
    End Select
End Sub

'--------------------------------
Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lblXix As Integer
    lblXix = CInt(lbl(Index).Tag)
    
    'como estaba el label cuando se empezo a editar el tamaño
    oldRecta.qLeft = lblX(lblXix).Left
    oldRecta.qTop = lblX(lblXix).Top
    oldRecta.qRight = lblX(lblXix).Width
    oldRecta.qBottom = lblX(lblXix).Height
    
    Shap.Visible = True
    TipoResize = CLng(Index)
    HabilitarSoloLblX CLng(lblXix)
    lbl(Index).Drag 1
    
    lblX(lblXix).Visible = False
    LabelsVisibles False
End Sub
Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lblXix As Integer
    lblXix = CInt(lbl(Index).Tag)
    
    lblX(lblXix).Width = 200
    
    lblX(lblXix).Visible = True
    AcomodarLabels lblX(lblXix)
    LabelsVisibles True
End Sub


'//
'//
'//
'//
'//
'//
'===============================================
'DATOS/Comienzo
'===============================================
Public Function GetObjetoXSeleccionadoIndex() As Long
    GetObjetoXSeleccionadoIndex = lblXSeleccionado
End Function

Public Sub AgregarObjetoX()
    Dim nwIx As Long
    nwIx = lblX.UBound + 1
    Load lblX(nwIx)
    lblX(nwIx).Left = 20
    lblX(nwIx).Top = 20
    lblX(nwIx).Width = 150
    lblX(nwIx).Height = 50
    lblX(nwIx).Visible = True
End Sub

Public Sub QuitarObjetoX(Index As Long)
    On Local Error GoTo XnoEncontrado1
    
    If Index < 1 Then Exit Sub
    If Index > lblX.UBound Then Exit Sub
    Unload lblX(Index)
    lblXSeleccionado = -1
    LabelsVisibles False
    
    Exit Sub
XnoEncontrado1:
    MsgBox "Error ObjetoX 001, el Index: " + CStr(Index) + " no existe.", vbCritical, "ObjetoX"
End Sub

Public Function GetObjetoX(Index As Long) As Object
    On Local Error GoTo XnoEncontrado4
    
    If Index < 1 Then Exit Function
    If Index > lblX.UBound Then Exit Function
    
    '------------------------------------
    Set GetObjetoX = lblX(Index)
    '------------------------------------
    
    Exit Function
XnoEncontrado4:
    MsgBox "Error ObjetoX 001, el Index: " + CStr(Index) + " no existe.", vbCritical, "ObjetoX"
End Function

Public Function SetObjetoX(Index As Long, lblRef As Object) As Object
    On Local Error GoTo XnoEncontrado4
    
    If Index < 1 Then Exit Function
    If Index > lblX.UBound Then Exit Function
    
    '------------------------------------
    lblX(Index).Left = lblRef.Left
    lblX(Index).Top = lblRef.Top
    lblX(Index).Width = lblRef.Width
    lblX(Index).Height = lblRef.Height
    '------------------------------------
    LabelsVisibles False
    
    Exit Function
XnoEncontrado4:
    MsgBox "Error ObjetoX 001, el Index: " + CStr(Index) + " no existe.", vbCritical, "ObjetoX"
End Function


Public Function GetPropiedadesObjetoX(Index As Long) As String
    On Local Error GoTo XnoEncontrado2
    
    If Index < 1 Then Exit Function
    If Index > lblX.UBound Then Exit Function
    
    '------------------------------------
    GetPropiedadesObjetoX = lblX(Index).Tag
    
    '------------------------------------
    
    Exit Function
XnoEncontrado2:
    MsgBox "Error ObjetoX 001, el Index: " + CStr(Index) + " no existe.", vbCritical, "ObjetoX"
End Function

Public Sub SetPropiedadesObjetoX(Index As Long, Propiedades As String)
    On Local Error GoTo XnoEncontrado3
    
    If Index < 1 Then Exit Sub
    If Index > lblX.UBound Then Exit Sub
    
    '------------------------------------
    lblX(Index).Tag = Propiedades
    '------------------------------------
    
    Exit Sub
XnoEncontrado3:
    MsgBox "Error ObjetoX 001, el Index: " + CStr(Index) + " no existe.", vbCritical, "ObjetoX"
End Sub
'===============================================
'DATOS/Fin
'===============================================


Private Sub UserControl_Initialize()
    CrearLabels
    LabelsVisibles False
    lblXSeleccionado = -1
End Sub

Private Sub UserControl_Resize()
    Picture1.Left = 0
    Picture1.Top = 0
    Picture1.Width = (UserControl.Width / 15)
    Picture1.Height = (UserControl.Height / 15)
End Sub


'//
'//
'//
'//
'//
'//
'============================================================
'modLabels.cls/Comienza
'============================================================
Private Sub CrearLabels()
    Dim i As Long
    For i = 1 To 8
        Load lbl(i)
        
    Next i
    
    For i = 0 To 8
        lbl(i).Visible = True
        lbl(i).Width = 8
        lbl(i).Height = 8
    Next i
    
    lbl(0).MousePointer = 8
    lbl(1).MousePointer = 7
    lbl(2).MousePointer = 6
    
    lbl(3).MousePointer = 9
    lbl(4).MousePointer = 5
    lbl(5).MousePointer = 9
    
    lbl(6).MousePointer = 6
    lbl(7).MousePointer = 7
    lbl(8).MousePointer = 8
    
    lbl(4).Visible = False
End Sub

Private Sub AcomodarLabels(EnDonde As Label)
    Dim i As Long
    
    
    'Arriba Izquierda
    lbl(0).Left = EnDonde.Left - lbl(1).Width
    'Arriba Medio
    lbl(1).Left = EnDonde.Left + (((EnDonde.Width) / 2) - (lbl(1).Width / 2))
    'Arriba Derecha
    lbl(2).Left = EnDonde.Left + ((EnDonde.Width)) - (lbl(2).Width) + lbl(1).Width
    
    lbl(0).Top = EnDonde.Top - lbl(1).Height
    lbl(1).Top = lbl(0).Top
    lbl(2).Top = lbl(0).Top
    '=====================================================
    'Arriba Izquierda
    lbl(3).Left = lbl(0).Left
    'Arriba Medio
    lbl(4).Left = lbl(1).Left
    'Arriba Derecha
    lbl(5).Left = lbl(2).Left
    
    lbl(3).Top = EnDonde.Top + ((EnDonde.Height) / 2) - (lbl(0).Height / 2)
    lbl(4).Top = lbl(3).Top
    lbl(5).Top = lbl(3).Top
    '=====================================================

    'Arriba Izquierda
    lbl(6).Left = lbl(0).Left
    'Arriba Medio
    lbl(7).Left = lbl(1).Left
    'Arriba Derecha
    lbl(8).Left = lbl(2).Left
    
    lbl(6).Top = EnDonde.Top + ((EnDonde.Height)) ' - (lbl(0).Height)
    lbl(7).Top = lbl(6).Top
    lbl(8).Top = lbl(6).Top
    
    For i = 0 To 8
        lbl(i).Tag = EnDonde.Index
        lbl(i).ZOrder
    Next i
End Sub

Private Sub LabelsVisibles(esVisible As Boolean)
    For i = 0 To 8
        lbl(i).Visible = esVisible
    Next i
    
    lbl(4).Visible = False
End Sub
'============================================================
'modLabels.cls/Fin
'============================================================


'//
'//
'//
'//
'//
'//
'============================================================
'modVarias.cls/Comienza
'============================================================
'Aqui hay 3 'On Local Error Resume Next'
'esta por el caso de que se haya borrado
'un ObjetoX y de error al hacer un for
'Atte: Manu-

Private Sub HabilitarSoloLblX(Index As Long)
    On Local Error Resume Next
    
    Dim i As Long
    For i = 1 To (lblX.UBound)
        If i <> Index Then
            lblX(i).Enabled = False
        Else
            lblX(i).Enabled = True
        End If
    Next i
End Sub

Private Sub HabilitarLblXTodos()
    On Local Error Resume Next
    
    Dim i As Long
    For i = 1 To (lblX.UBound)
        lblX(i).Enabled = True
    Next i
End Sub

Private Sub lblXVisibles()
    On Local Error Resume Next
    
    Dim i As Long
    For i = 1 To (lblX.UBound)
        lblX(i).Visible = True
    Next i
End Sub

'============================================================
'modVarias.cls/Fin
'============================================================

Public Function GetBackPic() As Object
    Set GetBackPic = Picture1
End Function

