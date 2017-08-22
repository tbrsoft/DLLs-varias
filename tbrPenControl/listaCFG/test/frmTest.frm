VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Pruebas"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P2 
      BackColor       =   &H00000000&
      Height          =   4545
      Left            =   3210
      ScaleHeight     =   4485
      ScaleWidth      =   3405
      TabIndex        =   1
      Top             =   60
      Width           =   3465
      Begin Proyecto1.ctlLIST lsCOMBO 
         Height          =   1995
         Left            =   540
         TabIndex        =   3
         Top             =   1110
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3519
      End
      Begin VB.Label lbSoloInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "para mostrar los solo info"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   1635
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   2884
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
      Height          =   405
      Left            =   7200
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L As New clsElemList
Dim Ap As String
Dim EstoyEn As String
'a la izquierda en la lista de opciones es "tree"
'a la derecha en lo que sea es "rait"

Dim NodToSel As Node
Dim FSO As New Scripting.FileSystemObject
Dim K As Long 'indice de el nodo elegido

Private Sub Form_Load()
    Ap = App.Path
    If Right(Ap, 1) <> "\" Then Ap = Ap + "\"
    Me.KeyPreview = True
    EstoyEn = "tree" 'treeview de opciones la izquierda
    
    lsCOMBO.Alignment = vbCenter
    
    L.Load Ap + "ejemplo3PM.txt"
    L.LoadOnTreeView TV

End Sub

Private Sub Form_Resize()
    TV.Top = 120
    TV.Left = 60
    
    If EstoyEn = "tree" Then
        TV.Width = (Me.Width - 230)
        P2.Visible = False
    End If
    
    If EstoyEn = "rait" Then
        TV.Width = (Me.Width - 180) / 2 - 60
        P2.Visible = True
    End If
    
    TV.Height = Me.Height - 600
    
    P2.Left = TV.Width + TV.Left + 60
    P2.Top = TV.Top
    P2.Height = TV.Height
    P2.Width = TV.Width
    
    'acomodar cosas internas
    UbicateLbSoloInfo
    UbicatelsCOMBO
End Sub

Private Sub lsCOMBO_Change(NewSel As String)
    'Form1.Text1.Text = L.Status
End Sub

Private Sub TV_Click()
    'mostrar a la derecha lo que corresponda
    'limpiar todo primero
    HideAll
    
    'obtener el ID del elemento elegido
    Dim SP() As String
    SP = Split(TV.SelectedItem.Key)
    
    K = CLng(SP(1)) 'el KEY es "NODO xx" siempre
    

    Select Case L.GetElement(K).eType
        Case SoloInfo
            shSoloInfo L.GetElement(K).Help
            
        Case EjecutarProceso 'SOLO SI APRIETA ENTER se ejecuta, aqui solo se muestra!!!
            shSoloInfo L.GetElement(K).Help
            
        Case ListaCombo
            UbicatelsCOMBO
            'cargar los elemtos
            lsCOMBO.setManager L.GetElement(K).Internal_ListaSImple
            'y mostrarla
            lsCOMBO.LoadList
            'y elegir la que se ha elegido
            lsCOMBO.SelElegida
            
            lsCOMBO.Visible = True
            lsCOMBO.SetTitulo L.GetElement(K).Help
    End Select
    
    'Form1.Text1.Text = L.Status
End Sub

Private Sub shSoloInfo(t As String)
    UbicateLbSoloInfo
    lbSoloInfo.Caption = t
    lbSoloInfo.Visible = True
End Sub

Private Sub HideAll()
    lbSoloInfo.Visible = False
    lsCOMBO.Visible = False
End Sub
    
Private Sub UbicateLbSoloInfo()
    lbSoloInfo.Top = 0
    lbSoloInfo.Left = 0
    lbSoloInfo.Width = P2.Width
    lbSoloInfo.Height = P2.Height
End Sub

Private Sub UbicatelsCOMBO()
    lsCOMBO.Top = 0
    lsCOMBO.Left = 0
    lsCOMBO.Width = P2.Width
    lsCOMBO.Height = P2.Height
End Sub

Private Sub EJECUTAR(orden As String)
    Select Case orden
        Case "listaNewMusicUSB"
            listMusicUSB
        Case "listaMusicaSinUso"
        
        Case "UpdateMusic" 'se carga aqui (no en el archivo base) una vez que el tipo busca en el pendrive
            'hacer lo que tenga que hacer y eliminar el nodo "Actualizar!"
            'SEGUIR AQUI
    End Select
    
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyZ
            If EstoyEn = "tree" Then
                L.TV_Prev
                TV_Click
            End If
            
            If EstoyEn = "rait" Then lsCOMBO.SelPrev
            
        Case vbKeyX
            If EstoyEn = "tree" Then
                L.TV_Next
                TV_Click
            End If
            
            If EstoyEn = "rait" Then lsCOMBO.SelNext
            
        Case vbKeyReturn
            If EstoyEn = "tree" Then
                'si es un nodo que tiene hijos abrirlo (ates de que sea ejecutale. NO USAR EJECUTALES COMO PADRES!
                Set NodToSel = TV.SelectedItem
                
                If NodToSel.Children > 0 Then
                    If NodToSel.Expanded Then
                        NodToSel.Expanded = False
                    Else
                        NodToSel.Expanded = True 'abro los hijos
                        NodToSel.Child.Selected = True 'eligo el primero hijo
                        TV_Click
                    End If
                Else
                    If L.GetElement(K).eType = EjecutarProceso Then
                        EJECUTAR L.GetElement(K).Internal_VerEXE.orden
                        'si la ejecucion desprendio hijos (cualquiera sea) me voy al primero
                        If NodToSel.Children > 0 Then
                            NodToSel.Expanded = True 'abro los hijos
                            NodToSel.Child.Selected = True 'eligo el primero hijo
                            TV_Click
                        End If
                    Else
                        EstoyEn = "rait"
                        P2.Visible = True
                        Form_Resize 'reacomoda
                    End If
                    
                End If
            Else
                'volver al arbol
                EstoyEn = "tree"
                P2.Visible = False
                Form_Resize 'reacomoda
                
                'si estaba eligiendo opciones marcar la elegida
                If L.GetElement(K).eType = ListaCombo Then
                    'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
                    L.GetElement(K).NodeOp.Text = L.GetElement(K).GetRes
                    'version vieja larga
                    '=L.GetElement(K).Caption + "=" + L.GetElement(K).Internal_ListaSImple.GetSelectOp
                End If
                
            End If
    End Select
End Sub

Private Sub listMusicUSB()
    'buscar en el pendrive, tratar de coincidir con los origenes de 3PM
    'mostrar cuales estan listos para copiarse (se encotraron igual) y cuales puede cambiar
    Set NodToSel = TV.SelectedItem 'supongo que hay uno elegido que despidio la orden!
    Dim IndicesNuevosNodos As Long
    IndicesNuevosNodos = 201
    'agregar cada uno de los origenes que tiene el pen
    Dim PD As Folder
    Set PD = FSO.GetFolder(Ap + "simulPD") 'ESTA LINEA ME LA TIENEN QUE DAR POR OTRO LADO, por ahora la simulo
    
    'necesito los origenes oficiales de 3pm, estos so de prueba
    Dim ORG(4) As String
    ORG(0) = "D:\MM\MUSIC\rock"
    ORG(1) = "D:\MM\MUSIC\pop"
    ORG(2) = "D:\MM\MUSIC\latinos"
    ORG(3) = "D:\MM\MUSIC\argento"
    ORG(4) = "D:\MM\MUSIC\regaee"
    
    Dim listaORG As String
    listaORG = "No usar|" + ORG(0) + "|" + ORG(1) + "|" + ORG(2) + "|" + ORG(3) + "|" + ORG(4)
    
    Dim CadaOrigen As Folder
    For Each CadaOrigen In PD.SubFolders
        'solo va a entrar una vez por que la prioridad en el ENTER es expandir si tiene hijos
        
        Dim newEL As tbrListaConfig.clsElem
        Set newEL = L.addElement
        
        'ahora mismo ver a donde se asigna
        newEL.Internal_ListaSImple.LoadFromString listaORG, "|", "PATHS" 'necesito la lista cargada para elegir!
        Dim elegido As String
        elegido = newEL.Internal_ListaSImple.TryToSelectFromVisibleOptions(CadaOrigen.Name)
        newEL.Caption = CadaOrigen.Name 'este elemento guarda el caption puro (= xxxx va aparte)
        newEL.eType = ListaCombo
        newEL.Help = "Definir donde se copiara este contenido"
        newEL.id = IndicesNuevosNodos
        Dim SP() As String
        SP = Split(NodToSel.Key)
        newEL.Padre = SP(1) 'NODO nn es el key del padre
        
        Set newEL.NodeOp = TV.Nodes.Add(NodToSel.Key, tvwChild, "NODO " + CStr(IndicesNuevosNodos), newEL.Caption + "=" + elegido)
        
        IndicesNuevosNodos = IndicesNuevosNodos + 1
        'SEGUIR AQUI
        'ver que sea usabe y que cada uno tenga un elemento en L con las opciones = origenes validos de 3PM
SIG:
    Next
    
    '//////////////////////////////////////////////////////////////////////////////////
    'ahora agregar un nodo que sea "CARGAR MUSICA" como "tio" de estos origenes
    Dim N2 As Node
    Set N2 = TV.Nodes.Add(NodToSel.Parent.Key, tvwChild, "NODO 200", "Actualizar ahora")
    N2.ForeColor = vbWhite
    N2.Bold = True
    N2.BackColor = &HFFC0C0
    Set newEL = L.addElement
    newEL.Caption = "Actualizar ahora"
    newEL.eType = EjecutarProceso
    newEL.Help = "Si ya ha buscado y elegido el destino de la musica puede comenzar aqui el proceso de carga de la musica"
    newEL.id = 200
    SP = Split(NodToSel.Parent.Key)
    newEL.Padre = SP(1) 'NODO nn es el key del padre
    newEL.Internal_VerEXE.orden = "UpdateMusic"
    '//////////////////////////////////////////////////////////////////////////////////
    
    
End Sub
