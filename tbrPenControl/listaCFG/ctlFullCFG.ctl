VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlFullCFG 
   BackColor       =   &H00404040&
   ClientHeight    =   7980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   Begin VB.PictureBox CM1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7710
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   1110
      Width           =   555
   End
   Begin VB.PictureBox P3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   7380
      Left            =   180
      ScaleHeight     =   492
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   418
      TabIndex        =   2
      Top             =   150
      Width           =   6270
      Begin VB.PictureBox P2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5370
         Left            =   180
         ScaleHeight     =   358
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   377
         TabIndex        =   4
         Top             =   180
         Width           =   5655
         Begin tbrListaConfig_CTL.cltPROG PROG 
            Height          =   585
            Left            =   1050
            TabIndex        =   11
            Top             =   3990
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   1032
         End
         Begin VB.CommandButton btCa 
            Caption         =   "Salir"
            Height          =   315
            Left            =   1950
            TabIndex        =   7
            Top             =   4890
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.CommandButton btOk 
            Caption         =   "Ejecutar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   300
            TabIndex        =   10
            Top             =   4890
            Visible         =   0   'False
            Width           =   1575
         End
         Begin tbrListaConfig_CTL.ctlNumeroSimple lsNumeroSimple 
            Height          =   2955
            Left            =   3630
            TabIndex        =   9
            Top             =   330
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   5212
         End
         Begin tbrListaConfig_CTL.ctlTextoSimple lsTextoSimple 
            Height          =   960
            Left            =   300
            TabIndex        =   8
            Top             =   945
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   1693
         End
         Begin tbrListaConfig_CTL.ctlLIST lsCOMBO 
            Height          =   1515
            Left            =   30
            TabIndex        =   6
            Top             =   2070
            Visible         =   0   'False
            Width           =   3165
            _ExtentX        =   4842
            _ExtentY        =   1402
         End
         Begin VB.Label lbSoloInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "para mostrar los solo info"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   420
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   3075
         End
      End
      Begin VB.TextBox lbHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   915
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   4005
      End
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   1050
      Left            =   6600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1852
      _Version        =   393217
      LabelEdit       =   1
      HotTracking     =   -1  'True
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ctlFullCFG.ctx":0000
   End
End
Attribute VB_Name = "ctlFullCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim AnchoControl As Long
Dim AltoControl As Long

Dim SuperConfig As New clsConfig

Dim L As New clsElemList
Dim Ap As String
Dim EstoyEn As String
'a la izquierda en la lista de opciones es "tree"
'a la derecha en lo que sea es "rait"

Dim NodToSel As Node
Dim FSo As New Scripting.FileSystemObject
Dim ik As Long 'indice de el nodo elegido

Dim PathMusic As String 'path de la musica en el pendrive

'SEGUIRAQUI PONER ORIGENES
'Dim OrsPC As New tbrMPaquet.clsOrigDiscoManager   'todos los origenes del disco duro 'necesito los origenes oficiales de 3pm, estos so de prueba

Dim PcKy As String 'clave de esta pc(me lo mandan de afuera)
Dim mSoftNow As String 'NO LO USO pero quiero en los logs (es CV.JUAMAI)
Dim pthSysRockola As String 'ejecutable del softweare de fonola que estoy usando!
Dim mPERM As New clsPERM  'PERMISOS DEL PENDRIVE PARA CADA PC QUE ESTA HABILITADO

Public Event Fin(Rstrt As Long) 'me piden que termine
Public Event Ejecutar(sOrden As String) 'piden que se ejecute un comando
Public Event KeyDown2(KeyCode2 As Integer, Shift2 As Integer)
Public Event KeyUp2(KeyCode2 As Integer, Shift2 As Integer)

Private jpgFondo As String 'imagen de fondo que setea el programa que la llama
Private selConf As String 'marco de las config elegidas
Private PathImgConfs As String 'carpeta con los iconos de las configs

Private mUbicationBySys As Long
'mprock = 1
'e2games = 2
'genesis martino = 3
'SEGUIRAQUI y 3pm ?????????????

Public Property Let UbicatSys(mUbiSys As Long)
    mUbicationBySys = mUbiSys
End Property

Public Property Get UbicatSys() As Long
    UbicatSys = mUbicationBySys
End Property

Private Sub ShwSEL()
    On Local Error GoTo errCL
    
    HideAll
    
    terr.Anotar "qcb", ik
    
    terr.AppendSinHist "ShSEL:" + L.GetElement(ik).Caption
    
    Select Case L.GetElement(ik).eType
        Case SoloInfo
            'traducir si es texto de caption o de ayuda!! (para casos que venga con variables"
            shInfoP2 L.GetElement(ik).Help

        Case EjecutarProceso 'SOLO SI APRIETA ENTER se ejecuta, aqui solo se muestra!!!
            shInfoP2 L.GetElement(ik).Help
            
            'mostrar los botones "ejecutar y salir", si no no se entiende
            btOk.Left = P2.Width / 2 - btOk.Width / 2
            btOk.Top = P2.Height / 2 - btOk.Height
            
            btCa.Left = P2.Width / 2 - btCa.Width / 2
            btCa.Top = P2.Height / 2 + 6
            
            btOk.ZOrder: btCa.ZOrder
            btOk.Visible = True: btCa.Visible = True
            
            
        Case ListaCombo
            UbicatelsCOMBO
            terr.Anotar "qcf"
            'cargar los elemetos
            lsCOMBO.setManager L.GetElement(ik).Internal_ListaSImple
            terr.Anotar "qcg"
            
            'cosas raras del manu para que sea transparente
            lsCOMBO.ImitarFondo P3.hdc, P2.Left, P2.Top
            
            'y mostrarla
            lsCOMBO.LoadList
            terr.Anotar "qch"
            'y elegir la que se ha elegido
            lsCOMBO.SelElegida
            terr.Anotar "qci"
            
            'lsCOMBO.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsCOMBO.Visible = True
            lsCOMBO.ZOrder
            
            terr.Anotar "qcj"
        
        Case TextoSimple
            ubicateTextoSimple
            lsTextoSimple.setManager L.GetElement(ik).Internal_TextoSimple
            lsTextoSimple.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsTextoSimple.Visible = True
            
            lsTextoSimple.ImitarFondo P3.hdc, P2.Left, P2.Top
            
            lsTextoSimple.ZOrder
        
        Case SelectPath
            ubicateTextoSimple
            lsTextoSimple.setManager L.GetElement(ik).Internal_TextoSimple
            lsTextoSimple.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsTextoSimple.Visible = True
            
            lsTextoSimple.ImitarFondo P3.hdc, P2.Left, P2.Top
            
            lsTextoSimple.ZOrder
        
        Case Numero
            ubicateTextoSimple
            lsNumeroSimple.setManager L.GetElement(ik).Internal_Numeros
            lsNumeroSimple.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsNumeroSimple.Visible = True
            
            lsNumeroSimple.ImitarFondo P3.hdc, P2.Left, P2.Top
            
            lsNumeroSimple.ZOrder
            
    End Select
    
    terr.Anotar "qck"
    Exit Sub
    
errCL:
    terr.AppendLog "qce", terr.ErrToTXT(Err)
    Resume Next
End Sub

Public Property Let SoftNow(newS As String)
    mSoftNow = newS
End Property

Public Property Get SoftNow() As String
    SoftNow = mSoftNow
End Property

Public Sub setPathError(s As String)
    terr.FileLog = s
    terr.LargoAcumula = 1900
    terr.AppendSinHist "INI" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Sub
    
Public Function GetElemList() As clsElemList
    Set GetElemList = L
End Function

Private Sub shInfoP2(t As String, _
        Optional sBold As Boolean = False, _
        Optional sForeColor As Long = vbWhite)
        
    UbicateLbSoloInfo
    
    lbSoloInfo.Caption = t
    lbSoloInfo.Font.Bold = sBold
    lbSoloInfo.ForeColor = sForeColor
    lbSoloInfo.Visible = True
    
End Sub

Private Sub HideAll()
    lbSoloInfo.Visible = False
    lsCOMBO.Visible = False
    lsTextoSimple.Visible = False
    lsNumeroSimple.Visible = False
    PROG.Visible = False
    btOk.Visible = False
    btCa.Visible = False
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

Private Sub ubicateTextoSimple()
    lsTextoSimple.Top = 0
    lsTextoSimple.Left = 0
    lsTextoSimple.Width = P2.Width
    lsTextoSimple.Height = P2.Height
End Sub

Private Sub ubicateNumeroSimple()
    lsNumeroSimple.Top = 0
    lsNumeroSimple.Left = 0
    lsNumeroSimple.Width = P2.Width
    lsNumeroSimple.Height = P2.Height
End Sub


Public Sub Load(F As String) 'cargar un archivo de configuracion
    terr.Anotar "qad", F
    'definir los permisos!!!
    
    L.Load F, mPERM 'Ap + "ejemplo3PM.txt"
    
    terr.Anotar "qae", mPERM.toString
    
    L.LoadOnTreeView TV
    Set SuperConfig.Nodos = TV.Nodes
    IniciarConfig
    
    terr.Anotar "qaf"
End Sub

Public Sub wError(t As String, APPEN As Boolean)
    If APPEN Then
        terr.AppendLog t
    Else
        terr.Anotar "qaz", t
    End If

End Sub

Private Sub btCa_Click()
    GoCancel
End Sub

Private Sub btOk_Click()
    GoOK
End Sub

Private Sub CM1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown2(KeyCode, Shift)
End Sub

Private Sub CM1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp2(KeyCode, Shift)
End Sub

Private Sub lbHelp_GotFocus()
    CM1.SetFocus  'el foco de esto hace que las flechas de direccion en el form donde se carga este control
    'no reciba los eventos keyUp y down de las flechas de direccion
End Sub

Private Sub lsCOMBO_ClickCancel()
    GoCancel
End Sub

Private Sub lsCOMBO_ClickOK()
    GoOK
End Sub

Private Sub lsCOMBO_GotFocus()
    CM1.SetFocus
End Sub

Private Sub lsNumeroSimple_ClickCancel()
    GoCancel
End Sub

Private Sub lsNumeroSimple_ClickOK()
    GoOK
End Sub

Private Sub lsNumeroSimple_GotFocus()
    CM1.SetFocus
End Sub

Private Sub lsTextoSimple_ClickCancel()
    GoCancel
End Sub

Private Sub lsTextoSimple_ClickOK()
    GoOK
End Sub

Private Sub lsTextoSimple_GotFocus()
    CM1.SetFocus
End Sub

Private Sub P3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If EstoyEn = "rait" Then
        GoCancel
        'RestoreHDC_Rait
        'P3.Refresh
    End If
    
    SuperConfig.DoClick CLng(X), CLng(Y)
    If (SuperConfig.SelectedItem Is Nothing) Then Exit Sub
    
    upSEL
    GoOK
    
End Sub

Private Sub TV_GotFocus()
    CM1.SetFocus
End Sub

Private Sub UserControl_GotFocus()
    CM1.SetFocus
End Sub

Private Sub UserControl_Initialize()
    Ap = App.path
    If Right(Ap, 1) <> "\" Then Ap = Ap + "\"
    'Me.KeyPreview = True
    EstoyEn = "tree" 'treeview de opciones la izquierda
    
    lbHelp.Font = "Verdana"
    lbHelp.FontSize = 8
    lbHelp.FontBold = False
    lbHelp.ForeColor = RGB(80, 80, 80)
    lbHelp.BackColor = vbWhite
    
    lsCOMBO.Alignment = vbCenter
    terr.LargoAcumula = 700
    
    AnchoControl = 800
    AltoControl = 600
    
End Sub

Public Sub AddOrigMusica(nFol As String)
    terr.Anotar "AddOR:", nFol
    'SEGUIRAQUI DAR DE ALTA DE NUEVO ORSPC 18/9/09
    'OrsPC.AddOrig nFol
End Sub

Public Sub SetPcKy(nk As String)
    PcKy = nk
End Sub

Public Function GetPcKy() As String
    GetPcKy = PcKy
End Function

Public Function Ubicar()
    
End Function

Private Function NumeroProporcional(exMedida As Long, exMaxMedida, newMaxMedida As Long) As Long
    NumeroProporcional = newMaxMedida * (((exMedida * 100) / exMaxMedida) / 100)
End Function

Private Sub UserControl_Resize()
    
    'acomodar cosas internas
    UbicateLbSoloInfo
    UbicatelsCOMBO
    ubicateTextoSimple
    ubicateNumeroSimple
    
    'necesito q ue el TV no tenga el foco y lo pone !!!
    'necesito otro objeto que tome el foco
    CM1.Width = 1
    CM1.Height = 1
    CM1.Top = P3.Top
    CM1.Left = P3.Left
    
    'P.ZOrder
    P2.ZOrder
    lbHelp.ZOrder
End Sub

'*****************************************************************
'*****************************************************************
'TECLADO
'*****************************************************************

Public Sub GoRight()
    terr.Anotar "qcl", EstoyEn
    If EstoyEn = "tree" Then
        L.TV_Next
        terr.Anotar "qcn"
        'TV_Click
        
        SuperConfig.ComandoAdelante
        SuperConfig.Renderizar
        
        upSEL
        
        P3.Refresh
        UserControl.Refresh
    End If
    terr.Anotar "qcm"
    
    If EstoyEn = "rait" Then
        If L.GetElement(ik).eType = ListaCombo Then
            lsCOMBO.SelNext
        End If
        
        If L.GetElement(ik).eType = Numero Then
            lsNumeroSimple.SelNext
        End If
    End If
    
    
End Sub

Public Sub GoLeft()
    terr.Anotar "qco", EstoyEn
    If EstoyEn = "tree" Then
        L.TV_Prev
        terr.Anotar "qcp"
        'TV_Click
            
        SuperConfig.ComandoAtras
        SuperConfig.Renderizar
        
        upSEL
        
        P3.Refresh
        UserControl.Refresh
        
    End If
    terr.Anotar "qcq"
    If EstoyEn = "rait" Then
        If L.GetElement(ik).eType = ListaCombo Then
            lsCOMBO.SelPrev
        End If
        
        If L.GetElement(ik).eType = Numero Then
            lsNumeroSimple.SelPrev
        End If
    End If
    
End Sub

Private Sub upSEL()
    If SuperConfig.EstoyEnVolver = False Then
        Dim sp() As String
        sp = Split(SuperConfig.SelectedItem.Key)
        ik = CLng(sp(1)) 'el KEY es "NODO xx" siempre
        
        'los procesos los remarco
        If L.GetElement(ik).eType = EjecutarProceso Then
            lbHelp.BackColor = vbBlack
            lbHelp.ForeColor = vbWhite
            lbHelp.Font.Bold = True
        Else
            lbHelp.BackColor = vbWhite
            lbHelp.ForeColor = vbBlack
            lbHelp.Font.Bold = False
        End If
        
        lbHelp.Text = L.GetElement(ik).Help
        
    Else
        lbHelp.BackColor = vbWhite
        lbHelp.ForeColor = vbBlack
        lbHelp.Font.Bold = False
        
        lbHelp.Text = "Volver al menu anterior"
    End If
End Sub

Public Sub GoOK()

    On Local Error GoTo errOKGO
    'si no hay nada elegido salir de aca

    terr.Anotar "qcr", EstoyEn
    'terr.AppendSinHist "GOK2112." + EstoyEn + ":" + CStr(L.GetElement(ik).eType) + ":" + L.GetElement(ik).Internal_VerEXE.orden
    If EstoyEn = "tree" Then
        'si es un nodo que tiene hijos abrirlo (antes de que sea ejecutale. NO USAR EJECUTALES COMO PADRES!
        'Set NodToSel = TV.SelectedItem
        
        If SuperConfig.EstoyEnVolver Then
            'no hago nada
            'es una negrada, ya hace el manu mas abajo
            terr.AppendSinHist "VOLV-OK"
        Else
            terr.Anotar "qcr34"
            Set NodToSel = SuperConfig.SelectedItem
            If NodToSel Is Nothing Then Exit Sub
            
            If NodToSel.Children > 0 Then
                terr.Anotar "qcs"
                If NodToSel.Expanded Then
                    terr.AppendSinHist "Des-Expande"
                    NodToSel.Expanded = False
                Else
                    terr.Anotar "qct"
                    NodToSel.Expanded = True 'abro los hijos
                    NodToSel.Child.Selected = True 'eligo el primero hijo
                    terr.AppendSinHist "Expande-Sel_01"
                    'TV_Click
                End If
            Else
                terr.AppendSinHist "To-RAIT"
                ShwSEL
                terr.Anotar "qcu"
                EstoyEn = "rait"
                P2.Visible = True
                UserControl_Resize 'reacomoda
                terr.Anotar "qdb"
            End If
        End If
    Else 'tengo abierto la derecha EstoyEn = rait
        terr.Anotar "qdc"
        'volver al arbol
        EstoyEn = "tree"
        P2.Visible = False
        UserControl_Resize 'reacomoda
        terr.Anotar "qdd"
        'si estaba eligiendo opciones marcar la elegida
        If L.GetElement(ik).eType = ListaCombo Then
            terr.Anotar "qde"
            L.GetElement(ik).Internal_ListaSImple.ConfirmOption 'pasa el valor temportal a elegido
            'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
            
        End If
        
        'si estaba eligiendo TEXTO
        If L.GetElement(ik).eType = TextoSimple Then
            terr.Anotar "qde"
            'recien aqui asigno lo escrito a el valor definido
            L.GetElement(ik).Internal_TextoSimple.ConfirmOption
            'es lo mismo que:
            'L.GetElement(ik).Internal_TextoSimple.Valor = L.GetElement(ik).Internal_TextoSimple.ValorTMP
            
            'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
        End If
        
        'si estaba eligiendo TEXTO
        If L.GetElement(ik).eType = SelectPath Then
            terr.Anotar "qde"
            'recien aqui asigno lo escrito a el valor definido
            L.GetElement(ik).Internal_TextoSimple.ConfirmOption
            'es lo mismo que:
            'L.GetElement(ik).Internal_TextoSimple.Valor = L.GetElement(ik).Internal_TextoSimple.ValorTMP
            
            'NO ALTERAR EL CAPTION DEL ELEMENTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
        End If
        
        'si estaba eligiendo NUMERO
        If L.GetElement(ik).eType = Numero Then
            terr.Anotar "qde6"
            'recien aqui asigno lo escrito a el valor definido
            L.GetElement(ik).Internal_Numeros.ConfirmOption
            
            'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
        End If
        
        terr.Anotar "qdf"
        
        If L.GetElement(ik).eType = EjecutarProceso Then
            terr.Anotar "qcv"
            'mostrar el proceso
            EstoyEn = "rait"
            UserControl_Resize
            Dim keyNodo As String
            keyNodo = NodToSel.Key 'para saber si se elimino!
            terr.Anotar "qcw"
            Ejecutar L.GetElement(ik).Internal_VerEXE.orden
            
            'me quedo mostrando las opciones de lo que se ejecuto
            'EstoyEn = "tree"
            'UserControl_Resize
            
            'la ejecucion puede haber eliminado este nodo ...
            If L.ExisteNodoByKey(keyNodo) = False Then
                terr.Anotar "qcx"
                Exit Sub
            End If
            terr.Anotar "qcy"
            
            '... o desprendio hijos (cualquiera sea) me voy al primero
            If NodToSel.Children > 0 Then
                NodToSel.Expanded = True 'abro los hijos
                NodToSel.Child.Selected = True 'eligo el primero hijo
                terr.Anotar "qcz"
                'TV_Click
            End If
        End If
    End If
    
    SuperConfig.ComandoEntrar
    SuperConfig.Renderizar
    P3.Refresh
    UserControl.Refresh
    
    terr.Anotar "qdg"
    
    Exit Sub
    
errOKGO:
    terr.AppendLog "errOKGO", terr.ErrToTXT(Err)
End Sub

Public Sub GoCancel()
    terr.Anotar "qcr33"
    terr.AppendSinHist "GoCancel33." + EstoyEn + ":" + CStr(L.GetElement(ik).eType) + ":" + L.GetElement(ik).Internal_VerEXE.orden
    If EstoyEn = "tree" Then
        'deberia cerrar esta config ?
    Else
        terr.Anotar "qdc33"
        'volver al arbol
        EstoyEn = "tree"
        P2.Visible = False
        UserControl_Resize 'reacomoda
        terr.Anotar "qdd33"
        'si estaba eligiendo opciones asegurarse que no grabe la elegida
        If L.GetElement(ik).eType = ListaCombo Then
            terr.Anotar "qde33"
            
        End If
        
        'si estaba eligiendo TEXTO
        If L.GetElement(ik).eType = TextoSimple Then
            terr.Anotar "qde33"
            'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
            'version vieja larga
            '=L.GetElement(K).Caption + "=" + L.GetElement(K).Internal_ListaSImple.GetSelectOp
        End If
        
        'si estaba eligiendo TEXTO
        If L.GetElement(ik).eType = SelectPath Then
            terr.Anotar "qde33"
            'NO ALTERAR EL CAPTION DEL ELEMNTO QUE ES PURO
            L.GetElement(ik).NodeOp.Text = L.GetElement(ik).GetRes
            'version vieja larga
            '=L.GetElement(K).Caption + "=" + L.GetElement(K).Internal_ListaSImple.GetSelectOp
        End If
        
        terr.Anotar "qdf33"
    End If
    
    terr.Anotar "qdg33"
End Sub


'*****************************************************************
'*****************************************************************
'FIN TECLADO
'*****************************************************************

'/////////////////////////////////////////////////////////////
'SEGUIRAQUI
'en principio se hizo que se ejecute todo aqui dentro (para pencontrol) despues me di cuenta
'que le quita reusabilidad, las ordenes se deberian ejecutar afuera
'asi como para no hacerlo ahora agrego algun carcter especial cuando quiero que se ejecute afuera
Public Sub Ejecutar(orden As String)
    'ejecuta la orden
    terr.Anotar "qea", orden
    terr.AppendSinHist "Ex=" + orden
    
    If Left(orden, 1) = "*" Then
        Dim orden2 As String
        orden2 = mID(orden, 2)
        RaiseEvent Ejecutar(orden2)
    Else
        Select Case orden
            Case "listaNewMusicUSB"
                listMusicUSB 1 'desde aqui esta en TV por lo tanto si uso el TV !
            Case "listaMusicaSinUso"
                'seguiraqui ver si va o es automatico
                
            Case "UpdateMusic" 'se carga aqui (no en el archivo base) una vez que el tipo busca en el pendrive
                'hacer lo que tenga que hacer y eliminar el nodo "Actualizar!"
                UpdateMusic 1 'esta en el TV por eso lo uso
            Case "QuickUpdate"
                QuickUpdate
            Case "end" 'salir sin reiniciar la pc
                myEnd 0
            Case "endNOrestart" 'salir sin reiniciar la pc
                myEnd 1
            Case "endrestart" 'salir reiniciando la pc (para que inicie mprock)
                myEnd 2
            Case "reinicountr" 'reiniciar el contador reiniciable
                ReiniCountR
        End Select
    End If
    
End Sub

Private Sub QuickUpdate() 'el pendrive solo tiene permisos para actualziar musica
    terr.Anotar "qae11"
    listMusicUSB 0
    terr.Anotar "qae12"
    UpdateMusic 0
    terr.Anotar "qae13"
End Sub

Private Sub ReiniCountR()
    'reiniciar el contador reiniciable
    'SEGUIRAQUI
End Sub

Private Sub myEnd(Rstrt As Long)  'TERMINAR
    RaiseEvent Fin(Rstrt)
End Sub

Private Sub UpdateMusic(useTv As Long)
    'ya se definio la musica, ahora copiarla
    
    'hay un modo que actualiza la musica sin usar el TV (para pendrive automatizados y conm pocos permisos)
    terr.Anotar "qeb"
    PROG.Porc 1 / 15, "Calculando espacio necesario ..."
    ShowProg
    
    'ver que espacio disponible hay en las unidades de los origenes que se van a usar
    'comparar tamaño necesario de cada origen en el pendrive y compararlo con el espacio en el disco duro (unidad en que esta el origen de 3pm)
    'indicar que se desea copiar en cada lugar ...
    
    Dim OrigPD() As clsElem 'paquete con todos los origenes del pendrive
    OrigPD = L.getMarcados("OrigPD")
    terr.Anotar "qec"
    Dim K As Long
    For K = 1 To UBound(OrigPD)
        'ver si se eligio para copiar!!
        Dim dest As String
        dest = OrigPD(K).Internal_ListaSImple.GetSelectOpInternal
        If dest <> "" And LCase(dest) <> "no usar" Then
            'SEGUIRAQUI SEGUIR COMPLETANDO ORIGENES
            'Dim OrPcAUsar As tbrMPaquet.clsOrigDisco
            'Set OrPcAUsar = OrsPC.GetOrig(dest)
            OrPcAUsar.AddFolderToCopy OrigPD(K).InternalReal
            terr.Anotar "qed", OrigPD(K).InternalReal
        End If
        
        PROG.Porc (2 / 15) + ((K / UBound(OrigPD)) / 15), "Calculando espacio necesario " + OrigPD(K).Caption
    Next K
    terr.Anotar "qee"
    'todos esos datos dan un SI o NO final
    Dim UnidadPasada As String
    UnidadPasada = OrsPC.HayEspacioOK
    If UnidadPasada <> "" Then
        terr.Anotar "qef"
        PROG.Porc 1, "La copia completa planificada hara que el DISCO: " + UnidadPasada + "pase del 90% de uso" + _
            vbCrLf + "Por la integridad y correcto funcionamiento esta copia no se permitira. " + _
            "Modifique las opciones de copiado y reintente"
        Exit Sub
    End If
    terr.Anotar "qeg"
    'hay lugar ! comenzar a copiar
    If OrsPC.GetTotalOrs > 0 Then
        PROG.Porc 0.1, "Comezando copia " + FSo.GetBaseName(OrsPC.GetOrigByIndex(1).path)
        PROG.Refresh
    Else
        'IDIOTA no deberias haber entrado!!!
        'SOLUCIONAR
    End If
    terr.Anotar "qeh"
    'HACER UNA LISTA DE TODAS LAS CARPETAS QUE SE DEBEN COPIAR
    PROG.Porc 0.2, "Planificando copia"
    PROG.Refresh
    
    Dim aCopiar As String
    aCopiar = ""
    For K = 1 To OrsPC.GetTotalOrs
        terr.Anotar "qei", K
        aCopiar = aCopiar + OrsPC.GetOrigByIndex(K).GetListFolderToCopy
        'separa internamete chr5 y SIEMPRE termina en chr6
    Next K
    
    terr.Anotar "qel"
    Dim Lista() As String
    Lista = Split(aCopiar, Chr(6))
    Dim mORI As String, mDEST As String
    Dim L2() As String
    Dim St As Single
    
    'seguiraqui alcompilar el 18/9/09 jodia y lo saque
    'Dim MP As New tbrMPaquet.clsMPaquet 'para el 3h.dt
    
    terr.Anotar "qem"
    Dim COPIASHECHAS As Long
    COPIASHECHAS = 0
    For K = 0 To UBound(Lista)
        terr.Anotar "qen", K, Lista(K)
        L2 = Split(Lista(K), Chr(5))
        If UBound(L2) > 0 Then 'el ultimo de todos estara vacio
            
            'revisar 3H.DT (info del disco)
            Dim Orig3H As String
            Dim Dest3H As String
            
            If Right(L2(0), 1) <> "\" Then
                Orig3H = L2(0) + "\3H.DT"
            Else
                Orig3H = L2(0) + "3H.DT"
            End If
            
            If Right(L2(1), 1) <> "\" Then
                Dest3H = L2(1) + "\" + FSo.GetBaseName(L2(0)) + "\3H.DT"
            Else
                Dest3H = L2(1) + FSo.GetBaseName(L2(0)) + "\3H.DT"
            End If
            terr.Anotar "qeo", L2(0), L2(1)
            'si ya lo tiene el origen eliminarlo para que se cree puro
            If FSo.FileExists(Orig3H) Then FSo.DeleteFile Orig3H, True
            
            'LO COPIO y en el destino defino el 3H.DT
            '/////////////////////////////
            myCopyFolder L2(0), L2(1), True
            terr.Anotar "qep"
            terr.AppendSinHist "UpMSC:______" + FSo.GetBaseName(L2(0)) + ":____:" + FSo.GetBaseName(L2(1))
            '/////////////////////////////
            COPIASHECHAS = COPIASHECHAS + 1
            'DEFINIENDO o CORRIGIENDO 3H.DT
            'si ya estaba de antes es porque recopie un disco que ya existia!
            'si ya lo tiene el destino revisarlo y solo cambiarle el origen del disco
            terr.Anotar "qeq", COPIASHECHAS
            'modo de ingreso=pendrive! (para saber y para bloqueos tipo martino)
            If MP.CheckAndCreate(PcKy, L2(1), 2) = 0 Then
                'cuando da cero es que se cargo y existia ok para esta pc, como puede haber cambiado modoin
                terr.Anotar "qer"
                MP.Grabar False, L2(1), PcKy, 2
            End If
            
            terr.Anotar "qes"
            'MOSTRAR PROGRESO/////////////////////////////
            St = Round(K / UBound(Lista), 2)
            PROG.Porc St, "Copiando " + _
                            vbCrLf + "[" + FSo.GetBaseName(L2(0)) + "]" + _
                            vbCrLf + "Al origen [" + FSo.GetBaseName(FSo.GetParentFolderName(L2(1))) + "]"
            PROG.Refresh
            
        End If
        
        terr.Anotar "qet"
    Next K
    terr.Anotar "qeu", COPIASHECHAS
    'avisar que se termino
    PROG.Porc 1, "FINALIZADO" + vbCrLf + "La musica se ha actualziado correctamente" + vbCrLf + _
        "Se han cargado " + CStr(COPIASHECHAS) + " discos"
    PROG.Refresh
    terr.Anotar "qev"
    If useTv = 1 Then
        terr.Anotar "qew"
        'elimiar el nodo de actualizar musica para evitar que haga todo de nuevo!
        'quitar todos los origenes, el actualizar y marcar como usado el nodo de buscar
        Dim CE As clsElem
        Set CE = L.GetElementByMarca("findmusicusb")
        CE.NodeOp.BackColor = &HFFC0C0
        CE.NodeOp.ForeColor = vbWhite
        CE.Internal_VerEXE.orden = "" 'YA NO ESTA MAS LA ORDEN DE EJECUTAR!
        CE.NodeOp.Text = "ACTUALIZACION COMPLETADA" 'lo cambio por uno informativo con datos de lo que se hizo
        CE.eType = SoloInfo
        CE.Help = "La actualización se ha completado" 'SEGUIRAQUI estaria bueno un detalle de lo copiado, resumen final
        terr.Anotar "qex"
        'eliminar nodos pero tambien los elementos
        L.TV_KillNodeSel
        L.TV_KillMarcados "OrigPD"
        'no elimine los clsElem por que imagino que seria un problema
        'SOLUCIONAR (de todas formas no jode aparentemente)
        
        'decirle a IK cual es el elegido !
        'If Not (TV.SelectedItem Is Nothing) Then
        If Not (SuperConfig.SelectedItem Is Nothing) Then
            terr.Anotar "qey"
            Dim sp() As String
            'sp = Split(TV.SelectedItem.Key)
            sp = Split(SuperConfig.SelectedItem.Key)
            ik = CLng(sp(1)) 'el KEY es "NODO xx" siempre
        End If
    End If
    terr.Anotar "qez"
End Sub

Private Sub listMusicUSB(useTv As Long)
    On Local Error GoTo NOLISTA
    
    'es posible que desee cargar todo virtualmente a cElem sin cargarlo al treeview,e sto sirve para casos de PDs que no muestran nada y solo actualizan
    
    'EN CADA PROCESO PONER LA BARRA
    terr.Anotar "qba"
    PROG.Porc 1 / 15, "Listando contenido del pendrive"
    ShowProg
    
    If useTv = 1 Then
        'buscar en el pendrive, tratar de coincidir con los origenes de 3PM
        'mostrar cuales estan listos para copiarse (se encotraron igual) y cuales puede cambiar
        'If TV.SelectedItem Is Nothing Then Exit Sub
        If SuperConfig.SelectedItem Is Nothing Then Exit Sub
        
        'Set NodToSel = TV.SelectedItem 'supongo que hay uno elegido que despidio la orden!
        Set NodToSel = SuperConfig.SelectedItem 'supongo que hay uno elegido que despidio la orden!
    End If
    
    Dim IndicesNuevosNodos As Long
    IndicesNuevosNodos = 201
    'agregar cada uno de los origenes que tiene el pen
    Dim PD As Folder
    terr.Anotar "qbb", PathMusic
    Set PD = FSo.GetFolder(PathMusic)  'pathmusic me la da mprock leyendo el pendrive antes de iniciar el ejecutable de este pendrive
    PROG.Porc 2 / 15, "Obteniendo origenes de discos del sistema"
    
    Dim listaORG As String
    listaORG = "No usar|" + OrsPC.GetFullOrigString("|")
    
    Dim CadaOrigen As Folder, CNT As Long
    CNT = 0
    For Each CadaOrigen In PD.SubFolders
        terr.Anotar "qbc", CadaOrigen.Name, CNT
        'solo va a entrar una vez por que la prioridad en el ENTER es expandir si tiene hijos
        
        Dim newEL As clsElem
        Set newEL = L.addElement
        
        'ahora mismo ver a donde se asigna
        newEL.Internal_ListaSImple.LoadFromString listaORG, "|", "PATHS" 'necesito la lista cargada para elegir!
        Dim elegido As String
        elegido = newEL.Internal_ListaSImple.TryToSelectFromVisibleOptions(CadaOrigen.Name)
        terr.Anotar "qbd", elegido
        newEL.Caption = CadaOrigen.Name 'este elemento guarda el caption puro (= xxxx va aparte)
        newEL.InternalReal = CadaOrigen.path
        
        PROG.Porc (2 / 15) + 10 * ((CNT / PD.SubFolders.Count) / 15), "Encontrado: " + newEL.Caption 'el porcetaje extraño representa desde 2/15 hasta 12/15
        esperar 0.5
        newEL.PlusInfo = "Tamaño total: " + CStr(Round(((CadaOrigen.Size / 1024) / 1024), 2)) + " MB" + vbCrLf + _
                         "Discos: " + CStr(CadaOrigen.SubFolders.Count)
        
        terr.Anotar "qbe", newEL.PlusInfo
        
        newEL.Marca = "OrigPD" 'despues obtener una coleccion con todos los elementos con esta marca
        newEL.eType = ListaCombo
        newEL.Help = "Definir donde se copiara este contenido"
        newEL.id = IndicesNuevosNodos
        
        If useTv = 1 Then
            Dim sp() As String
            sp = Split(NodToSel.Key)
            newEL.Padre = sp(1) 'NODO nn es el key del padre
            terr.Anotar "qbf"
            Set newEL.NodeOp = TV.Nodes.Add(NodToSel.Key, tvwChild, "NODO " + CStr(IndicesNuevosNodos), newEL.Caption + "=" + elegido)
            'Set newEL.NodeOp = SuperConfig.Nodes.Add(NodToSel.Key, tvwChild, "NODO " + CStr(IndicesNuevosNodos), newEL.Caption + "=" + elegido)
            newEL.NodeOp.ForeColor = vbRed
        End If
        
        IndicesNuevosNodos = IndicesNuevosNodos + 1
SIG:
        CNT = CNT + 1
    Next
    terr.Anotar "qbg"
    '//////////////////////////////////////////////////////////////////////////////////
    'ahora agregar un nodo que sea "CARGAR MUSICA" como "tio" de estos origenes
    
    PROG.Porc (13 / 15), "Agregando... "
    
    Set newEL = L.addElement
    newEL.Caption = "Actualizar ahora"
    newEL.eType = EjecutarProceso
    newEL.Help = "Si ya ha buscado y elegido el destino de la musica puede comenzar aqui el proceso de carga de la musica"
    newEL.id = 200
    
    If useTv = 1 Then
        Dim n2 As Node
        Set n2 = TV.Nodes.Add(NodToSel.Parent.Key, tvwChild, "NODO 200", "Actualizar ahora")
        'Set n2 = SuperConfig.Nodes.Add(NodToSel.Parent.Key, tvwChild, "NODO 200", "Actualizar ahora")
        n2.ForeColor = vbWhite
        n2.Bold = True
        n2.BackColor = &HFFC0C0
        sp = Split(NodToSel.Parent.Key)
        newEL.Padre = sp(1) 'NODO nn es el key del padre
    End If
    
    newEL.Internal_VerEXE.orden = "UpdateMusic"
    '//////////////////////////////////////////////////////////////////////////////////
    terr.Anotar "qbh"
    PROG.Porc (15 / 15), "Terminado"
    PROG.Visible = False
    
    Exit Sub
NOLISTA:
    terr.AppendLog "NOLISTA541:" + CStr(useTv), terr.ErrToTXT(Err)
    MsgBox "Errores al listar cotenido multimedia (" + terr.GetLastLog + ")-" + CStr(useTv)
End Sub

Private Sub ShowProg() 'mostrar progreso
    PROG.Top = 0
    PROG.Left = 0
    PROG.Width = P2.Width
    PROG.Height = P2.Height
    PROG.Visible = True
    PROG.ZOrder
End Sub

Private Function getPathProg(likeProg As String, likeEmpresa As String) As String
    'ver donde esta 3PM para saber donde esta lo que necesito
    Dim MIP As New tbrInstalledPrograms.tbrProgsInst
    Dim PTH As String
    MIP.LoadList
    PTH = MIP.GetPath2(likeProg, likeEmpresa)
    'PTh = "D:\dev\3PM kundera 718047"
    If Right(PTH, 1) <> "\" Then PTH = PTH + "\"
    
    getPathProg = PTH
End Function

Public Sub SetPathMusic(s As String)
    terr.Anotar "SPthMu:", PathMusic
    PathMusic = s
End Sub

Public Sub SetPathSysRockola(s As String)
    terr.Anotar "SPthSysRoc:", pthSysRockola
    pthSysRockola = s
End Sub

Public Sub PlayRockola()
    terr.AppendSinHist "reiniSYS:" + pthSysRockola
    If FSo.FileExists(pthSysRockola) Then
        Shell pthSysRockola, vbMaximizedFocus
    Else
        terr.AppendLog "NO-X.34"
        MsgBox "No se ha encontrado el sistema " + vbCrLf + FSo.GetBaseName(pthSysRockola) + vbCrLf + _
        "Inicielo manualmente" + vbCrLf + "(" + terr.FileLog + ")"
    End If
End Sub

'entrega los permisos para leer y escribir
Public Function GetPerms() As clsPERM
    Set GetPerms = mPERM
End Function

Public Function VS() As clsVARS
    Set VS = VS2
End Function

Public Sub SetImgFondo(s As String)
    jpgFondo = s
End Sub

Public Sub SetSelConf(s As String)
    selConf = s
End Sub

Public Sub SetPathImgConfs(s As String)
    If Left(s, 1) <> "\" Then s = s + "\"
    PathImgConfs = s
End Sub

Private Sub IniciarConfig()
    Dim rctBotones As New clsRectangulo
    Dim rctLista As New clsRectangulo
    Dim rctTitulo As New clsRectangulo
    Dim rctTitulo2 As New clsRectangulo
    
    SuperConfig.IniciarNodos TV, "NODO 0"
    SuperConfig.IniciarFuente Parent, "Arial", 14, False, False, False, False, RGB(5, 5, 30), RGB(200, 200, 200)
    
    
    '==================================
    'Medidas Dentro de la configuracion
    '==================================
    'P3 es el contenedor de toda la configuracion dentro del formulario
    P3.Top = 15:         P3.Left = 15
    P3.Width = AnchoControl:         P3.Height = AltoControl
    
    'SE UBICA DIFERENTE SEGUN EL SISTEMA QUE LO LLAME
    'p2 es cada configuracion para modificar que va a la derecha segun elemento elegido en la lista de la izquierda
    If mUbicationBySys = 1 Then 'mprock
        P2.Left = NumeroProporcional(591, 800, AnchoControl)
        P2.Top = NumeroProporcional(133, 600, AltoControl)
        P2.Width = NumeroProporcional(197, 600, AltoControl)
        P2.Height = NumeroProporcional(269, 800, AnchoControl)
        
        lbHelp.Left = NumeroProporcional(144, 800, AnchoControl)
        lbHelp.Top = NumeroProporcional(411, 600, AltoControl)
        lbHelp.Width = NumeroProporcional(643, 800, AnchoControl)
        lbHelp.Height = NumeroProporcional(129, 600, AltoControl)
        
        rctBotones.X = NumeroProporcional(155, 800, AnchoControl)
        rctBotones.Y = NumeroProporcional(30, 600, AltoControl)
        rctBotones.Ancho = 0
        rctBotones.Alto = 0
    
        rctLista.X = NumeroProporcional(143, 800, AnchoControl)
        rctLista.Y = NumeroProporcional(137, 600, AltoControl)
        rctLista.Ancho = NumeroProporcional(438, 800, AnchoControl)
        rctLista.Alto = NumeroProporcional(262, 600, AltoControl)
    
        rctTitulo.X = NumeroProporcional(141, 800, AnchoControl)
        rctTitulo.Y = NumeroProporcional(103, 600, AltoControl)
        rctTitulo.Ancho = NumeroProporcional(442, 800, AnchoControl)
        rctTitulo.Alto = NumeroProporcional(27, 600, AltoControl)
    
        rctTitulo2.X = NumeroProporcional(144, 800, AnchoControl)
        rctTitulo2.Y = NumeroProporcional(411, 600, AltoControl)
        rctTitulo2.Ancho = NumeroProporcional(643, 800, AnchoControl)
        rctTitulo2.Alto = NumeroProporcional(129, 600, AltoControl)
    End If
    
    If mUbicationBySys = 2 Then 'e2games
        P2.Left = NumeroProporcional(389, 640, AnchoControl)
        P2.Top = NumeroProporcional(108, 480, AltoControl)
        P2.Height = NumeroProporcional(300, 640, AnchoControl)
        P2.Width = NumeroProporcional(200, 480, AltoControl)
        
        lbHelp.Left = NumeroProporcional(50, 640, AnchoControl)
        lbHelp.Top = NumeroProporcional(416, 480, AltoControl)
        lbHelp.Width = NumeroProporcional(540, 640, AnchoControl)
        lbHelp.Height = NumeroProporcional(54, 480, AltoControl)
        
        rctBotones.X = NumeroProporcional(55, 640, AnchoControl)
        rctBotones.Y = NumeroProporcional(30, 480, AltoControl)
        rctBotones.Ancho = 0
        rctBotones.Alto = 0
    
        rctLista.X = NumeroProporcional(56, 640, AnchoControl)
        rctLista.Y = NumeroProporcional(145, 480, AltoControl)
        rctLista.Ancho = NumeroProporcional(320, 640, AnchoControl)
        rctLista.Alto = NumeroProporcional(235, 480, AltoControl)
    
        rctTitulo.X = NumeroProporcional(55, 640, AnchoControl)
        rctTitulo.Y = NumeroProporcional(112, 480, AltoControl)
        rctTitulo.Ancho = NumeroProporcional(320, 640, AnchoControl)
        rctTitulo.Alto = NumeroProporcional(28, 480, AltoControl)
    
        rctTitulo2.X = NumeroProporcional(10, 640, AnchoControl)
        rctTitulo2.Y = NumeroProporcional(458, 480, AltoControl)
        rctTitulo2.Ancho = NumeroProporcional(625, 640, AnchoControl)
        rctTitulo2.Alto = NumeroProporcional(18, 480, AltoControl)
    End If
    
    If mUbicationBySys = 3 Then 'martino
        
        P2.Left = NumeroProporcional(389, 640, AnchoControl)
        P2.Top = NumeroProporcional(108, 480, AltoControl)
        P2.Height = NumeroProporcional(300, 640, AnchoControl)
        P2.Width = NumeroProporcional(200, 480, AltoControl)
        
        lbHelp.Left = NumeroProporcional(50, 640, AnchoControl)
        lbHelp.Top = NumeroProporcional(416, 480, AltoControl)
        lbHelp.Width = NumeroProporcional(540, 640, AnchoControl)
        lbHelp.Height = NumeroProporcional(54, 480, AltoControl)
        
        rctBotones.X = NumeroProporcional(55, 640, AnchoControl)
        rctBotones.Y = NumeroProporcional(30, 480, AltoControl)
        rctBotones.Ancho = 0
        rctBotones.Alto = 0
    
        rctLista.X = NumeroProporcional(56, 640, AnchoControl)
        rctLista.Y = NumeroProporcional(145, 480, AltoControl)
        rctLista.Ancho = NumeroProporcional(320, 640, AnchoControl)
        rctLista.Alto = NumeroProporcional(235, 480, AltoControl)
    
        rctTitulo.X = NumeroProporcional(55, 640, AnchoControl)
        rctTitulo.Y = NumeroProporcional(112, 480, AltoControl)
        rctTitulo.Ancho = NumeroProporcional(320, 640, AnchoControl)
        rctTitulo.Alto = NumeroProporcional(28, 480, AltoControl)
    
        rctTitulo2.X = NumeroProporcional(10, 640, AnchoControl)
        rctTitulo2.Y = NumeroProporcional(458, 480, AltoControl)
        rctTitulo2.Ancho = NumeroProporcional(625, 640, AnchoControl)
        rctTitulo2.Alto = NumeroProporcional(18, 480, AltoControl)
    End If
    
    If mUbicationBySys = 4 Then 'multitech
        P2.Left = NumeroProporcional(591, 800, AnchoControl)
        P2.Top = NumeroProporcional(133, 600, AltoControl)
        P2.Width = NumeroProporcional(197, 600, AltoControl)
        P2.Height = NumeroProporcional(269, 800, AnchoControl)
        
        lbHelp.Left = NumeroProporcional(144, 800, AnchoControl)
        lbHelp.Top = NumeroProporcional(411, 600, AltoControl)
        lbHelp.Width = NumeroProporcional(643, 800, AnchoControl)
        lbHelp.Height = NumeroProporcional(129, 600, AltoControl)
        
        rctBotones.X = NumeroProporcional(155, 800, AnchoControl)
        rctBotones.Y = NumeroProporcional(30, 600, AltoControl)
        rctBotones.Ancho = 0
        rctBotones.Alto = 0
    
        rctLista.X = NumeroProporcional(143, 800, AnchoControl)
        rctLista.Y = NumeroProporcional(137, 600, AltoControl)
        rctLista.Ancho = NumeroProporcional(438, 800, AnchoControl)
        rctLista.Alto = NumeroProporcional(262, 600, AltoControl)
    
        rctTitulo.X = NumeroProporcional(141, 800, AnchoControl)
        rctTitulo.Y = NumeroProporcional(103, 600, AltoControl)
        rctTitulo.Ancho = NumeroProporcional(442, 800, AnchoControl)
        rctTitulo.Alto = NumeroProporcional(27, 600, AltoControl)
    
        rctTitulo2.X = NumeroProporcional(144, 800, AnchoControl)
        rctTitulo2.Y = NumeroProporcional(411, 600, AltoControl)
        rctTitulo2.Ancho = NumeroProporcional(643, 800, AnchoControl)
        rctTitulo2.Alto = NumeroProporcional(129, 600, AltoControl)
    End If
    
    If mUbicationBySys = 5 Then 'kume
        P2.Left = NumeroProporcional(441, 800, AnchoControl)
        P2.Top = NumeroProporcional(133, 600, AltoControl)
        P2.Width = NumeroProporcional(347, 600, AltoControl)
        P2.Height = NumeroProporcional(269, 800, AnchoControl)
        
        lbHelp.Left = NumeroProporcional(144, 800, AnchoControl)
        lbHelp.Top = NumeroProporcional(411, 600, AltoControl)
        lbHelp.Width = NumeroProporcional(643, 800, AnchoControl)
        lbHelp.Height = NumeroProporcional(129, 600, AltoControl)
        
        rctBotones.X = NumeroProporcional(155, 800, AnchoControl)
        rctBotones.Y = NumeroProporcional(30, 600, AltoControl)
        rctBotones.Ancho = 0
        rctBotones.Alto = 0
    
        rctLista.X = NumeroProporcional(143, 800, AnchoControl)
        rctLista.Y = NumeroProporcional(137, 600, AltoControl)
        rctLista.Ancho = NumeroProporcional(290, 800, AnchoControl)
        rctLista.Alto = NumeroProporcional(262, 600, AltoControl)
    
        rctTitulo.X = NumeroProporcional(141, 800, AnchoControl)
        rctTitulo.Y = NumeroProporcional(103, 600, AltoControl)
        rctTitulo.Ancho = NumeroProporcional(290, 800, AnchoControl)
        rctTitulo.Alto = NumeroProporcional(27, 600, AltoControl)
    
        rctTitulo2.X = NumeroProporcional(144, 800, AnchoControl)
        rctTitulo2.Y = NumeroProporcional(411, 600, AltoControl)
        rctTitulo2.Ancho = NumeroProporcional(643, 800, AnchoControl)
        rctTitulo2.Alto = NumeroProporcional(129, 600, AltoControl)
    End If
        
    If EstoyEn = "tree" Then P2.Visible = False
    If EstoyEn = "rait" Then P2.Visible = True
    
    '---------------------------------------------------------------------
    
    SuperConfig.IniciarRectangulos rctBotones, rctLista, rctTitulo, rctTitulo2
    SuperConfig.IniciarGrafios P3.hdc, 0, 0, AnchoControl, AltoControl, jpgFondo, selConf, PathImgConfs      ', LabelMatrix(0)
    
    SuperConfig.CargarNodos
    
    SuperConfig.Renderizar
    P3.Refresh
    UserControl.Refresh
End Sub

'el manu necesita un formulario
Public Sub SetFRMNEGRADA(F As Object)
    Set NegradaFrmManu = F
End Sub

