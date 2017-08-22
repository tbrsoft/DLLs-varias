VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlFullCFG 
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   Begin VB.PictureBox P3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   7410
      Left            =   150
      ScaleHeight     =   494
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   678
      TabIndex        =   2
      Top             =   675
      Width           =   10170
      Begin VB.PictureBox P2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5370
         Left            =   4080
         ScaleHeight     =   358
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   377
         TabIndex        =   4
         Top             =   360
         Width           =   5655
         Begin tbrListaConfig_CTL.ctlNumeroSimple lsNumeroSimple 
            Height          =   2955
            Left            =   3570
            TabIndex        =   9
            Top             =   330
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   5212
         End
         Begin tbrListaConfig_CTL.ctlTextoSimple lsTextoSimple 
            Height          =   960
            Left            =   330
            TabIndex        =   8
            Top             =   945
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   1693
         End
         Begin tbrListaConfig_CTL.ctlLIST lsCOMBO 
            Height          =   1515
            Left            =   1230
            TabIndex        =   6
            Top             =   3195
            Visible         =   0   'False
            Width           =   3165
            _ExtentX        =   4842
            _ExtentY        =   1402
         End
         Begin tbrListaConfig_CTL.cltPROG PROG 
            Height          =   825
            Left            =   150
            TabIndex        =   7
            Top             =   2820
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   1455
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
            Height          =   495
            Left            =   420
            TabIndex        =   5
            Top             =   135
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
         Left            =   225
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3285
         Width           =   4005
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "ctlFullCFG.ctx":0000
      Top             =   480
      Width           =   1305
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   4320
      Left            =   405
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   7620
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
      MouseIcon       =   "ctlFullCFG.ctx":0006
   End
End
Attribute VB_Name = "ctlFullCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
''=============================================================================
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Dim HDC_bup As Long 'guardo en memoria una parte del fondo para
''hacer transparentes los ocx.
''=============================================================================

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
Dim OrsPC As New tbrMPaquet.clsOrigDiscoManager   'todos los origenes del disco duro 'necesito los origenes oficiales de 3pm, estos so de prueba
Dim PcKy As String 'clave de esta pc(me lo mandan de afuera)
Dim mSoftNow As String 'NO LO USO pero quiero en los logs (es CV.JUAMAI)
Dim pthSysRockola As String 'ejecutable del softweare de fonola que estoy usando!
Dim mPERM As New clsPERM  'PERMISOS DEL PENDRIVE PARA CADA PC QUE ESTA HABILITADO

Public Event Fin(Rstrt As Long) 'me piden que termine
Public Event Ejecutar(sOrden As String) 'piden que se ejecute un comando
Private jpgFondo As String 'imagen de fondo que setea el programa que la llama
Private selConf As String 'marco de las config elegidas
Private PathImgConfs As String 'carpeta con los iconos de las configs

Private Sub ShwSEL()
    On Local Error GoTo errCL
    
    HideAll
    
    terr.Anotar "qcb", ik
    
    terr.AppendSinHist "ShSEL:" + L.GetElement(ik).Caption
    
    Select Case L.GetElement(ik).eType
        Case SoloInfo
            'traducir si es texto de caption o de ayuda!! (para casos que venga con variables"
            shSoloInfo L.GetElement(ik).Help

        Case EjecutarProceso 'SOLO SI APRIETA ENTER se ejecuta, aqui solo se muestra!!!
            shSoloInfo L.GetElement(ik).Help
            
        Case ListaCombo
            UbicatelsCOMBO
            terr.Anotar "qcf"
            'cargar los elemetos
            lsCOMBO.setManager L.GetElement(ik).Internal_ListaSImple
            terr.Anotar "qcg"
            
            'cosas raras del manu para que sea transparente
            lsCOMBO.ImitarFondo P3.HDC, P2.Left, P2.Top
            
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
            
            lsTextoSimple.ImitarFondo P3.HDC, P2.Left, P2.Top
            
            lsTextoSimple.ZOrder
        
        Case SelectPath
            ubicateTextoSimple
            lsTextoSimple.setManager L.GetElement(ik).Internal_TextoSimple
            lsTextoSimple.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsTextoSimple.Visible = True
            
            lsTextoSimple.ImitarFondo P3.HDC, P2.Left, P2.Top
            
            lsTextoSimple.ZOrder
        
        Case Numero
            ubicateTextoSimple
            lsNumeroSimple.setManager L.GetElement(ik).Internal_Numeros
            lsNumeroSimple.SetTitulo L.GetElement(ik).Help + vbCrLf + L.GetElement(ik).PlusInfo
            lsNumeroSimple.Visible = True
            
            lsNumeroSimple.ImitarFondo P3.HDC, P2.Left, P2.Top
            
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

Private Sub shSoloInfo(t As String)
    UbicateLbSoloInfo
    lbSoloInfo.Caption = t
    'lbSoloInfo.Visible = True
    
    lsCOMBO.Visible = False
    lsTextoSimple.Visible = False
    lsNumeroSimple.Visible = False
    PROG.Visible = False
    
End Sub

Private Sub HideAll()
    lbSoloInfo.Visible = False
    lsCOMBO.Visible = False
    lsTextoSimple.Visible = False
    lsNumeroSimple.Visible = False
    PROG.Visible = False
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

'Private Sub Command1_Click()
'    IniHDC HDC_bup, P2.Width, P2.Height
'    'BitBlt HDC_bup, 0, 0, P2.Width, P2.Height, P3.HDC, P2.Left, P2.Top, SRCCOPY
'    BitBlt HDC_bup, 0, 0, P2.Width, P2.Height, UserControl.HDC, 0, 0, SRCCOPY
'End Sub

Private Sub lsCOMBO_ClickCancel()
    GoCancel
End Sub

Private Sub lsCOMBO_ClickOK()
    GoOK
End Sub

Private Sub lsNumeroSimple_ClickCancel()
    GoCancel
End Sub

Private Sub lsNumeroSimple_ClickOK()
    GoOK
End Sub

Private Sub lsTextoSimple_ClickCancel()
    GoCancel
End Sub

Private Sub lsTextoSimple_ClickOK()
    GoOK
End Sub

Private Sub P3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If EstoyEn = "rait" Then
        GoCancel
        'RestoreHDC_Rait
        'P3.Refresh
    End If
    
    SuperConfig.DoClick CLng(x), CLng(y)
    If (SuperConfig.SelectedItem Is Nothing) Then Exit Sub
    
    upSEL
    GoOK
    
End Sub

Private Sub TV_GotFocus()
    Text1.SetFocus 'QUE NUNCA TENGA EL FOCO!
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
End Sub

Public Sub AddOrigMusica(nFol As String)
    terr.Anotar "AddOR:", nFol
    OrsPC.AddOrig nFol
End Sub

Public Sub SetPcKy(nk As String)
    PcKy = nk
End Sub

Public Function GetPcKy() As String
    GetPcKy = PcKy
End Function

Private Sub UserControl_Resize()
    P3.Top = 0
    P3.Left = 0
    
    'P3.Width = UserControl.Width - 30
    'P3.Height = UserControl.Height - 30
    P3.Width = 640
    P3.Height = 480
    
    If EstoyEn = "tree" Then P2.Visible = False
    If EstoyEn = "rait" Then P2.Visible = True
    
    lbHelp.Height = 55
    lbHelp.Left = 4
    lbHelp.Top = P3.Height - lbHelp.Height - 4
    lbHelp.Width = P3.Width - 8
    
    P2.Left = 389
    P2.Top = 108
    P2.Height = 300
    P2.Width = 200
    
    'acomodar cosas internas
    UbicateLbSoloInfo
    UbicatelsCOMBO
    ubicateTextoSimple
    ubicateNumeroSimple
    
    'necesito q ue el TV no tenga el foco y lo pone !!!
    'necesito otro objeto que tome el foco
    Text1.Top = P3.Top
    Text1.Left = P3.Left
    
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
        
        lbHelp.Text = L.GetElement(ik).Help
    Else
        lbHelp.Text = "Volver al menu anterior"
    End If
End Sub

Public Sub GoOK()

    'si no hay nada elegido salir de aca

    terr.Anotar "qcr"
    'terr.AppendSinHist "GOK2112." + EstoyEn + ":" + CStr(L.GetElement(ik).eType) + ":" + L.GetElement(ik).Internal_VerEXE.orden
    If EstoyEn = "tree" Then
        'si es un nodo que tiene hijos abrirlo (antes de que sea ejecutale. NO USAR EJECUTALES COMO PADRES!
        'Set NodToSel = TV.SelectedItem
        
        If SuperConfig.EstoyEnVolver Then
            'no hago nada
            'es una negrada, ya hace el manu mas abajo
            terr.AppendSinHist "VOLV-OK"
        Else
            Set NodToSel = SuperConfig.SelectedItem
            
            
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
            'version vieja larga
            '=L.GetElement(K).Caption + "=" + L.GetElement(K).Internal_ListaSImple.GetSelectOp
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

Private Sub Text1_Change()
    Text1.Text = "" 'para que no se llene
End Sub


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
            Dim OrPcAUsar As tbrMPaquet.clsOrigDisco
            Set OrPcAUsar = OrsPC.GetOrig(dest)
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
    
    Dim MP As New tbrMPaquet.clsMPaquet 'para el 3h.dt
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
    SuperConfig.IniciarNodos TV, "NODO 0"
    SuperConfig.IniciarFuente Parent, "Arial", 12, False, False, False, False, vbWhite, RGB(20, 20, 20)
    'SuperConfig.IniciarGrafios P.hdc, 0, 0, App.path + "\testPulenta\clsConfig\CONF.jpg", App.path + "\testPulenta\PNG\Selected.png", App.path + "\testPulenta\PNG\", LabelMatrix(0)
    SuperConfig.IniciarGrafios P3.HDC, 0, 0, jpgFondo, selConf, PathImgConfs ', LabelMatrix(0)
    'SuperConfig.IniciarGrafios Parent.p1.hdc, 0, 0, App.path + "\testPulenta\clsConfig\CONF.jpg", App.path + "\testPulenta\PNG\Selected.png", App.path + "\testPulenta\PNG\", LabelMatrix(0)
    
    SuperConfig.CargarNodos
    
    SuperConfig.Renderizar
    P3.Refresh
    UserControl.Refresh
End Sub

'el manu necesita un formulario
Public Sub SetFRMNEGRADA(F As Object)
    Set NegradaFrmManu = F
End Sub

'Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
'    Dim TempBMP As Long
'    Dim ObjCreado2 As Long
'
'    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
'    CualHDC = CreateCompatibleDC(0)
'    ObjCreado2 = SelectObject(CualHDC, TempBMP)
'
'    DeleteObject TempBMP
'    DeleteObject ObjCreado2
'End Sub
