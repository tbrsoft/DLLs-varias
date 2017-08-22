VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPrincipal 
   Caption         =   "Traductor de Proyectos"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sTab 
      Height          =   6495
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Controles"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdTraducir"
      Tab(0).Control(1)=   "lvwControles"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Codigo"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lvw"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdInsertar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdInsertar 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HideSelection   =   0   'False
         Left            =   120
         TabIndex        =   5
         Top             =   6000
         Width           =   5775
      End
      Begin VB.CommandButton cmdTraducir 
         Caption         =   "Traducir"
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         Top             =   6000
         Width           =   6615
      End
      Begin MSComctlLib.ListView lvwControles 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   2
         ToolTipText     =   "Seleccione las propiedades que deben ser traducidas."
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Propiedad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Lo que se modifique aqui se modificara en el codigo fuente!"
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9763
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11456
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuInsertarTraductor 
         Caption         =   "Insertar Traductor"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vbinstance As VBIDE.VBE

Dim vCompActual As VBComponent

Dim Prefijos() As String 'lista de prefijos antes de una cadena

Private Sub OKButton_Click()
    MsgBox "Operación de complemento en: " & vbinstance.FullName
End Sub

Public Sub LLenarTodo(pVBinstance)
    Set vbinstance = pVBinstance
    Dim vbp As VBProject
    Dim comp As VBComponent 'formularios + modulos
    'Dim vbCont As VBControl 'no lo uso martin al final, lo usa mas adelante
    
    Dim nod As Node
    tvw.Nodes.Clear
    For Each vbp In vbinstance.VBProjects
        'agrego un nodo padre por cada projecto en el ide
        Set tvw.Nodes.Add(, , vbp.Name, vbp.Name).Tag = vbp
        'dentro de este todos los formularios y modulos que contiene
        For Each comp In vbp.VBComponents
            Set tvw.Nodes.Add(vbp.Name, tvwChild, comp.Name, comp.Name).Tag = comp
        Next
    Next
    Me.Show
End Sub

Private Sub cmdInsertar_Click()
    If txt <> "" Then
        If txt.SelLength <> 0 Then
            aux = txt.SelText
            txt.SelText = "Trans(" + aux + ")"
        End If
        
        lvw.SelectedItem.Text = txt
        lvw_AfterLabelEdit 1, txt
        
        'elegir el que sigue viendo que exista!!!
        If lvw.SelectedItem.Index + 1 <= lvw.ListItems.Count Then
            lvw.ListItems(lvw.SelectedItem.Index + 1).Selected = True
        End If
        'y hacerle click
        lvw_ItemClick lvw.ListItems.Item(lvw.SelectedItem.Index)
    End If
End Sub

Private Sub cmdTraducir_Click()
    
    Dim aux As String
    aux = "Private Sub Traducir()"
    Dim li As ListItem
    For Each li In lvwControles.ListItems
        If li.Checked Then
            aux = aux + vbCrLf + vbTab + li.Text + "." + li.Tag.Name + "= Trans(" + Chr(34) + li.Tag.Value + Chr(34) + ")"
        End If
    Next
    
    aux = aux + vbCrLf + "End Sub"
    
    Dim vComp As VBComponent
    Dim M As Member
    Dim encontro As Boolean
    encontro = False
    Set vComp = tvw.SelectedItem.Tag
    
    'busco el evento form load para insertar el llamado a la funcion traducir
    For Each M In vComp.CodeModule.Members
        If M.Name = "Form_Load" Or M.Name = "MDIForm_Load" Then
            vComp.CodeModule.InsertLines M.CodeLocation + 1, _
                vbTab + "Traducir" + " 'Agregado por el complemento traductor "
            encontro = True
            Exit For
        End If
    Next
    
    'si no encuentra el evento form load
    If Not encontro Then
        If vComp.Type = vbext_ct_VBMDIForm Then
            vComp.CodeModule.InsertLines vComp.CodeModule.CountOfLines + 1, _
                "'-------Agregado por el complemento traductor------------" + vbCrLf + _
                "Private Sub MDIForm_Load()" + vbCrLf + _
                vbTab + "Traducir" + vbCrLf + _
                "End Sub"
        Else
            'es un form comun
            vComp.CodeModule.InsertLines vComp.CodeModule.CountOfLines + 1, _
                "'-------Agregado por el complemento traductor------------" + vbCrLf + _
                "Private Sub Form_Load()" + vbCrLf + _
                vbTab + "Traducir" + vbCrLf + _
                "End Sub"
        End If
    End If
    
    'y aca escribo el metodo que traduce las propiedades del form
    vComp.CodeModule.InsertLines vComp.CodeModule.CountOfLines + 1, _
    "'-------Agregado por el complemento traductor------------" + vbCrLf + _
    aux

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
    cargarPanelCodigo
End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMaximized Or Me.WindowState = vbNormal Then
        sTab.Height = Me.Height - 550 - sTab.Top
        sTab.Width = Me.Width - sTab.Left - 200
        lvw.Width = sTab.Width - 200
        lvw.Height = sTab.Height - 600 - cmdInsertar.Height
        lvw.ColumnHeaders(1).Width = lvw.Width - 100
        lvwControles.Width = lvw.Width
        lvwControles.Height = lvw.Height
        tvw.Height = sTab.Height
        cmdTraducir.Top = lvwControles.Height + 500
        cmdTraducir.Width = lvw.Width
        txt.Top = cmdTraducir.Top
        cmdInsertar.Top = cmdTraducir.Top
        txt.Width = lvw.Width - 100 - cmdInsertar.Width
        cmdInsertar.Left = txt.Width + txt.Left + 100
    End If
End Sub

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim vComp As VBComponent
    Set vComp = tvw.Nodes(lvw.Tag).Tag
    vComp.CodeModule.ReplaceLine CLng(lvw.SelectedItem.Tag), NewString
    'vComp.CodeModule.CodePane.SetSelection
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt = Item.Text
    'aca selecciono el texto a envolver
    a = InStr(1, txt, Chr(34))
    
    'si me voy a la ultima comilla?
    'b = InStrRev(txt, Chr(34))
    'no sirve demasiado ...
    
    b = InStr(a + 1, txt, Chr(34))
    
    'XXXXXXX
    'ES POSIBLE QUE SOLO HAYA UNA COMILLA _
        cuando se lista solo se pide una comilla
    'no es posible que b=0 con el instrev pero si que b=a
    If b = 0 Or b = a Then Exit Sub
    'saliendo en b=0 evita un error
    
    txt.SelStart = a - 1
    txt.SelLength = b - a + 1
    'txt.SetFocus
End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    cmdInsertar_Click
    lvw.SetFocus
    End If
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuPopUp
End Sub

Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 1 Then cargarPanelCodigo
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next
Dim comp
Set comp = Node.Tag
If TypeOf comp Is VBComponent Then
    
    Set vCompActual = comp
    cargarPanelCodigo
    
    lvwControles.ListItems.Clear
    Dim vControl As VBControl
    Dim vProperty As Property
    Dim li As ListItem
    
    'cargar todas las propiedades de los objetos que tengan valores tipo string
    For Each vControl In vCompActual.Designer.VBControls
        For Each vProperty In vControl.Properties
            If TypeName(vProperty.Value) = "String" Then
                'ver si tiene indice!!!
                If vControl.Properties("Index") > -1 Then
                    Set li = lvwControles.ListItems.Add(, , vControl.Properties("Name") + "(" + CStr(vControl.Properties("Index")) + ")")
                Else
                    Set li = lvwControles.ListItems.Add(, , vControl.Properties("Name"))
                End If
                
                Set li.Tag = vProperty
                li.Bold = IIf(vProperty.Name = "Caption" Or vProperty.Name = "Text", True, False)
                li.ListSubItems.Add , , vProperty.Name
                li.ListSubItems.Add , , vProperty.Value
            End If
        Next
    Next
    
Else
    lvw.ListItems.Clear
    lvwControles.ListItems.Clear
    Set vCompActual = Nothing
End If
End Sub

Private Sub cargarPanelCodigo()
    If Not vCompActual Is Nothing Then
        lvw.ListItems.Clear
        
        Dim Linea As String
      
        lvw.Tag = vCompActual.Name
        ReDim Prefijos(0)
        'una vuelta para buscar los prefijos y mostrar algo mas depurado
        For I = 1 To vCompActual.CodeModule.CountOfLines
            
            Linea = vCompActual.CodeModule.Lines(I, 1)
            
            'no tomar lineas comentadas
            If Left(Trim(Linea), 1) = "'" Then GoTo SIG
            
            'ver que no haya una comilla simple antes de la primera comilla!
            If InStr(1, Linea, "'") > 0 Then 'asegurarse que haya una comilla
                If InStr(1, Linea, "'") < InStr(1, Linea, Chr$(34)) Then GoTo SIG
            End If
            
            Dim px As String 'prefijo actual
            If InStr(1, Linea, Chr$(34)) <> 0 Then
                px = GetPrefijo(Linea)
                AddPrefijo px
            End If
SIG:
        Next
        
        frmAnulla.LoadPX Prefijos
        
        'cuando vuelva se habran redefinido los prefijos a evitar
        
        'ahora cargo evitando los prefijos que no van
        'y dejando un codigo mas limpio
        
        For I = 1 To vCompActual.CodeModule.CountOfLines
             Linea = vCompActual.CodeModule.Lines(I, 1)
             
             'XXXXXXXXX
             'no toma lineas comentadas
             If Left(Trim(Linea), 1) = "'" Then GoTo SIG2
             'tampoco una cadena que uso mucho
             'If InStr(1, linea, "terr", vbTextCompare) Then GoTo SIG
             
             If InStr(1, Linea, Chr$(34)) <> 0 Then
                'si existe el prefijo en la lista de los que se ignoran no lo uso
                If ExistePrefijo(GetPrefijo(Linea)) = False Then
                    'en el tag pongo el numero de linea
                    lvw.ListItems.Add(, , Linea).Tag = I
                End If
            End If
            
SIG2:
        Next
        
    End If
    
    MsgBox "Listo"

End Sub

Private Function ExistePrefijo(sPrefix As String) As Boolean

    ExistePrefijo = False
    If UBound(Prefijos) = 0 Then Exit Function
        
    For H = 1 To UBound(Prefijos)
        If LCase(Prefijos(H)) = LCase(sPrefix) Then
            ExistePrefijo = True
            Exit Function
        End If
    Next H
End Function

Private Sub AddPrefijo(sPrefix As String)
    Dim H As Long
    
    'ver si ya existe!
    H = UBound(Prefijos) + 1
    If H > 1 Then
        For H = 1 To UBound(Prefijos)
            If LCase(Prefijos(H)) = LCase(sPrefix) Then Exit Sub
        Next H
    End If
    
    'lo agrego nomas
    H = UBound(Prefijos) + 1
    ReDim Preserve Prefijos(H)
    Prefijos(H) = sPrefix

End Sub

Public Sub SetNewPrefijos(nPX() As String)
    'definir los prefijos que realmente se van a usar
    Prefijos = nPX
End Sub

Private Function GetPrefijo(sLinea As String) As String
    'el prefijo esta antes de "
    'una apertura de parentesis no es un prefijo
    
    'busco a partir de que punto busco el prefijo
    Dim FinPrefijo As Long
    FinPrefijo = InStr(1, sLinea, Chr(34))
    'si hay un inicio de parentesis no lo cuento
    If Mid(sLinea, FinPrefijo - 1, 1) = "(" Then FinPrefijo = FinPrefijo - 1
    If Mid(sLinea, FinPrefijo - 1, 1) = " " Then FinPrefijo = FinPrefijo - 1
    
    'veo los posibles inicios del prefijo y tomo el que este mas adelante
    Dim IniPrefijo(3) As Long
    IniPrefijo(0) = InStrRev(sLinea, ".", FinPrefijo - 1) 'pic.loadpicture("asass")
    IniPrefijo(1) = InStrRev(sLinea, " ", FinPrefijo - 1) 'v = "juju"
    IniPrefijo(2) = InStrRev(sLinea, "=", FinPrefijo - 1) 'v = "juju"
    IniPrefijo(3) = InStrRev(sLinea, "(", FinPrefijo - 1) 'v = leerconfig(gpf("asas"))
    
    Dim M As Long
    M = Max(IniPrefijo) + 1

    GetPrefijo = Mid(sLinea, M, FinPrefijo - M)
    
End Function

Private Function Min(v() As Long) As Long
    Dim H As Long
    Dim Mini As Long: Mini = 99999
    
    For H = 0 To UBound(v)
        '>0 es por que no encontro, no USAR!
        If v(H) < Mini And v(H) > 0 Then Mini = v(H)
    Next H
    
    Min = Mini
End Function

Private Function Max(v() As Long) As Long
    Dim H As Long
    Dim Maxi As Long: Maxi = -99999
    
    For H = 0 To UBound(v)
        If v(H) > Maxi Then Maxi = v(H)
    Next H
    
    Max = Maxi
End Function

