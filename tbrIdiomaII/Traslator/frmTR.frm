VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTR 
   Caption         =   "Traduccion"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txResumen 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5100
      Width           =   3255
   End
   Begin VB.TextBox txTags 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2490
      Width           =   3255
   End
   Begin VB.TextBox txObservaciones 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6270
      Width           =   5205
   End
   Begin VB.CheckBox chkNoEntiendo 
      Caption         =   "No entiendo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3810
      TabIndex        =   8
      Top             =   6360
      Width           =   2235
   End
   Begin VB.CheckBox chkNoTerminada 
      Caption         =   "Me falta revisar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3810
      TabIndex        =   7
      Top             =   6060
      Width           =   2235
   End
   Begin VB.CheckBox chkTerminada 
      Caption         =   "Terminada"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3810
      TabIndex        =   6
      Top             =   6660
      Width           =   2235
   End
   Begin VB.TextBox txExpli 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3390
      Width           =   3255
   End
   Begin VB.TextBox txTRAD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4110
      Width           =   4875
   End
   Begin VB.TextBox txBASE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2190
      Width           =   4875
   End
   Begin MSComctlLib.ListView lvIDs 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   12197
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvTAGs 
      Height          =   2055
      Left            =   7830
      TabIndex        =   1
      Top             =   120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvIDMs 
      Height          =   2055
      Left            =   3810
      TabIndex        =   4
      Top             =   120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Observaciones del traductor para programador:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   6030
      Width           =   4005
   End
   Begin VB.Label Label3 
      Caption         =   "Tag usados:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8790
      TabIndex        =   14
      Top             =   2250
      Width           =   2145
   End
   Begin VB.Label Label2 
      Caption         =   "Variables usadas:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   4860
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "Explicacion:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   3150
      Width           =   975
   End
   Begin VB.Menu mnFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnSave 
         Caption         =   "Grabar"
      End
      Begin VB.Menu mnExport 
         Caption         =   "Exportar"
      End
      Begin VB.Menu MnExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnCadena 
      Caption         =   "Cadena"
      Begin VB.Menu mnAddCadena 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnAddvariable 
         Caption         =   "Agregar Variable"
      End
      Begin VB.Menu mnAddTag 
         Caption         =   "Agregar Tag"
      End
   End
End
Attribute VB_Name = "frmTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mnPhr As tbrPrhase.clsPhraseMNG 'manejador de todo esta traduccion

'********************************
'cosas que necesito a mano siempre
Private IDIOMA As String 'ultimo idioma elegido para acceder rapido a el sin tener que leer el LVW cada vez
Private TagActual As String 'id anterior
Private sID As String 'ID alegido del LvwIDs
Private PHRActual As clsPRHASE
'********************************

Private Sub chkNoEntiendo_Click()
    PHRActual.GetTransObj(IDIOMA).NoEntiendo = CBool(chkNoEntiendo.Value)
End Sub

Private Sub chkNoTerminada_Click()
    If chkNoTerminada.Value > 0 Then chkTerminada.Value = 0
    PHRActual.GetTransObj(IDIOMA).NoTerminada = CBool(chkNoTerminada.Value)
    PHRActual.GetTransObj(IDIOMA).Terminada = CBool(chkTerminada.Value)
End Sub

Private Sub chkTerminada_Click()
    If chkTerminada.Value > 0 Then chkNoTerminada.Value = 0
    PHRActual.GetTransObj(IDIOMA).NoTerminada = CBool(chkNoTerminada.Value)
    PHRActual.GetTransObj(IDIOMA).Terminada = CBool(chkTerminada.Value)
End Sub

'quedeaqui
'2- no probe grabar vars(0) o sea la explicacion ni otras variables
'3- no se muestran los porcentajes de tags traducidos, deberia tener una funcion para actualizarlo
'4- necesito una herramienta para agregar cadenas que se pueda bloquear
'5- necesito saber cuando se va a grabar cada traduccion, si en el change del txt o con boton
'6- Necesito los chks de las propiedades nuevas de los mTrans de cada phr


Private Sub Form_Load()

    Traducir

    'acomodar los listviwes
    'IDIOMAS
    lvIDMs.View = lvwReport
    lvIDMs.ColumnHeaders.Add , , T.GetText("000006") '"Idioma"
    lvIDMs.ColumnHeaders.Add , , T.GetText("000019") '"Hecho" ' x / y, cuantas cadenas estan traducidas
    lvIDMs.ColumnHeaders.Add , , "%" ' valor en porcentaje de lo anterios
    
    'TAGs
    lvTAGs.View = lvwReport
    lvTAGs.ColumnHeaders.Add , , "Tag"
    lvTAGs.ColumnHeaders.Add , , T.GetText("000019") '"Hecho" ' x / y, cuantas cadenas estan traducidas
    lvTAGs.ColumnHeaders.Add , , "%" ' valor en porcentaje de lo anterios
    
    'IDs a traducir
    lvIDs.View = lvwReport
    lvIDs.ColumnHeaders.Add , , "ID"
    lvIDs.ColumnHeaders.Add , , T.GetText("000020") '"Fecha base" 'fecha en que se modifico esta base
    lvIDs.ColumnHeaders.Add , , T.GetText("000021") '"Fecha Trad" 'fecha en que se traduxjo para el idioma elegido
    
    lvIDMs.LabelEdit = lvwManual
    lvIDMs.FullRowSelect = True
    lvIDMs.HideSelection = False
    
    lvTAGs.LabelEdit = lvwManual
    lvTAGs.FullRowSelect = True
    lvTAGs.HideSelection = False
    
    lvIDs.LabelEdit = lvwManual
    lvIDs.FullRowSelect = True
    lvIDs.HideSelection = False
    
End Sub

'cargar un archivo de idioma
Public Sub OpenPhr(M As tbrPrhase.clsPhraseMNG)
    
    Set mnPhr = M
    
    'IDIOMAS disponibles
    CargarIdiomas
    
    'TAGs usados
    CargarTags
    
    'lista de cadenas de traduiccion segun sus ids
    CargarIds
    
    Me.Show
    
End Sub

Private Sub CargarIdiomas()

    lvIDMs.ListItems.Clear
    Dim RET() As String, J As Long
    RET = mnPhr.GetIdiomas
    
    Dim ItmX As ListItem
    
    Dim Oks As Long, Tots As Long
    For J = 1 To UBound(RET)
        
        Tots = mnPhr.GetCadenasCantidad
        Oks = mnPhr.GetCadenasOk(RET(J))
        
        
        Set ItmX = lvIDMs.ListItems.Add(J, , RET(J))
            ItmX.SubItems(1) = CStr(Oks) + "/" + CStr(Tots)
            ItmX.SubItems(2) = CStr(Round(Oks / Tots * 100, 2)) + " %"
    Next J
End Sub

Private Sub CargarTags(Optional IDM As String = "")
    lvTAGs.ListItems.Clear
    Dim RET() As String
    RET = mnPhr.GetTags
    
    Dim ItmX As ListItem
    
    'cargar un tag que sea "TODOS"
    Set ItmX = lvTAGs.ListItems.Add(1, , "TODOS")
    
    Dim Oks As Long, Tots As Long
    
    For J = 1 To UBound(RET) 'POR AHORA no pongo estadisticas por que se necesita que haya un idioma elegido
        If IDM <> "" Then
            Tots = mnPhr.GetCadenasCantidad(RET(J))
            Oks = mnPhr.GetCadenasOk(IDM, RET(J))
        End If
        
        Set ItmX = lvTAGs.ListItems.Add(J + 1, , RET(J))
            If IDM <> "" Then
                ItmX.SubItems(1) = CStr(Oks) + "/" + CStr(Tots)
                ItmX.SubItems(2) = CStr(Round(Oks / Tots * 100, 2)) + " %"
            End If
            
    Next J
End Sub

Private Sub CargarIds(Optional IDM As String = "", Optional sTag As String = "", Optional SelId As String = "")
    'IDs a traducir
    Dim J As Long
    Dim ItmX As ListItem
    
    Dim fBase As Long, fTrans As Long 'fechas como numeros
    Dim sBase As String, sTrans As String 'fechas string
    
    lvIDs.ListItems.Clear
    Dim Usar As Boolean, Usados As Long
    Usados = 0
    For J = 1 To mnPhr.GetPhrCantidad
        
        fBase = mnPhr.GetPhrByNum(J).FechaBase
        fTrans = mnPhr.GetPhrByNum(J).GetTransObj(IDM).fechaTrans
        'If fTrans = 0 Then fTrans = CLng(Now) 'para que no de error de que esta venida
        'lo saque para quede en rojo si no esta ???
        sBase = CStr(CDate(fBase))
        If fTrans = 0 Then
            sTrans = ""
        Else
            sTrans = CStr(CDate(fTrans))
        End If
        
        Usar = False
        If sTag <> "" Then
            Usar = mnPhr.GetPhrByNum(J).HasTag(sTag)
        Else
            Usar = True 'van todos
        End If
        
        If Usar Then
            Usados = Usados + 1
            Set ItmX = lvIDs.ListItems.Add(Usados, , mnPhr.GetPhrByNum(J).sID)
                ItmX.SubItems(1) = sBase
                ItmX.SubItems(2) = sTrans
        
            If mnPhr.GetPhrByNum(J).GetTransObj(IDM).Terminada Then ItmX.ForeColor = &H800000
            If mnPhr.GetPhrByNum(J).GetTransObj(IDM).NoEntiendo Then ItmX.ForeColor = &H800080
            If fTrans < fBase Then ItmX.ForeColor = vbRed
            
            'dejarlo marcado si me lo piden
            If SelId = mnPhr.GetPhrByNum(J).sID Then
                ItmX.Selected = True
                ItmX.EnsureVisible
            End If
            
        End If
           
    Next J
End Sub

Private Sub lvIDMs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IDIOMA = lvIDMs.SelectedItem.Text
    CargarIds IDIOMA, TagActual
    
    'aqui debo cargar los tags para que se vean los porcentajes que estan en el idioma actual
    CargarTags IDIOMA
End Sub

Private Sub lvIDs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If IDIOMA = "" Then Exit Sub
    
    'mostrar la base y la traduccion de esto
    sID = lvIDs.SelectedItem.Text
    
    Set PHRActual = mnPhr.getPHR(sID)
    
    txBASE = PHRActual.BaseText
    txTRAD = PHRActual.GetTrans(IDIOMA)
    chkNoTerminada.Value = Abs(CLng(PHRActual.GetTransObj(IDIOMA).NoTerminada))
    chkTerminada.Value = Abs(CLng(PHRActual.GetTransObj(IDIOMA).Terminada))
    chkNoEntiendo.Value = Abs(CLng(PHRActual.GetTransObj(IDIOMA).NoEntiendo))
    txObservaciones.Text = PHRActual.GetTransObj(IDIOMA).Observaciones
    txExpli.Text = PHRActual.GetVar
    txTags.Text = PHRActual.GetStrTagsByColon
    
    
    Dim RET As String 'texto completo de la explicacion y sus variables
    RET = ""
    Dim J As Long
    For J = 1 To PHRActual.GetVarCantidad
        RET = RET + "Variable %" + CStr(J) + "%:" + vbCrLf + PHRActual.GetVar(J) + vbCrLf
    Next J
    
    txResumen.Text = RET
    
End Sub

Private Sub lvTAGs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TagActual = lvTAGs.SelectedItem.Text
    If TagActual = "TODOS" Then TagActual = ""
    
    CargarIds IDIOMA, TagActual
End Sub

Private Sub mnAddCadena_Click()
    'solo le pido el ID, lo demas lo llena en los textos
    Dim RET As String
    RET = InputBox("Inserte ID", , mnPhr.GetAutoID)
    
    Set PHRActual = mnPhr.QuickAddPhr(RET, "DEFINIR TEXTO", "", "")
    
    'cargarlo y elegirlo en la lista
    CargarIds , , RET
End Sub

Private Sub mnAddTag_Click()
    Dim RET2 As String
    RET2 = InputBox("Escriba el nuevo tag")
    
    PHRActual.AddTag RET2
    
    'si el tag no existia agregarlo a la lista
    CargarTags
End Sub

Private Sub mnAddvariable_Click()
    'que me diga el numero que quiere y el texto
    Dim RET As Long
    RET = CLng(InputBox("Indique numero de variable"))
    
    Dim RET2 As String
    RET2 = InputBox("Indique explicacion de esa variable")
    
    PHRActual.SetVar RET, RET2
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub mnExport_Click()
    mnPhr.ExportAll AP + "Exports"
End Sub

Private Sub mnSave_Click()
    mnPhr.Save "LAST"
End Sub

Private Sub txBASE_Change()
    PHRActual.BaseText = txBASE
End Sub

Private Sub txExpli_Change()
    PHRActual.SetVar 0, txExpli.Text
End Sub

Private Sub txObservaciones_Change()
    PHRActual.GetTransObj(IDIOMA).Observaciones = txObservaciones
End Sub

Private Sub txTRAD_Change()
    PHRActual.SetTrans txTRAD.Text, IDIOMA
End Sub

Private Sub Traducir()
    mnFile.Caption = T.GetText("000001")
    mnSave.Caption = T.GetText("000004")
    MnExit.Caption = T.GetText("000007")
    mnExport.Caption = T.GetText("000011")
    
    mnCadena.Caption = T.GetText("000012")
    mnAddCadena.Caption = T.GetText("000013")
    mnAddvariable.Caption = T.GetText("000014")
    mnAddTag.Caption = T.GetText("000015")
    
    chkNoTerminada.Caption = T.GetText("000016")
    chkTerminada.Caption = T.GetText("000017")
    chkNoEntiendo.Caption = T.GetText("000018")
    
End Sub
