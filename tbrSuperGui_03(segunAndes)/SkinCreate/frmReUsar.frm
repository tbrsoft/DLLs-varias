VERSION 5.00
Begin VB.Form frmReUsar 
   Caption         =   "Elija su objeto"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   904
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   1890
      Top             =   930
   End
   Begin VB.HScrollBar HS 
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   4770
      Width           =   3135
   End
   Begin VB.VScrollBar VS 
      Height          =   3285
      Left            =   5460
      TabIndex        =   4
      Top             =   1020
      Width           =   315
   End
   Begin VB.PictureBox picTodo 
      BackColor       =   &H00000080&
      Height          =   3375
      Left            =   2340
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   2
      Top             =   930
      Width           =   3045
      Begin VB.PictureBox picTipo 
         Height          =   885
         Index           =   0
         Left            =   180
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdNEW 
      Caption         =   "nuevo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2610
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdTipo 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "frmReUsar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'este formulario esta pensado para que objetos previamnete definidos se carguen automáticamente
'en un formulario por tipo
'de esta forma en vez de crearse objetos predefinidos se usan objetos completos creados previamente

Dim FolReUtil As String 'carpeta con todos los objetos reutilizables

'es indispensable un objeto GUI global para que se incialize el qAlgun FRM
Private G As tbrSuperGui_3.clsGUI

Dim Tipos() As tbrSuperGui_3.ObjFullPadre  'todos los de un mismo tipo sera hijos de un mismo padre para mostrarse juntos uno debajo del otro
Dim tipoElegido As Long 'en cada momento que frm esta elegido

'ultimo objeto elegido para usar en el skineador
Public ObjElegido As tbrSuperGui_3.objFULL

Private Obj_Que_Necesita__DibujarTexto() As Object 'lista de objetos que se refrescaran con el timer

Public Sub CargarTodo(Optional mModal As Long = 0)
    'ver todos los objetos que hay y crear un PictureBox por cada tipo
    'para que se muetren separados por tipo
    
    On Local Error GoTo ErrCarg
        
    TERR.Anotar "aba", FolReUtil
    
    Init
    
    If FSO.FolderExists(FolReUtil) = False Then
        'MsgBox "No hay objetos reutilizables !!"
    End If
    
    TERR.Anotar "abb"
    CargarPadres 'base en la que entran todos los objetos
    
    'ver cada objeto de la carpeta folReUtil
    Dim FI As File
    Dim FO As Folder
    
    Set FO = FSO.GetFolder(FolReUtil)
    
    Dim LastY() As Long 'ultima coordenada usada
    ReDim LastY(G.MNG.GetPadresMaxID) 'una coordenada y para cada tipo!
    
    'cada uno de los objetos que se van cargando
    Dim OG() As New tbrSuperGui_3.objFULL  'creo un objeto que lo cargara y mostrar
    Dim C As Long
    Dim H6 As Long 'resultado de la carga de los lods
    For Each FI In FO.Files
        TERR.Anotar "abc", FI.path, C
        'debe fijarse solo de que tipo es
        C = C + 1
        ReDim Preserve OG(C)
        H6 = OG(C).Load(FI.path, FolReUtil + "fol_" + FI.Name)
        TERR.Anotar "abd-60", H6
        'LAS COORDENADAS X e Y no importan AQUI ! yo los dibujo uno por debajo del otro
        OG(C).oSimple.X = 5
        OG(C).oSimple.Y = LastY(OG(C).Tipo) + 5
        
        TERR.Anotar "abd-59", LastY(OG(C).Tipo), OG(C).oSimple.SGOName
        
        LastY(OG(C).Tipo) = LastY(OG(C).Tipo) + OG(C).oSimple.H + 5
        
        Tipos(OG(C).Tipo).AppendSGO OG(C)
        
        TERR.Anotar "abd-58"
        
        'reacomodar el pic segun sea necesario
        picTipo(OG(C).Tipo).Height = LastY(OG(C).Tipo) + 5
        If OG(C).oSimple.W > picTipo(OG(C).Tipo).Width Then
            TERR.Anotar "abd", OG(C).oSimple.W, picTipo(OG(C).Tipo).Width
            picTipo(OG(C).Tipo).Width = OG(C).oSimple.W + 10
        End If
        
        TERR.Anotar "abd-56"
        'algunos tipos necesita "dibujar texto cada muy poco tiempo!
        If OG(C).Tipo = en_tbrPromociones2 Then
            TERR.Anotar "abd-57"
            Add_Obj_Que_Necesita__DibujarTexto OG(C).oManu
        End If
        
    Next
    
    TERR.Anotar "abe"
    Me.Show mModal
    
    Exit Sub
ErrCarg:
    TERR.AppendLog "ErrCarg", TERR.ErrToTXT(Err)
    'Resume Next
End Sub

Private Sub refreshPDR(J As Long)
    TERR.Anotar "abf", J
    picTipo(J).Visible = True
    Tipos(J).INIT_GRAPH "CLOSE"
    Tipos(J).INIT_GRAPH
    picTipo(J).Refresh
End Sub


Private Sub LoadPicTipo(I As Long)
    On Local Error Resume Next 'puede ser que ya se haya cargado
    TERR.Anotar "abg"
    Load picTipo(I)
    Load cmdTipo(I)
    cmdTipo(I).Caption = Tipos(0).getStrTipo(I)
End Sub

Private Sub CargarPadres()
    'cargar todos los tipos posibles de objetos
    On Local Error GoTo errPDR
    TERR.Anotar "abh"
    
    
    cmdTipo(0).Top = 0
    cmdTipo(0).Left = 5
    
    Dim J As Long
    For J = 1 To Tipos(0).CantTiposDatos
        TERR.Anotar "abi", J
        ReDim Preserve Tipos(J)
        Set Tipos(J) = G.MNG.AddPadre("FRM_" + Tipos(0).getStrTipo(J)) 'crea a cada uno con nombre segun tipo de objetos que contendra
        
        TERR.Anotar "abk"
        LoadPicTipo J
        cmdTipo(J).Top = cmdTipo(J - 1).Top + cmdTipo(J - 1).Height + 3
        cmdTipo(J).Left = cmdTipo(J - 1).Left
        cmdTipo(J).Visible = True
        
        picTipo(J).Left = 3 ' cmdTipo(0).Left + cmdTipo(0).Width + 5
        picTipo(J).Top = 3 ' cmdTipo(1).Top
        picTipo(J).Visible = False
        TERR.Anotar "abl", J, Tipos(J).sHDC
    Next J
    
    UpdateHDCs
    
    'showMeHdcs
    
    TERR.Anotar "abm"
    Exit Sub
errPDR:
    TERR.AppendLog "errPDR", TERR.ErrToTXT(Err)
End Sub

Private Sub cmdNEW_Click()

    On Local Error GoTo errNEW
    TERR.Anotar "abn"
    
    'si si es un archivo que se va a grabar
    Dim PAD As String
    PAD = InputBox("Inserte nombre del objeto nuevo", "Nuevo objeto", "OBJ001")

    'crear un nuevo objeto de la clase eklegida y grabarlo aqui como repositorio
    Dim OG As New objFULL
    'seguiraqui, que bueno que seria que cada objeto tuviera una base de lo predeterminado
    OG.oSimple.SGOName = PAD
    OG.Tipo = tipoElegido
    OG.oSimple.W = 150
    OG.oSimple.H = 70
    OG.oSimple.X = 5
    OG.oSimple.Y = picTipo(tipoElegido).Height + 5
    
    '************************************************
    'propiedades especificas segun objeto!
    If OG.Tipo = en_clsLabel Then
        OG.oSimple.SetProp "TextoActual", OG.oSimple.SGOName
    End If
    
    If OG.Tipo = en_tbrPromociones2 Then
        'las propiedades "texto_num" son despues agregadas como textos
        OG.oSimple.SetProp "Texto_1", "1º" + vbCrLf + "Promocion" + vbCrLf + "Chonga !!!"
        OG.oSimple.SetProp "Texto_2", "2º" + vbCrLf + "roProcion" + vbCrLf + "Chonga !!!"
        OG.oSimple.SetProp "Texto_3", "3º" + vbCrLf + "propaganda" + vbCrLf + "Chonga !!!"
        OG.oSimple.SetProp "PixelSalteo", "3"
        OG.oSimple.SetProp "TiemposEntreTextos", "40"
    End If
    
    If OG.Tipo = en_clsTemasManager Then
        OG.oSimple.W = 250
        OG.oSimple.H = 350
        OG.oSimple.SetProp "AlphaB", "1"
    End If
    
    If OG.Tipo = en_tbrTextoSelect Then
        OG.oSimple.W = 450
        OG.oSimple.H = 35
    End If
    
    If OG.Tipo = en_clsDiscoManager Then
        OG.oSimple.W = 600
        OG.oSimple.H = 400
    End If
    '************************************************
    '************************************************
    OG.CreateManu 'genere un nuevo objeto manu basado en oSimple interno
    '************************************************
    
    
    '************************************************
    'algunas clases req uieren que primero se cree el objeto del manu y despues manosearlo
    '************************************************
    '************************************************
    
    Dim CM As New CommonDialog, F As String
    
    If OG.Tipo = en_clsPNGBoton Then
        'elegir el PNG a usar !
        
        CM.InitDir = FolReUtil + "pngs_reusables"
        CM.Filter = "Imagenes PNG |*.png"
        CM.ShowOpen
        
        F = CM.FileName
        
        If F <> "" Then
            Dim PB As New clsPNGBoton
            Set PB = OG.oManu
            PB.SetPNGUnSel F
            'no alcanza con esto, debe ser parte de la comeccion de archivos fiMG
            OG.getFIMG.AddFileByPath F, , "PngUnSel"
            
            'SOLO A LOS FINES DE QUE LO CARGUE Y LEA EL WI Y HE PARA QUE SE GRABEN OK!
            PB.IniciarPNGs F
        End If
    End If
    
    If OG.Tipo = en_clsDiscoManager Then
        'elegir el PNG a usar !
        
        CM.InitDir = FolReUtil + "pngs_MarcosDiscos"
        CM.Filter = "Imagenes PNG |*.png"
        CM.DialogTitle = "PNG DEL MARCO DE DISCOS INTERIORES"
        CM.ShowOpen
        
        F = CM.FileName
        
        If F <> "" Then
            Dim DM As New clsDiscoManager
            Set DM = OG.oManu
            DM.SetPNGMarcoDisco F
            'no alcanza con esto, debe ser parte de la comeccion de archivos fiMG
            OG.getFIMG.AddFileByPath F, , "PNGMarcoDisco"
            
            'SOLO A LOS FINES DE QUE LO CARGUE Y LEA EL WI Y HE PARA QUE SE GRABEN OK!
            DM.IniciarPNGs F
        End If
    End If
    
    '************************************************
    '************************************************
    
    TERR.Anotar "abo", PAD, tipoElegido
    Tipos(tipoElegido).AppendSGO OG
    
    'agrandar el picture box para que entre lo nuevo
    picTipo(tipoElegido).Height = picTipo(OG.Tipo).Height + OG.oSimple.H + 5
    
    If OG.oSimple.W > picTipo(tipoElegido).Width Then
        picTipo(tipoElegido).Width = OG.oSimple.W + 10
    End If
    
    TERR.Anotar "abp", tipoElegido
    'mostrar todo
    refreshPDR tipoElegido 'esta carga grafica termina de llenar todas las propiedades a valores predeterminados
    'es pos eso que va ANTES de SAVE
    'si se desea grabar deberia siem,pre hacerse un init graph
    
    TERR.Anotar "abq", FolReUtil, PAD
    'mostrar su edicion y permitir grabarlo
    OG.Save FolReUtil + PAD + "." + Tipos(0).getStrTipo(OG.Tipo)
    
    'marcarlo por si lo quiere agregar
    Set ObjElegido = OG
    TERR.Anotar "abr", tipoElegido
    OG.GuiEdit Tipos(tipoElegido) 'mostrar el editor
    
    TERR.Anotar "abs"
    UpdateScrolls
    
    'si requiere refrescadas ...
    If OG.Tipo = en_tbrPromociones2 Then
        Add_Obj_Que_Necesita__DibujarTexto OG.oManu
    End If
    
    Exit Sub
errNEW:
    TERR.AppendLog "errNEW", TERR.ErrToTXT(Err)
End Sub

Private Sub cmdTipo_Click(Index As Integer)
    'esconder todo y mostrar solo el que va
    
    TERR.Anotar "abt", Index
    
    tipoElegido = Index
    Dim J As Long
    For J = 1 To cmdTipo.UBound
        If Index <> J Then
            cmdTipo(J).Font.Bold = False
            picTipo(J).Visible = False
        End If
    Next J
    
    'solo mostrar ele elgido
    UpdateHDCs 'NO DEBERIA SER NECESARIO QUEDEAQUI
    
    TERR.Anotar "abz", Index, Tipos(Index).sName, picTipo(Index).hdc, Tipos(Index).sHDC
    cmdTipo(Index).Font.Bold = True
    picTipo(Index).Visible = True
    refreshPDR CLng(Index)
    
    UpdateScrolls
        
    Timer1.Enabled = True
    'showMeHdcs
End Sub

Private Sub Init()
    TERR.Anotar "abu"
    Me.AutoRedraw = True
    Me.ScaleMode = 3
    
    picTipo(0).AutoRedraw = True
    picTipo(0).BorderStyle = 0
    picTipo(0).BackColor = vbWhite
    picTipo(0).ScaleMode = 3
    
    TERR.Anotar "abv"
    FolReUtil = AP + "ReUtilizar\"
    Set G = New clsGUI
    G.SetParaPantallaPixeles 1024, 768
    G.SetPathLog AP + "LogReUtil.log"
    
    TERR.Anotar "abw"
    'hay una coleccion de padres, uno por cada tipo de objetos, el indice cero tiene uso especial
    ReDim Tipos(0) 'este es para calcular cantidades de tipos y otros
    Set Tipos(0) = G.MNG.AddPadre("INICIAL")
    'tipos(0) es el index 1 en g.mng.padres
    
    ReDim Obj_Que_Necesita__DibujarTexto(0)
    Timer1.Enabled = False
    Timer1.Interval = 50
End Sub

Private Sub CerrarTodo()
    TERR.Anotar "abx"
    Dim J As Long, mx As Long
    mx = G.MNG.GetPadresMaxID
    For J = 2 To mx ' Tipos(0).CantTiposDatos
        Tipos(J - 1).INIT_GRAPH "CLOSE" 'tipos(0) es el index 1 en g.mng.padres
    Next J
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TERR.Anotar "aby"
    CerrarTodo
End Sub

Private Sub showMeHdcs()
    'solo pruebas
    For J = 1 To Tipos(0).CantTiposDatos
        TERR.Anotar "aca", Tipos(J).sHDC, picTipo(J).hdc
    Next J
    TERR.AppendLog "TEST2"
End Sub

Private Sub UpdateHDCs()
    For J = 1 To Tipos(0).CantTiposDatos
        Tipos(J).sHDC = picTipo(J).hdc
    Next J
End Sub

Private Sub Form_Resize()

    On Local Error Resume Next

    picTodo.Top = cmdTipo(1).Top
    picTodo.Left = cmdTipo(0).Left + cmdTipo(0).Width + 5
    
    picTodo.Width = Me.Width / 15 - picTodo.Left - 35
    picTodo.Height = Me.Height / 15 - picTodo.Top - 55
    
    VS.Top = picTodo.Top
    VS.Left = picTodo.Left + picTodo.Width
    VS.Height = picTodo.Height
    
    HS.Top = picTodo.Top + picTodo.Height
    HS.Left = picTodo.Left
    HS.Width = picTodo.Width
    
    cmdTipo(0).Top = 5
    cmdTipo(0).Left = 5
    cmdNEW.Top = cmdTipo(0).Top
    cmdNEW.Left = cmdTipo(0).Left + cmdTipo(0).Width + 5
    
    UpdateScrolls
    
End Sub

Private Sub UpdateScrolls()
    'Ver si deben estar enabled o no!
    VS.Enabled = (picTipo(tipoElegido).Height > picTodo.Height)
    HS.Enabled = (picTipo(tipoElegido).Width > picTodo.Width)
    
    Dim Dif As Single
    If VS.Enabled Then 'ver que movimiento tiene
        Dif = picTipo(tipoElegido).Height - picTodo.Height
        VS.LargeChange = Dif / 4
        VS.SmallChange = Dif / 16
        'seguiraqui, asegurarse que ninguno de cero
        VS.Min = 0
        VS.Max = CLng(Dif)
        VS.Value = Abs(picTipo(tipoElegido).Top)
        
    End If
    
    If HS.Enabled Then 'ver que movimiento tiene
        Dif = picTipo(tipoElegido).Width - picTodo.Width
        HS.LargeChange = Dif / 4
        HS.SmallChange = Dif / 16
        
        HS.Min = 0
        HS.Max = CLng(Dif)
        HS.Value = Abs(picTipo(tipoElegido).Left)
        
    End If
End Sub

Private Sub HS_Change()
    picTipo(tipoElegido).Left = -HS.Value
End Sub

Private Sub picTipo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ver donde carajo estoy
    Dim J As Long, J2 As Long
    J2 = Tipos(Index).GetSgoMaxID
    
    Dim ixSEL As Long 'indice del objeto elegido
    ixSEL = -1
    For J = 1 To J2
        Dim obX As Long, obY As Long, obW As Long, obH As Long
        obX = Tipos(Index).GetSgoByID(J).oSimple.X
        obY = Tipos(Index).GetSgoByID(J).oSimple.Y
        obH = Tipos(Index).GetSgoByID(J).oSimple.H
        obW = Tipos(Index).GetSgoByID(J).oSimple.W
        
        If X >= obX And X <= obX + obW Then
            If Y >= obY And Y <= obY + obH Then
                'se hizo click en este !!
                Set ObjElegido = Tipos(Index).GetSgoByID(J)
                ObjElegido.GuiEdit Tipos(Index)
                ixSEL = J
                Exit For
            End If
        End If
    Next J
    
    If ixSEL = -1 Then 'no le pego a ningun objeto!
        Tipos(Index).GuiEdit  'mostrar las propiedades del form!
        Exit Sub
    End If
    
    'ademas de abrir el editor ver si necesita algo mas (SEGUN OBJETO DEL MANU)
    If Tipos(Index).GetSgoByID(ixSEL).Tipo = en_clsTemasManager Then
        Dim OB As clsTemasManager
        Set OB = Tipos(Index).GetSgoByID(ixSEL).oManu
        OB.DoClick_GetElementoIndex CLng(X), CLng(Y)
    End If
End Sub

Private Sub VS_Change()
    picTipo(tipoElegido).Top = -VS.Value
End Sub


'NEGRADA DEL MANU !!!
Private Sub Timer1_Timer()
    
    Dim J As Long
    For J = 1 To UBound(Obj_Que_Necesita__DibujarTexto)
        Obj_Que_Necesita__DibujarTexto(J).dibujartexto
    Next J
    
    picTipo(tipoElegido).Refresh
    
End Sub

Private Sub Add_Obj_Que_Necesita__DibujarTexto(OB As Object)
    Dim J As Long
    J = UBound(Obj_Que_Necesita__DibujarTexto) + 1
    ReDim Preserve Obj_Que_Necesita__DibujarTexto(J)
    Set Obj_Que_Necesita__DibujarTexto(J) = OB
End Sub
