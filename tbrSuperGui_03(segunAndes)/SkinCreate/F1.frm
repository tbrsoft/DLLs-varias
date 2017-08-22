VERSION 5.00
Begin VB.Form F1 
   Caption         =   " "
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9210
   Icon            =   "F1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VS 
      Height          =   3285
      Left            =   6420
      TabIndex        =   4
      Top             =   270
      Width           =   315
   End
   Begin VB.HScrollBar HS 
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdPadre 
      Caption         =   "Nuevo_Padre"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.PictureBox PIC 
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   1950
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   450
      Width           =   4395
      Begin VB.PictureBox PicPadre 
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Index           =   0
         Left            =   930
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Menu mnSkin 
      Caption         =   "Skin"
      Begin VB.Menu mnNewSkin 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnOpenSkin 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnCloseSkin 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu mnGrabarSkin 
         Caption         =   "Grabar"
      End
      Begin VB.Menu mnGrabarComoSkin 
         Caption         =   "Grabar como ..."
      End
      Begin VB.Menu sep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnReutils 
         Caption         =   "Reutils"
      End
      Begin VB.Menu mnsep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnForm 
      Caption         =   "Formulario"
      Begin VB.Menu mnAddForm 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnEditForm 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnKillForm 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu mnObj 
      Caption         =   "Objeto"
      Begin VB.Menu mnAddObj 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnAddObj_Reutil 
         Caption         =   "Agregar from Reutil"
      End
      Begin VB.Menu mnEditObj 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnKillObj 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private G As tbrSuperGui_3.clsGUI 'gui abierto
Private PadreActual As ObjFullPadre 'el padre sobre el que se esta trabajando en cada momento
Private ixPadreActual As Long 'indice del picturBox del padre actual
'ultimo objeto elegido para editar
Private ObjElegido As tbrSuperGui_3.objFULL

Private Sub cmdPadre_Click(Index As Integer)
    SelectPadre CLng(Index) 'mostrar el elegido
    UpdateScrolls
End Sub

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    TERR.FileLog = AP + "RegSkin.log"
    TERR.LargoAcumula = 800
    
    If FSO.FolderExists(AP + "SKINS") = False Then
        FSO.CreateFolder AP + "SKINS"
    End If
    
    Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next

    'Me.ScaleMode = 3
    cmdPadre(0).Left = 4
    cmdPadre(0).Top = 4
    
    PIC.Left = 4
    PIC.Top = cmdPadre(0).Height + 10
    PIC.Width = (Me.Width / 15) - 15 - VS.Width
    PIC.Height = (Me.Height / 15) - PIC.Top - 50 - HS.Height
    
    VS.Top = PIC.Top
    VS.Left = PIC.Left + PIC.Width
    VS.Height = PIC.Height
    
    HS.Top = PIC.Top + PIC.Height
    HS.Left = PIC.Left
    HS.Width = PIC.Width
    
    UpdateScrolls
    
End Sub

Private Sub UpdateScrolls()
    'Ver si deben estar enabled o no!
    VS.Enabled = (PicPadre(ixPadreActual).Height > PIC.Height)
    HS.Enabled = (PicPadre(ixPadreActual).Width > PIC.Width)
    
    Dim Dif As Single
    If VS.Enabled Then 'ver que movimiento tiene
        Dif = PicPadre(ixPadreActual).Height - PIC.Height
        VS.LargeChange = Dif / 4
        VS.SmallChange = Dif / 16
        'seguiraqui, asegurarse que ninguno de cero
        VS.Min = 0
        VS.Max = CLng(Dif)
        VS.Value = Abs(PicPadre(ixPadreActual).Top)
        
    End If
    
    If HS.Enabled Then 'ver que movimiento tiene
        Dif = PicPadre(ixPadreActual).Width - PIC.Width
        HS.LargeChange = Dif / 4
        HS.SmallChange = Dif / 16
        
        HS.Min = 0
        HS.Max = CLng(Dif)
        HS.Value = Abs(PicPadre(ixPadreActual).Left)
        
    End If
End Sub

'abrir uno usado o nuevo
Private Sub OpenSkin(sFile As String)
    On Local Error GoTo ErrOPP
    TERR.Anotar "aao", sFile
    
    'si sFile esta vacio lo creo
    Dim isNew As Boolean
    If sFile = "" Then
        isNew = True
    Else
        If FSO.FileExists(sFile) = False Then
            isNew = True
        Else
            isNew = False
        End If
    End If
    
    TERR.Anotar "aaa", isNew
    Set G = New tbrSuperGui_3.clsGUI
    G.SetParaPantallaPixeles 1024, 768
    G.SetPathLog AP + "RegGUI3.log"
    
    TERR.Anotar "aab"
    If isNew Then
    
        'obligatoriamente un padre por lo menos
        Dim tmpPadre As ObjFullPadre
        Set tmpPadre = G.MNG.AddPadre("Form1")
        'darle un tamaño generico
        tmpPadre.GetSgoInterno.X = 0
        tmpPadre.GetSgoInterno.Y = 0
        tmpPadre.GetSgoInterno.W = 300
        tmpPadre.GetSgoInterno.H = 300
        'GRABARLO!
        Dim PAD As String
        PAD = InputBox("Inserte nombre del skin nuevo", "Nuevo skin", "sk8_name")
        TERR.Anotar "aau", PAD
        G.SaveFile AP + "SKINS\" + PAD + ".SK8" 'skin version 8
        TERR.Anotar "aav"
        
    Else
    
        Dim FolSkin As String 'carpeta para trabajar con el skin
        FolSkin = AP + "SKINs\" + FSO.GetBaseName(sFile)
        TERR.Anotar "aaq", FolSkin
        If FSO.FolderExists(FolSkin) Then
            TERR.Anotar "aar"
            FSO.DeleteFolder FolSkin, True
        End If
        FSO.CreateFolder FolSkin
        TERR.Anotar "aas"
        G.LoadFile sFile, FolSkin
        TERR.Anotar "aat"
        
    End If
    
    TERR.Anotar "aam"
    'MOSTRARLO
    Loadpadres 'ya carga el primer padre en pantalla
    
    TERR.Anotar "aan"
    UpdateScrolls
    
    
    Exit Sub
    
ErrOPP:
    TERR.AppendLog "errOP902", TERR.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub HS_Change()
    PicPadre(ixPadreActual).Left = -HS.Value
End Sub

Private Sub mnAddForm_Click()
    Dim PAD As String
    PAD = InputBox("Inserte nombre del formulario nuevo", "Nuevo formulario para skin", "Nuevo_Padre")
    If PAD = "" Then Exit Sub 'CANCELO
    
    Dim tmpPadre As tbrSuperGui_3.ObjFullPadre
    Set tmpPadre = G.MNG.AddPadre(PAD)
    tmpPadre.GetSgoInterno.X = 0
    tmpPadre.GetSgoInterno.Y = 0
    tmpPadre.GetSgoInterno.W = 400
    tmpPadre.GetSgoInterno.H = 400
    
    Loadpadres -1 '-1 es el ultimo
    
    UpdateScrolls
End Sub

Private Sub mnAddObj_Click()
    Dim OG As New tbrSuperGui_3.objFULL
    
    Set ObjElegido = OG
    UpdateNameMenusObj
End Sub

Private Sub mnAddObj_Reutil_Click()
    PadreActual.AppendSGO frmReUsar.ObjElegido
    RefreshPadre
End Sub

Private Sub mnEditForm_Click()
    PadreActual.GuiEdit    'mostrar las propiedades del form!
End Sub

Private Sub mnGrabarSkin_Click()
    If FSO.FileExists(G.path) Then
        G.SaveFile "" 'con el nombre que tenia antes
    Else
        MsgBox "No hay donde grabar !!"
    End If
End Sub

Private Sub mnKillForm_Click()
    'SEGUIRAQUI NO ME INTERESA eliminar por ahora
End Sub

Private Sub mnKillObj_Click()
    'seguiraqui
End Sub

Private Sub mnNewSkin_Click()
    OpenSkin ""
End Sub

Private Sub mnOpenSkin_Click()
    Dim CM As New CommonDialog
    CM.InitDir = AP + "SKINS"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    OpenSkin F
    
End Sub

Private Sub mnReutils_Click()
    frmReUsar.CargarTodo
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

'mostrar los padres-formularios en la parte superior
Private Sub Loadpadres(Optional selectLast As Long = 1)
    On Local Error GoTo ErrLP9002
    TERR.Anotar "aad"
    'descargar los anteriores
    UnloadPadres
    
    Dim J As Long
    For J = 1 To G.MNG.GetPadresMaxID
        TERR.Anotar "aae", J, G.MNG.GetPadresByID(J).sName
        Load cmdPadre(J)
        
        Dim LeftBt As Long
        If J = 1 Then
            LeftBt = PIC.Left
        Else
            LeftBt = cmdPadre(J - 1).Left + cmdPadre(J - 1).Width
        End If
        
        cmdPadre(J).Left = LeftBt
        cmdPadre(J).Top = PIC.Top - cmdPadre(J).Height - 2
        cmdPadre(J).Caption = G.MNG.GetPadresByID(J).sName
        cmdPadre(J).Visible = True
        
        TERR.Anotar "aaf"
        Load PicPadre(J)
        PicPadre(J).AutoRedraw = True
        PicPadre(J).BorderStyle = 0
        PicPadre(J).Left = G.MNG.GetPadresByID(J).GetSgoInterno.X
        PicPadre(J).Top = G.MNG.GetPadresByID(J).GetSgoInterno.Y
        PicPadre(J).Width = G.MNG.GetPadresByID(J).GetSgoInterno.W
        PicPadre(J).Height = G.MNG.GetPadresByID(J).GetSgoInterno.H
        
        G.MNG.GetPadresByID(J).sHDC = PicPadre(J).hdc
        
        'si tiene imagen de fondo cargarla
        Dim iFondo As String
        iFondo = G.MNG.GetPadresByID(J).pathImgFondo
        If FSO.FileExists(iFondo) Then
            PicPadre(J).PaintPicture LoadPicture(G.MNG.GetPadresByID(J).pathImgFondo), 0, 0, PicPadre(J).Width, PicPadre(J).Height
        End If
        'SE HACE VISIBLE solo cuando se elige !
        
    Next J
    
    TERR.Anotar "aag"
    'generalmente muestra el primero salvo que le pidan otro (o el ultimo con -1)
    If selectLast = -1 Then 'quiere mostrar el ultimo de la coleccion
        selectLast = PicPadre.Count - 1
    End If
    
    SelectPadre selectLast 'siemrpe hay un padre al menos
    
    Exit Sub
ErrLP9002:
    TERR.AppendLog "errLP9002", TERR.ErrToTXT(Err)
End Sub

'mostrar el pic dºel padre elegido
Private Sub SelectPadre(I As Long)
    
    TERR.Anotar "aaj", cmdPadre.Count, PicPadre.Count
    Dim J As Long
    For J = 1 To cmdPadre.Count - 1
        If I = J Then
            PicPadre(J).Visible = True
            cmdPadre(J).Font.Bold = True
            cmdPadre(J).Font.Size = 12
            Set PadreActual = G.MNG.GetPadre(cmdPadre(J).Caption) 'parece negrada ...
            UpdateNameMenusForm
            ixPadreActual = I
            RefreshPadre 'hace los initgraph
            PicPadre(J).Refresh 'seguiraqui esto no estaba y el label por lo menos andaba ok
        Else
            PicPadre(J).Visible = False
            cmdPadre(J).Font.Bold = False
            cmdPadre(J).Font.Size = 10
        End If
    Next J
    
    TERR.Anotar "aak"
End Sub

Private Sub UnloadPadres()
    TERR.Anotar "aah", cmdPadre.Count, PicPadre.Count
    Dim J As Long
    For J = 1 To cmdPadre.Count - 1
        Unload cmdPadre(J)
    Next J
    
    For J = 1 To PicPadre.Count - 1
        Unload PicPadre(J)
    Next J
    
    TERR.Anotar "aal"
    
End Sub

Private Sub RefreshPadre()
    PadreActual.INIT_GRAPH "CLOSE"
    PadreActual.INIT_GRAPH
    
    PicPadre(ixPadreActual).Refresh
    
    UpdateScrolls
End Sub

Private Sub PicPadre_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ver donde carajo estoy
    Dim J As Long, J2 As Long
    J2 = PadreActual.GetSgoMaxID
    
    Dim ixSEL As Long 'indice del objeto elegido
    ixSEL = -1
    For J = 1 To J2
        Dim obX As Long, obY As Long, obW As Long, obH As Long
        obX = PadreActual.GetSgoByID(J).oSimple.X
        obY = PadreActual.GetSgoByID(J).oSimple.Y
        obH = PadreActual.GetSgoByID(J).oSimple.H
        obW = PadreActual.GetSgoByID(J).oSimple.W
        
        If X >= obX And X <= obX + obW Then
            If Y >= obY And Y <= obY + obH Then
                'se hizo click en este !!
                Set ObjElegido = PadreActual.GetSgoByID(J)
                UpdateNameMenusObj
                ObjElegido.GuiEdit PadreActual
                ixSEL = J
                Exit For
            End If
        End If
    Next J
    
    If ixSEL = -1 Then 'no le pego a ningun objeto!
        PadreActual.GuiEdit   'mostrar las propiedades del form!
        Exit Sub
    End If
    
    'ademas de abrir el editor ver si necesita algo mas (SEGUN OBJETO DEL MANU)
    If PadreActual.GetSgoByID(ixSEL).Tipo = en_clsTemasManager Then
        Dim OB As clsTemasManager
        Set OB = PadreActual.GetSgoByID(ixSEL).oManu
        OB.DoClick_GetElementoIndex CLng(X), CLng(Y)
    End If
End Sub

Private Sub VS_Change()
    PicPadre(ixPadreActual).Top = -VS.Value
End Sub

Private Sub UpdateNameMenusObj()
    mnEditObj.Caption = "Editar " + ObjElegido.oSimple.SGOName
    mnKillObj.Caption = "Eliminar " + ObjElegido.oSimple.SGOName
End Sub

Private Sub UpdateNameMenusForm()
    mnEditForm.Caption = "Editar " + PadreActual.sName
    mnKillForm.Caption = "Eliminar " + PadreActual.sName
End Sub

