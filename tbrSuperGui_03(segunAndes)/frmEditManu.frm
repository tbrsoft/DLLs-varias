VERSION 5.00
Begin VB.Form frmEditManu 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "GetFile"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   7140
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "guardar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   8370
      Width           =   1050
   End
   Begin VB.TextBox txTEXTO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   18
      Text            =   "label"
      Top             =   2820
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "como"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CheckBox chk_alphaHabilitado 
      Caption         =   "alphaHabilitado"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   6810
      Width           =   2175
   End
   Begin VB.TextBox txt_AlphaCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      TabIndex        =   11
      Text            =   "0"
      Top             =   6360
      Width           =   555
   End
   Begin VB.ComboBox cmbAlignV 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditManu.frx":0000
      Left            =   150
      List            =   "frmEditManu.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1770
      Width           =   1965
   End
   Begin VB.ComboBox cmbAlignH 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditManu.frx":0038
      Left            =   150
      List            =   "frmEditManu.frx":0045
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   1965
   End
   Begin VB.CheckBox chkEstirable 
      Caption         =   "Estirable"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   2190
      Width           =   1665
   End
   Begin VB.Frame frRect 
      Caption         =   "Rect"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2025
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Text            =   "0"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Text            =   "0"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Text            =   "0"
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtRect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Text            =   "0"
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   2550
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_FontColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FontColor"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   2145
   End
   Begin VB.Label lbl_FontColorSel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FontColorSel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   5430
      Width           =   2145
   End
   Begin VB.Label Lbl_FontColorUnSel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FontColorUnSel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   4980
      Width           =   2145
   End
   Begin VB.Label lblAphaCant 
      Caption         =   "AlphaCantidad"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   6390
      Width           =   1605
   End
   Begin VB.Label lbl_AlpahColorLong 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AlpahColorLong"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   2145
   End
   Begin VB.Label lbl_FntColorContorno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FntColorContorno"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   4530
      Width           =   2145
   End
   Begin VB.Label cmdFont 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FONT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3330
      Width           =   2145
   End
End
Attribute VB_Name = "frmEditManu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O As objFULL 'ObjFullQueRepresenta

Private Sub chkEstirable_Click()
    O.oSimple.Estirable = CBool(chkEstirable.Value)
End Sub

Private Sub cmbAlignH_Click()
    O.oSimple.AlignementH = cmbAlignH.ListIndex
End Sub

Private Sub cmbAlignV_Click()
    O.oSimple.AlignementV = cmbAlignV.ListIndex
End Sub

'basico para cargar archivo externo
Private Function GetFileCMDLG(dlgTitle As String, sFilter As String) As String
    Dim CM As New CommonDialog
    CM.DialogTitle = dlgTitle
    CM.Filter = sFilter
    CM.ShowOpen
    GetFileCMDLG = CM.FileName
End Function

Private Sub cmdAddFile_Click(Index As Integer)
    'todos los archivos que necesite la clase actual se carga al inicio
    'cuando lo cargo uso el tag
    'ejemplo para cambiar el pngBoton (la imagen)
    'cmd.tag ="Agregar imagen PngBoton\\Imagenes PNG|*.png\\pngUnSel)
    
    Dim S As String, SP() As String
    S = cmdAddFile(Index).Tag
    SP = Split(S, "\\")
    
    Dim dialogTIT As String
    Dim dialogFilter As String
    Dim sIdArchivo As String
    
    dialogTIT = SP(0)
    dialogFilter = SP(1)
    sIdArchivo = SP(2)
    'usados hasta ahora
    'PNGMarcoDisco -> discos manager
    'PNGUnSel -> pngBoton
    
    Dim F As String
    F = GetFileCMDLG(SP(0), SP(1))
    
    If F = "" Then Exit Sub 'si no elige nada se va
    
    Select Case LCase(sIdArchivo)
        Case LCase("PNGMarcoDisco")
            Dim DM As New clsDiscoManager
            Set DM = O.oManu
            DM.SetPNGMarcoDisco F
            'no alcanza con esto, debe ser parte de la comeccion de archivos fiMG
            O.getFIMG.AddFileByPath F, , "PNGMarcoDisco"
            
            'SOLO A LOS FINES DE QUE LO CARGUE Y LEA EL WI Y HE PARA QUE SE GRABEN OK!
            DM.IniciarPNGs F
            
        Case LCase("PngUnSel")
            
            Dim PB As New clsPNGBoton
            Set PB = O.oManu
            PB.SetPNGUnSel F
            'no alcanza con esto, debe ser parte de la comeccion de archivos fiMG
            O.getFIMG.AddFileByPath F, , "PngUnSel"
            
            'SOLO A LOS FINES DE QUE LO CARGUE Y LEA EL WI Y HE PARA QUE SE GRABEN OK!
            PB.IniciarPNGs F
            
    End Select
    
    
            
    
    
End Sub

Private Sub cmdFont_Click()
    Dim c As New CommonDialog
    'que cargue los valores que estan
    c.FontName = cmdFont.Font.Name
    c.FontSize = cmdFont.Font.Size
    c.Bold = cmdFont.Font.Bold
    
    c.ShowFont
    
    'ahora que actualice si hubo cambios
    cmdFont.Font.Name = c.FontName
    cmdFont.Font.Size = c.FontSize
    cmdFont.Font.Bold = c.Bold
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "FontName", c.FontName
    O.oSimple.SetProp "FontSize", CStr(c.FontSize)
    O.oSimple.SetProp "FontBold", CStr(CLng(c.Bold))
    
End Sub

Private Function hideAll()
    'ocultar todos los objetos que no necesariamente estan en todas las clases
    cmdFont.Visible = False
    lbl_FntColorContorno.Visible = False
    lbl_AlpahColorLong.Visible = False
    
    lblAphaCant.Visible = False
    txt_AlphaCantidad.Visible = False
    
    chk_alphaHabilitado.Visible = False
    lbl_FontColorSel.Visible = False
    Lbl_FontColorUnSel.Visible = False
    
    lbl_FontColor.Visible = False
    
    lblTexto.Visible = False
    txTEXTO.Visible = False
End Function

Private Sub Command1_Click()
    Dim c As New CommonDialog
    c.ShowSave
    Dim F As String
    F = c.FileName
    
    If F <> "" Then
        Dim J As Long
        J = O.Save(F)
        
        If J = 0 Then
            MsgBox "Se grabo ok"
        Else
            MsgBox "Error al grabar: " + CStr(J)
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    RefreshShow
    O.Save
End Sub

Private Sub lbl_AlpahColorLong_Click()
    Dim c As New CommonDialog
    c.RGBResult = lbl_AlpahColorLong.BackColor   'ver si anda, que el cmdlg empiece mostrando lo elegido
    c.ShowColor
    
    lbl_AlpahColorLong.BackColor = c.RGBResult
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "AlpahColorLong", CStr(c.RGBResult)
    
End Sub

Private Sub lbl_FntColorContorno_Click()
    Dim c As New CommonDialog
    c.RGBResult = lbl_FntColorContorno.BackColor 'ver si anda, que el cmdlg empiece mostrando lo elegido
    c.ShowColor
    
    lbl_FntColorContorno.BackColor = c.RGBResult
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "FntColorContorno", CStr(c.RGBResult)
    
End Sub

Private Sub lbl_FontColor_Click()
    Dim c As New CommonDialog
    c.RGBResult = lbl_FontColor.BackColor 'ver si anda, que el cmdlg empiece mostrando lo elegido
    c.ShowColor
    
    cmdFont.ForeColor = c.RGBResult
    lbl_FontColor.BackColor = c.RGBResult
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "FontColor", CStr(c.RGBResult)
    
    
End Sub

Private Sub lbl_FontColorSel_Click()
    Dim c As New CommonDialog
    c.RGBResult = lbl_FontColorSel.BackColor  'ver si anda, que el cmdlg empiece mostrando lo elegido
    c.ShowColor
    
    lbl_FontColorSel.BackColor = c.RGBResult
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "FontColorSel", CStr(c.RGBResult)
    
    
End Sub

Private Sub Lbl_FontColorUnSel_Click()
    Dim c As New CommonDialog
    c.RGBResult = Lbl_FontColorUnSel.BackColor  'ver si anda, que el cmdlg empiece mostrando lo elegido
    c.ShowColor
    
    Lbl_FontColorUnSel.BackColor = c.RGBResult
    
    'darle los valores al objFull para que se muestre con estos cambios
    O.oSimple.SetProp "FontColorUnSel", CStr(c.RGBResult)
    
    
End Sub

Private Sub txTEXTO_Change()
    O.oSimple.SetProp "TextoActual", txTEXTO.Text
    cmdFont.Caption = txTEXTO.Text
End Sub

Private Sub txtRect_Validate(Index As Integer, Cancel As Boolean)

    If IsNumeric(txtRect(Index).Text) = False Then
        Cancel = True
        Exit Sub
    End If

    Select Case Index
        Case 0
            O.oSimple.X = CLng(txtRect(0).Text)
        Case 1
            O.oSimple.Y = CLng(txtRect(1).Text)
        Case 2
            O.oSimple.W = CLng(txtRect(2).Text)
        Case 3
            O.oSimple.H = CLng(txtRect(3).Text)
    
    End Select
    
End Sub

'vuleve a cargar el objeto del manu
'cerrar y reahcer los graficos y fuentes
Public Sub RefreshShow()
    'seguiraqui QUE MANU lo revise y apruebe
    
    'refrescar todo (en realidad los objetos no deben refrescarse, es el formpadre con todos sus objetos)
    'de todas formas queda logico que un objeto se refresque cuando cambian sus propiedades
    
    O.GetPadre.INIT_GRAPH "CLOSE"
    O.GetPadre.INIT_GRAPH
    'quedeaqui, si no le hago un refresh al picture no se ve un joraca
    
End Sub

Public Sub SetObjFull(obj As objFULL)
    
    'paso el padere para poder refresacarlo a medida que cambie
    Set O = obj
    hideAll 'esconder todas las propiedades para que solo cargue las que usa
    
    'leerle todas las propiedades y cargarlas!
    'actualziarlas!!!
    O.UpdatePropiedades
    
    txtRect(0).Text = O.oSimple.X
    txtRect(1).Text = O.oSimple.Y
    txtRect(2).Text = O.oSimple.W
    txtRect(3).Text = O.oSimple.H
    
    cmbAlignH.ListIndex = O.oSimple.AlignementH
    cmbAlignV.ListIndex = O.oSimple.AlignementV
    
    chkEstirable.Value = Abs(CLng(O.oSimple.Estirable))
    
    'las demas propiedades pueden estar o no asi que se cargan segun correspondan
    Dim J As Long, PP As clsPropis
    
    'estas propiedades se van acomodando en Top para que no queden huecos vacios
    Dim LastY As Long
    LastY = chkEstirable.Top + chkEstirable.Height + 30
    
    For J = 1 To O.oSimple.GetPropCantidad
        Set PP = O.oSimple.GetPropByID(J) 'NO permite buscar en otros valores predeterminados (eso es un problema)
        Set PP = O.oSimple.GetProp(PP.NameProp) 'suena feo pero busca valores predeterminados de cada una
        
        Select Case LCase(PP.NameProp)
        
            Case "textoactual"
                lblTexto.Top = LastY: LastY = LastY + lblTexto.Height + 15
                txTEXTO.Top = LastY: LastY = LastY + txTEXTO.Height + 60
                
                txTEXTO.Text = PP.ValueProp
                txTEXTO.Visible = True
                lblTexto.Visible = True
                
            Case "fontname"
                cmdFont.Top = LastY: LastY = LastY + cmdFont.Height + 60
                
                cmdFont.Font.Name = PP.ValueProp
                cmdFont.Visible = True
                
            Case "fontsize"
                cmdFont.Font.Size = CLng(PP.ValueProp)
            
            Case "fontbold"
                cmdFont.Font.Bold = CBool(CLng(PP.ValueProp))
                
            Case "fontcolor"
                lbl_FontColor.Top = LastY: LastY = LastY + lbl_FontColor.Height + 60
                
                lbl_FontColor.Visible = True
                cmdFont.ForeColor = CLng(PP.ValueProp)
                lbl_FontColor.BackColor = CLng(PP.ValueProp)
            
            Case LCase("FntColorContorno")
                lbl_FntColorContorno.Top = LastY: LastY = LastY + lbl_FntColorContorno.Height + 60
                
                lbl_FntColorContorno.BackColor = CLng(PP.ValueProp)
                lbl_FntColorContorno.Visible = True
            
            Case LCase("AlpahColorLong")
                lbl_AlpahColorLong.Top = LastY: LastY = LastY + lbl_AlpahColorLong.Height + 60
                
                lbl_AlpahColorLong.BackColor = CLng(PP.ValueProp)
                lbl_AlpahColorLong.Visible = True
            
            Case LCase("AlphaCantidad")
                txt_AlphaCantidad.Top = LastY
                lblAphaCant.Top = LastY: LastY = LastY + lbl_AlpahColorLong.Height + 60
                
                txt_AlphaCantidad.Text = PP.ValueProp
                lblAphaCant.Visible = True
                txt_AlphaCantidad.Visible = True
            
            Case LCase("alphaHabilitado")
                chk_alphaHabilitado.Top = LastY: LastY = LastY + chk_alphaHabilitado.Height + 60
                
                chk_alphaHabilitado.Value = Abs(CLng(PP.ValueProp))
                chk_alphaHabilitado.Visible = True

            Case LCase("FontColorUnSel")
                Lbl_FontColorUnSel.Top = LastY: LastY = LastY + Lbl_FontColorUnSel.Height + 60
                
                Lbl_FontColorUnSel.BackColor = CLng(PP.ValueProp)
                Lbl_FontColorUnSel.Visible = True
            
            Case LCase("FontColorSel")
                lbl_FontColorSel.Top = LastY: LastY = LastY + lbl_FontColorSel.Height + 60
                
                lbl_FontColorSel.BackColor = CLng(PP.ValueProp)
                lbl_FontColorSel.Visible = True
        
        End Select
    
    Next J
    
    'agregar los botones que cargan archivos externos segun el tipo de objeto
    UnLoadCmdFiles
    
    'cada clase con la cantidad de archivos que corresponda
    If O.Tipo = en_clsDiscoManager Then
        Load cmdAddFile(1)
        cmdAddFile(1).Caption = "PNGMarcoDisco"
        cmdAddFile(1).Tag = "Agregar marco disco\\Imagenes PNG|*.png\\PNGMarcoDisco"
    End If
    
    If O.Tipo = en_clsPNGBoton Then
        Load cmdAddFile(1)
        cmdAddFile(1).Caption = "PngUnSel"
        cmdAddFile(1).Tag = "Agregar imagen PngBoton\\Imagenes PNG|*.png\\pngUnSel"
    End If
    
    'acomodar los botones en TOP y hacerlos visibles
    For J = 1 To cmdAddFile.Count - 1
        cmdAddFile(J).Top = LastY: LastY = LastY + cmdAddFile(J).Height + 60
        cmdAddFile(J).Visible = True
    Next J
    
    
    
    
    
    '**********************************************************************
    '**********************************************************************
    LastY = LastY + Command2.Height + 180
    Command1.Top = LastY
    Command2.Top = LastY: LastY = LastY + Command2.Height + 60
    '**********************************************************************
    '**********************************************************************
    LastY = LastY + Command2.Height + 120
    Me.Height = LastY
    '**********************************************************************
    '**********************************************************************
End Sub

Private Sub UnLoadCmdFiles()
    Dim J As Long
    For J = 1 To cmdAddFile.Count - 1
        Unload cmdAddFile(J)
    Next J
End Sub
