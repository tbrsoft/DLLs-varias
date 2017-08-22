VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTraductor 
   Caption         =   "Traductor de documentos de Idioma"
   ClientHeight    =   9420
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   6210
      TabIndex        =   11
      Top             =   510
      Width           =   2055
   End
   Begin VB.CheckBox chkFaltantes 
      Caption         =   "Mostrar solo faltantes"
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Top             =   0
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
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
      Height          =   855
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   0
      Width           =   5805
   End
   Begin VB.Frame Frame1 
      Height          =   4545
      Left            =   30
      TabIndex        =   1
      Top             =   2220
      Width           =   8115
      Begin VB.CommandButton Command1 
         Caption         =   "KILL"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1200
         Left            =   810
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Form1.frx":0442
         Top             =   1860
         Width           =   5175
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Aceptar este y seguir"
         Height          =   1155
         Left            =   90
         TabIndex        =   4
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox txtORIG 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   810
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   420
         Width           =   5175
      End
      Begin VB.TextBox txtExplic 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   3270
         Width           =   5205
      End
      Begin VB.Label Label2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Texto original"
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   8
         Top             =   210
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "Texto traducido"
         Height          =   315
         Index           =   1
         Left            =   780
         TabIndex        =   7
         Top             =   1650
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Explicacion del creador"
         Height          =   315
         Index           =   2
         Left            =   810
         TabIndex        =   6
         Top             =   3060
         Width           =   1995
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   990
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1879
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Line Line1 
      X1              =   -90
      X2              =   8190
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar texto"
      Height          =   315
      Index           =   3
      Left            =   6240
      TabIndex        =   12
      Top             =   300
      Width           =   2055
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir..."
      End
      Begin VB.Menu mnRECs 
         Caption         =   "Recientes"
         Begin VB.Menu mnRECENTS 
            Caption         =   "ListaRec"
            Index           =   0
         End
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnuGuardarComo 
         Caption         =   "Guardar como..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnJoin 
         Caption         =   "Join IanSource+Ian"
      End
      Begin VB.Menu mn2342 
         Caption         =   "-"
      End
      Begin VB.Menu mnOLD 
         Caption         =   "Cadenas Viejas y feas"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
   End
   Begin VB.Menu mnLanSources 
      Caption         =   "LanSources"
      Begin VB.Menu mnAddChainLSource 
         Caption         =   "Agregar cadena"
      End
   End
End
Attribute VB_Name = "frmTraductor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private path As String
Private estaGrabado As Boolean 'para quie no salga sin grabar
Dim cCadenas As Long 'total de cadenas
Dim cCadenasT As Long 'total de cadenas traducidas

Dim Aux2() As String 'ultimo archivo abierto

Dim AP As String, sFileRecent As String 'archivo con los ultimos usados
Dim FSO As New Scripting.FileSystemObject

Private Sub chkFaltantes_Click()
    LoadAux2
End Sub

Private Sub Command1_Click()
    If MsgBox("Seguro que elimina" + vbCrLf + lvw.SelectedItem.Text, vbQuestion + vbYesNo) = vbYes Then
        lvw.ListItems.Remove lvw.SelectedItem.Index
        'ademas para que se elimine de verdad debe salir de aux
        DeleteAux2 CLng(Label2)
    End If
End Sub

Private Sub Form_Load()
    estaGrabado = True
    
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    sFileRecent = AP + "recents.txt"
    Dim R() As String
    R = ListRecents
    
    If UBound(R) = 0 Then
        mnRECs.Visible = False
    Else
        Dim j As Long
        For j = 1 To UBound(R)
            
            If j > 1 Then Load mnRECENTS(j - 1)
            mnRECENTS(j - 1).Caption = R(j)
            mnRECENTS(j - 1).Visible = True
        
        Next
    End If
    
End Sub

Private Sub Form_Resize()
    lvw.Width = Me.Width - lvw.Left - 200
    Text1.Left = lvw.Left
    'Text1.Width = lvw.Width
    txt.Width = Me.Width - txt.Left - 200
    txtExplic.Width = txt.Width
    txtORIG.Width = txt.Width
    
    Frame1.Width = Me.Width - Frame1.Left - 100
    Frame1.Top = Me.Height - Frame1.Height - 1200
    
    On Local Error Resume Next
    
    lvw.ColumnHeaders(1).Width = (lvw.Width / 3) - 190
    lvw.ColumnHeaders(2).Width = (lvw.Width / 3) - 190
    lvw.ColumnHeaders(3).Width = (lvw.Width / 3) - 190
    
    If LCase(Right(path, 9)) = "lansource" Then
        lvw.Height = Frame1.Top + Frame1.Height
    Else
        lvw.Height = Frame1.Top - 260 - Line1.Y1
    End If
    
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    
    txt.Text = Item.ListSubItems(1).Text
    txtORIG.Text = Item.Text
    txtExplic.Text = Item.ListSubItems(2).Text
    
    'mostrar en colores diferentes segun si se tradujo o no
    If txtORIG.Text <> txt.Text Then
        txt.ForeColor = &H8000&
    Else
        txt.ForeColor = vbBlack
    End If
    
    Label2.Caption = Item.ListSubItems(3).Text

End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton Then Me.PopupMenu mnuPopUp
End Sub

Private Sub mnAddChainLSource_Click()

    MsgBox "No funciona asi. Siempre se graba cono LAN (creo)" + vbCrLf + _
        "Si quiere agregar una entre al lansource y agreguela manualmente (notese que son dobles cada cadenas)"
    

    'Dim LI As ListItem
    'Set LI = lvw.ListItems.Add(, , "<Nuevo>")
    'LI.ListSubItems.Add , , "<Nuevo>"
    'LI.ListSubItems.Add , , "<Nuevo>"
End Sub

Private Sub mnJoin_Click()
    'si el programador agrega, corrige o quita cadenas EL TRADUCTOR SE CAGA y queda regalado
    'el iansource es igual que el ian pero tiene un campo mas que es el contexto (en realidad por error el ian tambien lo tiene)
    
    'entonces cuando el programador le da el nuevo iansource actualizado el traductor para no perder su
    'trabajo fuciona su ian de trabajo con el iansource recibido
    
    Dim I1 As String 'iansource nuevo
    Dim I2 As String 'ian en proceso del traductor
    
    Dim cd As New CommonDialog
    cd.Filter = "Indique el NUEVO iansource recibido (*.lanSource)|*.lansource"
    
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    I1 = cd.FileName
    
    cd.Filter = "Indique su IAN en uso (*.lan)|*.lan"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    I2 = cd.FileName
    
    'ahora debo pasar todas las traducciones del ian al iansource
    'ver que queda huerfano en el ian y guardarselo al paquetazo del traductor para reutilizar
    
    'meto cada uno en una matriz manejable
    Dim M1() As String, M2() As String
    'loaddictionary carga el AUX2..
    LoadDictionary I1
    M1 = Aux2
    
    LoadDictionary I2
    M2 = Aux2
    
    Dim I As Long, j As Long
    Dim SP1() As String, SP2() As String
    
    'contadores para mostrar
    Dim CadenasEncontradas As Long
    Dim CadenasNuevas As Long
    Dim CadenasHuerfanas As Long
    
    CadenasEncontradas = 0
    CadenasNuevas = 0
    CadenasHuerfanas = 0
    
    For I = 1 To UBound(M1)
        SP1 = Split(M1(I), Chr(6))
        
        For j = 1 To UBound(M2)
            SP2 = Split(M2(j), Chr(6))
            
            If Trim(LCase(SP1(1))) = Trim(LCase(SP2(1))) Then
                'bingo!
                CadenasEncontradas = CadenasEncontradas + 1
                SP1(2) = SP2(2)
                'la marco como usada y
                'actualizo la matriz para que ya los descarte
                M2(j) = SP2(0) + Chr(6) + "" + Chr(6) + "" + Chr(6) + ""
                Exit For 'para no perder tiempo
            End If
            
            If j = UBound(M2) Then CadenasNuevas = CadenasNuevas + 1
            
        Next j
        
        'actualizo la matriz para que se grabe joia
        M1(I) = SP1(0) + Chr(6) + SP1(1) + Chr(6) + SP1(2) + Chr(6) + SP1(3)
        
    Next I
    
    'meto en M1 lo quedo sin usar en M2
    For j = 1 To UBound(M2)
        SP2 = Split(M2(j), Chr(6))
        
        'ademas de no haber sido usado deberia estar traducido, de no ser asi no tiene sentido.
        'los huerfanos sirven si contienen texto traducido
        
        If (SP2(1) <> "" And SP2(2) <> "") And (SP2(1) <> SP2(2)) Then
            
            I = UBound(M1) + 1
            ReDim Preserve M1(I)
            '!!! indica que no se debe usar para traducir y se debe marcar para usar y luego eliminar
            M1(I) = CStr(I) + Chr(6) + "HUERFANO: " + SP2(1) + Chr(6) + SP2(2) + Chr(6) + SP2(3)
            CadenasHuerfanas = CadenasHuerfanas + 1
            
        End If
    Next j
    
    
    'listo, M1 tiene lo que necesito
    'genero el texto a grabar
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(I1 + ".JOIN.LAN", True)
    
        For j = 1 To UBound(M1)
        
            SP1 = Split(M1(j), Chr(6))
            TE.Write SP1(1) + Chr(31) + SP1(2) + Chr(31) + SP1(3) + Chr(30)
        Next j
            
    TE.Close
    
    MsgBox "Se ha finalizado!!" + vbCrLf + _
            "Se grabo para que siga traduciendo en:" + vbCrLf + _
            I1 + ".JOIN.LAN" + vbCrLf + _
            "   Resumen:" + vbCrLf + _
            "      Cadenas Identicas:" + CStr(CadenasEncontradas) + vbCrLf + _
            "      Cadenas Nuevas:" + CStr(CadenasNuevas) + vbCrLf + _
            "      Cadenas Huerfanas:" + CStr(CadenasHuerfanas)
            
    lvw.ListItems.Clear
    
End Sub

Private Sub mnOLD_Click()
    Dim cd As New CommonDialog
    cd.Filter = "Fuente de idioma (*.lanSource)|*.lansource" + _
                "|Archivo de idioma (*.lan)|*.lan" + _
                "|Archivo de idioma (*.idm)|*.idm"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    
    Dim TE As TextStream, TX As String
    
    Set TE = FSO.OpenTextFile(cd.FileName, ForReading, False)
        TX = TE.ReadAll
    TE.Close
    
    Load frmOLD
    
    Dim SP() As String, SP2() As String
    SP = Split(TX, Chr(30))
    
    Dim H As Long
    For H = 0 To UBound(SP)
        SP2 = Split(SP(H), Chr(31))
        If InStr(SP(H), Chr(31)) > 0 Then
            'si son iguales noooo
            If SP2(0) <> SP2(1) Then frmOLD.List1.AddItem SP2(1)
        End If
    Next H
    
    frmOLD.Show
    
End Sub

Private Sub mnRECENTS_Click(Index As Integer)
    Abrir mnRECENTS(Index).Caption
    estaGrabado = True
End Sub

Private Sub mnuAbrir_Click()

    If estaGrabado = False Then
        If MsgBox("No ha grabado los cambios, ¿Desea abrir otro archivo de todas formas?", vbYesNo) = no Then
            Exit Sub
        End If
    End If

    Dim cd As New CommonDialog
    cd.Filter = "Fuente de idioma (*.lanSource)|*.lansource" + _
                "|Archivo de idioma (*.lan)|*.lan" + _
                "|Archivo de idioma (*.idm)|*.idm"
    cd.ShowOpen
    Dim f As String
    f = cd.FileName
    If f = "" Then Exit Sub
    
    Abrir f
    
    lvw.ColumnHeaders(1).Width = (lvw.Width / 3) - 190
    lvw.ColumnHeaders(2).Width = (lvw.Width / 3) - 190
    lvw.ColumnHeaders(3).Width = (lvw.Width / 3) - 190
    
    estaGrabado = True
    AddRedcent f
    
    If Frame1.Visible = False Then
        Text1.Text = "Para comenzar a traducir use 'Guardar como...' Este es un archivo de definicion de lenguaje" + _
            vbCrLf + Text1.Text
    End If
    
End Sub

Private Sub Abrir(sFile As String)
    LoadDictionary sFile
    path = sFile
    cCadenas = lvw.ListItems.Count
    If cCadenas = 0 Then cCadenas = 1
    
    Text1.Text = "INFO DEL ARCHIVO:" + vbCrLf + _
        "Cadenas traducidas:" + CStr(cCadenasT) + " de " + CStr(cCadenas) + " (" + CStr(Round((cCadenasT / cCadenas) * 100, 3)) + "%)"
End Sub

Private Sub cmdSiguiente_Click()
    On Error Resume Next
    lvw.SelectedItem.ForeColor = vbBlue
    If lvw.SelectedItem.ListSubItems(1).Text <> txt.Text Then
        lvw.SelectedItem.ListSubItems(1).Text = txt.Text
        
        'EL AUX2 es el que va a grabar NO PERDERLO !!
        ReplacePartAUX2 CLng(Label2.Caption), 2, txt.Text
        
        cCadenasT = cCadenasT + 1
        'actualziar el texto de arriba
        Text1.Text = "INFO DEL ARCHIVO:" + vbCrLf + _
        "Cadenas traducidas:" + CStr(cCadenasT) + " de " + CStr(cCadenas) + " (" + CStr(Round((cCadenasT / cCadenas) * 100, 3)) + "%)"
        
    End If
    Set lvw.SelectedItem = lvw.ListItems(lvw.SelectedItem.Index + 1)
    
    lvw_ItemClick lvw.SelectedItem 'para que lo pase todo
    
    txt.SetFocus
    'txt = lvw.SelectedItem.ListSubItems(1).Text
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
    
    estaGrabado = False
End Sub

Private Sub DeleteAux2(ix As Long)
    Dim SP() As String
    Dim j As Long
    For j = 1 To UBound(Aux2)
        SP = Split(Aux2(j), Chr(6))
        If CLng(SP(0)) = ix Then
            'eliminar este y parchar el espacio vacio
            Dim K As Long
            For K = j To UBound(Aux2) - 1
                Aux2(K) = Aux2(K + 1)
            Next K
            
            ReDim Preserve Aux2(UBound(Aux2) - 1)
            Exit For
        End If
    
    Next j
End Sub

Private Sub ReplacePartAUX2(idAux2 As Long, idPart As Long, CHG As String)
    Dim SP() As String
    SP = Split(Aux2(idAux2), Chr(6))
    
    SP(idPart) = CHG
    
    Aux2(idAux2) = SP(0) + Chr(6) + SP(1) + Chr(6) + SP(2) + Chr(6) + SP(3)
    
End Sub

Private Sub mnuAgregar_Click()
    Dim LI As ListItem
    Set LI = lvw.ListItems.Add(, , "<Nuevo>")
    LI.ListSubItems.Add , , "<Nuevo>"
End Sub

Private Sub mnuGuardar_Click()

    'sacar los filtros para que grabe todo!!!
    chkFaltantes.Value = 0
    txtFilter.Text = ""

    If path <> "" Then
        Dim LI As ListItem
        Dim aux As String
        For Each LI In lvw.ListItems
            aux = aux + LI.Text + Chr$(31) + LI.ListSubItems(1).Text + Chr$(31) + LI.ListSubItems(2).Text + Chr$(30)
        Next
        EscribirArchivo path, aux
    Else
        mnuGuardarComo_Click
    End If
End Sub

Private Sub mnuGuardarComo_Click()

    Dim cd As New CommonDialog
    cd.flags = cdlOFNOverwritePrompt
    cd.DefaultExt = ".lan"
    cd.Filter = "Archivo de idioma (*.lan)|*.lan|Fuente de idioma (*.lanSource)|*.lansource"
    cd.ShowSave
    If cd.FileName = "" Then Exit Sub
    
    'sacar los filtros para que grabe todo!!!
    chkFaltantes.Value = 0
    txtFilter.Text = ""
    
    Dim LI As ListItem
    Dim aux As String
    For Each LI In lvw.ListItems
        'grabo      texto original       traduccion                           contexto (solo si es iansource deberia ser!!!)
        aux = aux + LI.Text + Chr$(31) + LI.ListSubItems(1).Text + Chr$(31) + LI.ListSubItems(2).Text + Chr$(30)
    Next
    EscribirArchivo cd.FileName, aux
    
    Abrir cd.FileName

End Sub

Private Sub mnuSalir_Click()
    'VER SI GRABO!
    'XXXXX
    
    If estaGrabado = False Then
        If MsgBox("No ha grabado los cambios, ¿Desea salir igual?", vbYesNo) = no Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then cmdSiguiente_Click
End Sub

Public Sub EscribirArchivo(path As String, cadena As String)
   ' On Error GoTo e
    
    Dim f, ts
    Dim S As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO.FileExists(path) Then
       FSO.CreateTextFile path
    End If
    
    Set f = FSO.GetFile(path)
    Set ts = f.OpenAsTextStream(2, 0)
     
    ts.Write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub

Private Sub LoadDictionary(path As String)
    'On Error GoTo e
   
    Dim S As String
    Dim aux() As String
    Dim par() As String
    
    S = LeerArchivo(path)
        
    aux = Split(S, Chr$(30)) 'separador oficla de cada cadena ?
    
    Dim I2 As Long 'index matriz interna completa de resumen
    ReDim Aux2(0)
    'NO ES LO MISMO UN ARCHIVO LAN QUE EL LANSOURCE QUE TIENE COMENTARIOS
    cCadenasT = 0
    
    If LCase(Right(path, 9)) = "lansource" Then
        'la estructura es:
        'Chr$(30) + dic.Keys(I) + Chr$(31) + dic.Items(I)
        '%97% es salto de carro
        '%98% es inicio de comentario
        '%99% fin de la cadena
    
    
        'para que no quiera traducir
        Frame1.Visible = False
        
        For I = 0 To UBound(aux)
            If aux(I) <> "" Then
                Dim RealText As String, Ayuda As String
                Dim TMP() As String, B As Long, C As Long
                par = Split(aux(I), Chr$(31))
                'primero dejo solo lo que se va a traducir diferenciandolo de la ayuda
                B = InStr(par(0), "%98%")
                If B = 0 Then
                    B = Len(par(0))
                    RealText = Left(par(0), B - 4)
                Else
                    RealText = Left(par(0), B - 1)
                End If
                'ahora le pongo saltos de carro donde van
                RealText = Replace(RealText, "%97%", vbCrLf)
                C = InStr(par(0), "%99%")
                If C = 0 Then C = Len(par(0)) 'si no esta es todo el texto
                If B > C Then 'cuando no existe el %98%
                    Ayuda = ""
                Else
                    Ayuda = Mid(par(0), B + 4, C - (B + 4))
                    Ayuda = Replace(Ayuda, "%97%", vbCrLf)
                End If
                
                I2 = I2 + 1
                ReDim Preserve Aux2(I2)
                Aux2(I2) = CStr(I2) + Chr(6) + RealText + Chr(6) + RealText + Chr(6) + Ayuda
            End If
        Next I
        
    End If
    
    If LCase(Right(path, 3)) = "lan" Then
        Frame1.Visible = True
        For I = 0 To UBound(aux)
            If aux(I) <> "" Then
            
                I2 = I2 + 1
                ReDim Preserve Aux2(I2)
            
                par = Split(aux(I), Chr$(31))
                
                
                Aux2(I2) = CStr(I2) + Chr(6) + par(0) + Chr(6) + par(1) + Chr(6)
                
                If UBound(par) > 1 Then
                    Aux2(I2) = Aux2(I2) + par(2)
                End If
                
                If par(0) <> par(1) Then
                    cCadenasT = cCadenasT + 1
                End If
                
                
            End If
        Next
    End If
    
    LoadAux2
    
    Form_Resize
    
    Exit Sub
e:
    MsgBox "No se encontro el archivo de idioma."
End Sub

Private Sub LoadAux2()

    'carga la lista final teniendo en cuenta si hay filtros
    Dim f As Boolean, cumpleF As Boolean
    f = CBool(chkFaltantes.Value)
        
    Dim S As String, cumpleS As Boolean
    S = txtFilter.Text
    
    lvw.ListItems.Clear
    lvw.ColumnHeaders.Add , , "Original"
    lvw.ColumnHeaders.Add , , "Traduccion"
    lvw.ColumnHeaders.Add , , "Contexto / ayuda"
    lvw.ColumnHeaders.Add , , "_ID"
    
    Dim itmX As ListItem
    
    Dim I As Long, CantShow As Long
    Dim PartesAux2() As String
    For I = 1 To UBound(Aux2)
        
        PartesAux2 = Split(Aux2(I), Chr(6))
        
        cumpleF = False 'solo lo mostrar si cumple todos los filtros
        cumpleS = False
        
        If f Then
            If PartesAux2(1) <> PartesAux2(2) Then
                cumpleF = False
            Else
                cumpleF = True
            End If
        Else
            cumpleF = True
        End If
        
        If Trim(S) <> "" Then
            
            If InStr(1, PartesAux2(1), S, vbTextCompare) Then
                cumpleS = True
            End If
            
            If InStr(PartesAux2(2), S) Then
                cumpleS = True
            End If
            
        Else
            cumpleS = True
        End If
        
        If (cumpleF And cumpleS) Then 'si se tiene que ver que conserve los colores indicadores originales
            'AGREGARLOS!
            CantShow = CantShow + 1
            Set itmX = lvw.ListItems.Add(CantShow, , PartesAux2(1))
                itmX.SubItems(1) = PartesAux2(2)
                itmX.SubItems(2) = PartesAux2(3)
                itmX.SubItems(3) = PartesAux2(0)
        
            If PartesAux2(1) = PartesAux2(2) Then
                itmX.ForeColor = vbRed
                itmX.ListSubItems(1).ForeColor = vbRed
                itmX.ListSubItems(2).ForeColor = vbRed
            Else
                itmX.ForeColor = vbBlue
                itmX.ListSubItems(1).ForeColor = vbBlue
                itmX.ListSubItems(2).ForeColor = vbBlue
            End If
        End If
    Next I
End Sub

Private Function LeerArchivo(path As String) As String
    Dim FSO 'As FileSystemObject
    Dim f
        
    Dim S As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If path = "" Then
        LeerArchivo = ""
    Else
        Set f = FSO.GetFile(path)
    End If
    Set ts = f.OpenAsTextStream(1)
    
    S = ts.ReadAll
    
    ts.Close
    LeerArchivo = S
End Function

Private Sub txtFilter_Change()
    LoadAux2
End Sub

'manejar la lista de ultimos archivos usados
Private Sub AddRedcent(sR As String)
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(sFileRecent, ForAppending, True)
        TE.WriteLine sR
    TE.Close
End Sub

Private Function ListRecents() As String()
    Dim TE As TextStream, Ret() As String
    
    ReDim Ret(0) 'para que nunca devuelva nulo
    
    Set TE = FSO.OpenTextFile(sFileRecent, ForReading, True)
        
        Dim j As Long
        Do While Not TE.AtEndOfStream
            Dim sFile As String
            'ver si el arch existe
            sFile = TE.ReadLine
            
            If FSO.FileExists(sFile) Then
            
                j = UBound(Ret) + 1
                ReDim Preserve Ret(j)
                Ret(j) = sFile
            End If
            
        Loop
    TE.Close
    
    ListRecents = Ret
End Function
