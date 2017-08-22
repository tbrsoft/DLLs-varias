VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'lista de los elementos jerarquica de las opciones
'lo uso para que al insertar el pendrive de 3pm este permita
'una lista de acciones a hacer y las configuraciones mas importantes

'todos los elemntos de esta lista
Private mElementos() As New clsElem
Private FSo As New Scripting.FileSystemObject
'Private tmpElem As clsElem
Private TV 'objeto treeview interno que maneja todo
Private mPermHeredado As clsPERM  'al hacerle un load se le dicen los permisos

Public Function Load(F As String, thPerms As clsPERM) As Long
    'el archivo tiene que ser simple ya que pienso escribirlo a mano
    'ID | caption | help | tipo | padres separados por espacio (los hijos de cero estan en la raiz | TAGS
    'EJ: 1 | Agregar musica | Inserte la musica que trae en el pendrive | 5 | 0
    'los tags pueden ser por ejemplo codigo de la funcion o proceso que larga un tipo 4 (proceso)
    
    'el segundo parametro son los permisos de mostrar o no
    Set mPermHeredado = thPerms
    On Local Error GoTo errLOAD
    
    terr.Anotar "raa", F, thPerms.toString
    
    If FSo.FileExists(F) = False Then
        Load = 1
        terr.Anotar "qag", F
        Exit Function
    End If
    
    Dim ts As TextStream
    Dim tx As String, sp() As String
    Dim Baneados() As Long 'lista de nodos que no deben mostrarse por que estan bloqueados
    ReDim Baneados(0) 'NO USAR EL CERO
    Set ts = FSo.OpenTextFile(F, ForReading)
OTROMAS:
        Do While Not ts.AtEndOfStream
            tx = ts.ReadLine
            sp = Split(tx, "|")
            'terr.Anotar "qah", tx se hace muy largo
            
            Dim renglonOK As Boolean
            renglonOK = False
            If UBound(sp) >= 5 Then renglonOK = True
            'comentarios con # o //
            If UBound(sp) >= 0 Then
                If Left(Trim(sp(0)), 1) = "#" Then renglonOK = False
                If Left(Trim(sp(0)), 2) = "//" Then renglonOK = False
            Else
                renglonOK = False
            End If
            
            If renglonOK Then 'SOLO SI ES UN RENGLON DE DATOS, LOS DEMAS SON COMENTARIOS
                
                'terr.Anotar "qah2", tx 'aqui hace que el log sea solo de donde entro
                
                'VER SI ES EL PADRE DE TODOS!!!
                If CLng(Trim(sp(0))) = 0 Then
                    'crear el nodo padre de todo
                    mElementos(0).id = 0
                    mElementos(0).Marca = "SUPER-ROOT"
                    mElementos(0).Caption = VS2.TraslateVars(Trim((sp(2))))
                    mElementos(0).Help = VS2.TraslateVars(Replace(Trim(sp(3)), "&&", vbCrLf))
                    mElementos(0).eType = SoloInfo
                    
                    terr.Anotar "qai", Trim(sp(2))
                Else
                
                    'ver que haya permisos para mostrar esto!
                    Dim CargarEste As Boolean
                    Dim thMARCA As String
                    thMARCA = Trim(sp(1))
                    Dim thID As Long 'id de este nodo
                    thID = CLng(Trim(sp(0)))
                    Dim thIdPadre As Long
                    thIdPadre = CLng(Trim(sp(5)))
                
                    terr.Anotar "rab", thMARCA, thID, thIdPadre, thPerms.toString
                    
                    CargarEste = (mPermHeredado.GetPermByIndexNodo(thID) <> 0) 'predeterminado SI SE VE
                    
                    'SEGUIRAQUI
                    'si un elemento no se carga sus hijos no deberian cargarse!!
                    Dim P11 As Long
                    If CargarEste = False Then
                        'agregarlo a una lista de baneados!
                        P11 = UBound(Baneados) + 1
                        ReDim Preserve Baneados(P11)
                        Baneados(P11) = thID
                        terr.Anotar "rac", thID
                        GoTo OTROMAS
                    End If
                    
                    Dim P12 As Long
                    'ver que este o le padre de este no esten baneados!
                    For P12 = 1 To UBound(Baneados)
                        If Baneados(P12) = thID Then
                            terr.Anotar "rad", thID
                            'esta baneado!
                            GoTo OTROMAS
                        End If
                        
                        If Baneados(P12) = thIdPadre Then
                            'esta baneado indirectamente (por su padre), marcarlo para que sus hijos no entren!
                            P11 = UBound(Baneados) + 1
                            ReDim Preserve Baneados(P11)
                            Baneados(P11) = thID
                            terr.Anotar "rae", thID, thIdPadre
                            GoTo OTROMAS
                        End If
                        
                    Next P12
                    
                    'si llego hasta aca tiene permiso!
                    Dim tmpElem As New clsElem
                    Set tmpElem = addElement
                    tmpElem.id = CLng(Trim(sp(0)))
                    tmpElem.Marca = Trim(sp(1))
                    
                    terr.Anotar "qaj", CLng(Trim(sp(0))), Trim(sp(1)), Trim(sp(2)), Trim(sp(3))
                    terr.Anotar "qaj22", sp(4), sp(5)
                    
                    tmpElem.Caption = VS2.TraslateVars(Trim((sp(2))))
                    tmpElem.Help = VS2.TraslateVars(Replace(Trim(sp(3)), "&&", vbCrLf))
                    tmpElem.eType = CLng(Trim(sp(4)))
                    tmpElem.Padre = CLng(Trim(sp(5)))
                    
                    'cargar subcontenido segun tipo de elemento
                    If tmpElem.eType = EjecutarProceso Then
                        terr.Anotar "qak", Trim(sp(6))
                        tmpElem.Internal_VerEXE.orden = VS2.TraslateVars(Trim(sp(6))) 'no la he usado pero puedo necesitar que traduzca
                    End If
                    
                    If tmpElem.eType = ListaCombo Then
                        terr.Anotar "qal", Trim(sp(6))
                        'puede ser que sean lista donde lo mismo que se lee es lo que necesito
                        'o puede ser una lista donde necesito que se vea algo pero se grabe otra cosa
                        'por ejemplo en Activar H2K quiero mostrar NO y SI pero en realidad hay un boolean por atras
                        'cuando necesite esto voy a usar la "\" para separa dentro de cada elemento
                        'de la lista lo que se muestra y lo que vale realmente
                        
                        'EN EL CASO DE LISTA KEY SP(6)="%%listakey%% que esta bien
                        'pero no se sabe
                        sp(6) = VS2.TraslateVars(Trim(sp(6)))
                        
                        If InStr(sp(6), "·") Then 'es el punto medio el separador
                            tmpElem.Internal_ListaSImple.LoadFromString sp(6), "&&", "·"
                        Else
                            tmpElem.Internal_ListaSImple.LoadFromString VS2.TraslateVars(Trim(sp(6))), "&&"
                        End If
                        
                    End If
                    
                    If tmpElem.eType = TextoSimple Then
                        terr.Anotar "qal22", Trim(sp(6))
                        tmpElem.Internal_TextoSimple.Predeterminado = VS2.TraslateVars(Trim(sp(6)))
                    End If
                    
                    If tmpElem.eType = SelectPath Then
                        terr.Anotar "qal29", Trim(sp(6))
                        tmpElem.Internal_TextoSimple.Predeterminado = VS2.TraslateVars(Trim(sp(6)))
                        'sp(7) puede ser "folder" o "files jpg gif"
                        Dim seven As String, sp7() As String
                        seven = Trim(sp(7))
                        sp7 = Split(seven)
                        If LCase(sp7(0)) = "folder" Then tmpElem.Internal_TextoSimple.ShowFolder = 1
                        If LCase(sp7(0)) = "files" Then
                            tmpElem.Internal_TextoSimple.ShowFile = 1
                            Dim y As Long, fils As String, fils2 As String
                            fils = "(": fils2 = ""
                            For y = 1 To UBound(sp7)
                                fils = fils + "*." + sp7(y) + " "
                                fils2 = fils2 + "*." + sp7(y)
                                If y < UBound(sp7) Then fils2 = fils2 + ";" ' si se separa por "|" hay un renglon para cada tipo y generalmente no es lo que hace falta
                            Next y
                            fils = fils + ")|"
                            fils = "Archivos " + fils + fils2
                            '"Imagnes (*.jpg *.gif)|*.jpg; *.gif"
                            '
                            tmpElem.Internal_TextoSimple.Filter = fils
                            
                        End If
                    End If
                    
                    If tmpElem.eType = Numero Then
                        terr.Anotar "qam"
                        Dim sp2() As String
                        sp2 = Split(Trim(sp(6)))
                        tmpElem.Internal_Numeros.MinVal = CLng(sp2(0))
                        tmpElem.Internal_Numeros.MaxVal = CLng(sp2(1))
                        tmpElem.Internal_Numeros.Step = CLng(sp2(2))
                        'traducirlo, por ejemplo se usa AP como variable
                        tmpElem.Internal_Numeros.Predeterminado = CLng(sp2(3))
                    End If
                End If
            End If
        Loop
    ts.Close
    
    terr.AppendLog "NodeCharged.-"
    
    Exit Function
errLOAD:
    terr.AppendLog "LoadCFG091823", terr.ErrToTXT(Err)
    Resume Next
End Function

Public Function addElement() As clsElem
    Dim mID As Long
    'el ID se debe generar solo segun disponibilidad
    mID = UBound(mElementos) + 1
    ReDim Preserve mElementos(mID)
    'mElementos(mID).ID = mID 'el archivo de cofig ya trae los ids
    Set addElement = mElementos(mID)
End Function

Private Sub Class_Initialize()
    ReDim mElementos(0)
End Sub

'devuelve el indice en la matriz de un elemento dado
Private Function Get_IdMatriz_FromIdElem(isSearch As Long) As Long
    Dim a As Long
    For a = 0 To UBound(mElementos)
        If mElementos(a).id = idsearch Then
            Get_IdMatriz_FromIdElem = a
            Exit Function
        End If
    Next a
    
    Get_IdMatriz_FromIdElem = -1
End Function

Public Sub LoadOnTreeView(TreeV As Object)
    
    terr.Anotar "qan"
    Set TV = TreeV
    
    'manejabilidad
    TV.LabelEdit = tvwManual 'que no se puedan editar los textos
    TV.TabStop = False 'para manejar yo lo de las teclas y no duplicar pulsaciones
    TV.HideSelection = False
    TV.SingleSel = False 'se abre cuando se elige
    
    'apariencia
    TV.Style = 0
    TV.Appearance = cc3D
    TV.FullRowSelect = False
    TV.Indentation = 140
    TV.Font.Name = "Courier New"
    TV.Font.Size = "10"
    
    'vaciar
    TV.Nodes.clear
    
    terr.Anotar "qao"
    'TV.Nodes.Add relative, relation, Key, Text
    'la raiz siempre es el cero
    Dim NOD As Node
    Set NOD = TV.Nodes.Add(, tvwFirst, "NODO 0", mElementos(0).Caption)
    NOD.Expanded = True
    
    Dim a As Long, NodPadre As String, NodEste As String
    For a = 1 To UBound(mElementos)
        terr.Anotar "qap", a, mElementos(a).Caption, mElementos(a).id, mElementos(a).Padre
        NodPadre = "NODO " + CStr(mElementos(a).Padre)
        NodEste = "NODO " + CStr(mElementos(a).id)
        Set NOD = TV.Nodes.Add(NodPadre, tvwChild, NodEste, mElementos(a).Caption)
        NOD.ForeColor = RGB(200, 200, 250)
        NOD.BackColor = RGB(90, 90, 120)
        'NOD
        Set mElementos(a).NodeOp = NOD 'muy practico para interactuar con el tv a traves de mi clase
    Next a
    
    TV.Nodes("NODO 0").Selected = True
End Sub

Public Function GetElement(idElem As Long) As clsElem
    terr.Anotar "rca", idElem
    Dim a As Long
    For a = 0 To UBound(mElementos)
        terr.Anotar "rce", a
        If mElementos(a).id = idElem Then
            Set GetElement = mElementos(a)
            terr.Anotar "rcb", mElementos(a).Caption, mElementos(a).id
            Exit Function
        End If
    Next a
    terr.Anotar "rcc"
    Set GetElement = Nothing
End Function

Public Function GetElementByMarca(laMarca As String) As clsElem
    terr.Anotar "rcd"
    Dim a As Long
    For a = 0 To UBound(mElementos)
        terr.Anotar "rcf", a
        If mElementos(a).Marca = laMarca Then
            Set GetElementByMarca = mElementos(a)
            terr.Anotar "rcg", mElementos(a).Caption, mElementos(a).Marca
            Exit Function
        End If
    Next a
    Set GetElementByMarca = Nothing
End Function

Public Function Status() As String
    terr.Anotar "rch"
    Dim a As Long, t As String
    For a = 0 To UBound(mElementos)
        t = t + mElementos(a).Caption
        If mElementos(a).eType = EjecutarProceso Then t = t + " / Proc:" + mElementos(a).Internal_VerEXE.orden
        If mElementos(a).eType = ListaCombo Then t = t + " / ListaSel:" + mElementos(a).Internal_ListaSImple.GetSelectOp
        If mElementos(a).eType = TextoSimple Then t = t + " / TextoSimple:" + mElementos(a).Internal_TextoSimple.Valor
        If mElementos(a).eType = SelectPath Then t = t + " / TextoSimple:" + mElementos(a).Internal_TextoSimple.Valor
        If mElementos(a).eType = SoloInfo Then t = t + " / nFo:" + mElementos(a).Help
        
        t = t + vbCrLf
    Next a
    terr.Anotar "rci", t
    Status = t
End Function

Public Sub TV_Next()
    
    terr.Anotar "rcj"
    If TV.SelectedItem Is Nothing Then Exit Sub
    
    'la primera opcion es un hijo (si esta expandido)
    If TV.SelectedItem.Children > 0 And TV.SelectedItem.Expanded Then
        TV.SelectedItem.Child.Selected = True
        Exit Sub
    End If
    
    terr.Anotar "rck"
    'ver si sigue un hermano
    If Not (TV.SelectedItem.Next Is Nothing) Then
        TV.SelectedItem.Next.Selected = True
        Exit Sub
    End If
    
    'recursivamente buscar tios-next
    Dim ND As Node
    Set ND = TV.SelectedItem
    Do
        If Not (ND.Parent Is Nothing) Then
            Set ND = ND.Parent
            If Not (ND.Next Is Nothing) Then
                ND.Next.Selected = True
                Exit Sub
            End If
            terr.Anotar "rcl"
        Else
            terr.Anotar "rcm"
            Exit Sub
        End If
            
    Loop
    
    terr.Anotar "rcn"
End Sub

Public Sub TV_Prev()
    
    terr.Anotar "rco"
    If TV.SelectedItem Is Nothing Then Exit Sub
    
    'primera opcion es ultimo hijo del previous
    If Not (TV.SelectedItem.Previous Is Nothing) Then
        If Not (TV.SelectedItem.Previous.Child Is Nothing) And TV.SelectedItem.Previous.Expanded Then
            If Not (TV.SelectedItem.Previous.Child.LastSibling Is Nothing) Then
                terr.Anotar "rcp"
                'hasta aca parece divertido pero es necesario que se hijo los ultimos hijos de todos si estan abiertos
                Dim ND As Node
                Set ND = TV.SelectedItem.Previous.Child.LastSibling
                Do
                    If ND.Children > 0 And ND.Expanded Then
                        Set ND = ND.Child.LastSibling
                    Else
                        Exit Do
                    End If
                Loop
                ND.Selected = True
                Exit Sub
            End If
        End If
        terr.Anotar "rcq"
    End If
    
    terr.Anotar "rcr"
    'opcion es hermano
    If Not (TV.SelectedItem.Previous Is Nothing) Then
        TV.SelectedItem.Previous.Selected = True
        Exit Sub
    End If
    
    terr.Anotar "rcs"
    'el padre puede seguir
    If Not (TV.SelectedItem.Parent Is Nothing) Then
        TV.SelectedItem.Parent.Selected = True
        Exit Sub
    End If
    
    terr.Anotar "rct"
    
End Sub

Public Sub TV_KillMarcados(laMarca As String)
    
    terr.Anotar "rcu", laMarca
    
    Dim a As Long, e As Long
    For a = 0 To UBound(mElementos)
        If mElementos(a).Marca = laMarca Then
            'eliminarlo!
            terr.Anotar "rcv", mElementos(a).Marca
            TV_KillNodeByIndex mElementos(a).NodeOp.index
        End If
    Next a
    
    terr.Anotar "rcw"
End Sub

Public Sub TV_KillNodeSel()
    terr.Anotar "rcx"
    If TV.SelectedItem Is Nothing Then Exit Sub
    TV.Nodes.Remove TV.SelectedItem.index
    TV_Next
    terr.Anotar "rcy"
End Sub

Public Sub TV_KillNodeByIndex(i As Long)
    terr.Anotar "rcz", i
    If TV.Nodes(i) Is Nothing Then Exit Sub
    TV.Nodes.Remove i
End Sub


'devuelve una matriz de elementos con detrminada marca
Public Function getMarcados(laMarca As String)
    terr.Anotar "rda", laMarca
    Dim Els() As clsElem
    ReDim Els(0)
    Dim a As Long, e As Long
    For a = 0 To UBound(mElementos)
        terr.Anotar "rdb", a
        If mElementos(a).Marca = laMarca Then
            e = UBound(Els) + 1
            ReDim Preserve Els(e)
            Set Els(e) = mElementos(a)
            terr.Anotar "rdc"
        End If
    Next a
    terr.Anotar "rdd"
    getMarcados = Els
End Function

Public Function ExisteNodoByKey(K As String) As Boolean
    On Local Error GoTo NoExiste
    terr.Anotar "rde"
    ExisteNodoByKey = Not (IsNull(TV.Nodes(K))) 'solo funciona para el true!
    
    Exit Function
NoExiste:
    If Err.Number = 35601 Then 'nop existe mas el nodo
        ExisteNodoByKey = False
    End If
End Function

'para grabar en alguna configuracion (opcion completa o de a uno)
Public Function GetFinalValues() As String()
    terr.Anotar "rdf"
    Dim Tmp() As String
    
    Dim a As Long, e As String
    For a = 0 To UBound(mElementos)
        ReDim Tmp(a)
        '"marca" es un ID (no verificado por duplicados, pero supuestamente no duplicado salvo excepciones donde se necesitan varios elementos (origenes de discos) con algo en comun)
        e = mElementos(a).Marca + "=" + mElementos(a).ValorElegido
        Tmp(a) = e
        terr.Anotar "rdg", a, e
    Next a

    GetFinalValues = Tmp
    terr.Anotar "rdh"
End Function

'la configuracion al grabarse puede buscar valores nuevos aqui
Public Function GetFinalValue(laMarca As String) As String
    terr.Anotar "rdi", laMarca
    Dim a As Long
    For a = 0 To UBound(mElementos)
        If LCase(mElementos(a).Marca) = LCase(laMarca) Then
            terr.Anotar "rdj", mElementos(a).ValorElegido
            GetFinalValue = mElementos(a).ValorElegido
            Exit Function
        End If
    Next a

    terr.Anotar "rdk"
    GetFinalValue = "NULL"

End Function

'puede ser necesario que lea valores grabados en una archivo (cuando se usa como configuración)
Public Function SetFinalValue(LVals() As String) As String
    terr.Anotar "rdl"
    'el parametro es la lista de valores separados por "="
    Dim a As Long, a2 As Long
    Dim lValProp As String, lValValue As String
    
    For a = 1 To UBound(LVals)
        'descartar valores chongos
        If LVals(a) = "" Then GoTo sig4
        
        
        Dim sp() As String
        sp = Split(LVals(a), "=")
        lValProp = sp(0)
        lValValue = sp(1)
        terr.Anotar "rdm", a, lValProp, lValValue
        'ver si este valor esta en este objeto
        For a2 = 0 To UBound(mElementos)
            terr.Anotar "rdo3", a2
            If LCase(mElementos(a2).Marca) = LCase(lValProp) Then
            
                If LCase(lValProp) = "teclader" Then
                    lValProp = lValProp
                End If
                
                terr.Anotar "rdn", mElementos(a2).eType
                'VER DE QUE TIPO ES PARA ASIGNARLE LO QUE CORRESPONDE
                If mElementos(a2).eType = ListaCombo Then
                    mElementos(a2).Internal_ListaSImple.TryToSelectFromInternalOptions lValValue
                    
                End If
                If mElementos(a2).eType = TextoSimple Then mElementos(a2).Internal_TextoSimple.Valor = lValValue
                If mElementos(a2).eType = SelectPath Then mElementos(a2).Internal_TextoSimple.Valor = lValValue
                
                If mElementos(a2).eType = Numero Then
                    mElementos(a2).Internal_Numeros.Valor = CLng(lValValue)
                End If
                
                'SEGUIRAQUI ver con multi y los nuevos tipos que se definan
                Exit For
            End If
        Next a2
sig4:
        terr.Anotar "rdo2", a
    Next a
    terr.Anotar "rdo"
End Function

