VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstbrCfg05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'hay categorias organizadas como un arbol
'cada config puede estar en una o varias
Private Type CategoriaCFG
    sNombre As String
    sId As Long 'para poder vincularlas
    sIdHijos(10) As Long 'cada una de las categorias que son padres de esta
End Type

'cada elemento de la configuracion tiene varias partes
Private Type CFs
    sNombre As String 'nombre del elemento
    sValor As String 'valor actual
    sDefaul As String 'valor predeterminado
    sHelp As String 'ayuda al usuario sobre la config
    sIdCategoria(10) As Long 'cada una de los id de las categoria a las que pertenece y en las que se debe mostrar
    
    'segun la info que guarda y el modo de cargarse (por ejemplo PATH que necesita cmdialog)
    ListaValues As String 'lista de valores aceptados separados por "|"
End Type

'en primer lugar se deben setear los valores predeterminados de cada funcion
'parece molesto pero obliga a organizar la configuracion, cada elemento debe ser dado de alta
'puede no ser dado de alta pero su valor si no existe sera "" siempre
'luego simplemente se usa GetV para ver que valor tiene (y si no esta el predefinido)
'setV setea un elemento de la configuracion sin grabar todo el archivo de nuevo
'hay que grabarlo al finalizar los cambios

'los valores prederminados no se graban en la config, solo el nombre y el valor actual

'si quieren ser desprolijos y no dar de alta los valores predeterminados pueden usar
'GetV2 que es igual que GetV pero tiene un segundo parametro para el valor predeterminado

'agregado feb09. la configuracion de los valores predeterminados es un archivo de texto
'si esto no fuera asi deberia duplicar el modo de inicializacion en 3pm y en config3pm

'ademas cada config puede estar en una o varias categorias

Private MainArchivo As String
Private CategS() As CategoriaCFG 'no definido aun SEGUIRAQUI
Private Conf() As CFs

'para grabar inteligentemente
Private HayCambiosPendientes As Boolean
Dim fso As New Scripting.FileSystemObject

Public Function SetMainFile(nf As String)
    tErr.Anotar "c5_SMF", nf
    MainArchivo = nf
End Function

Public Function GetV(Conf2 As String) As String
    'leer el archivo de configuracion y devolver valor
    tErr.Anotar "c5_aaa", Conf2
    
    Dim tmp As String
    
    tmp = "NO EXISTE"
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        If UCase(Conf(H).sNombre) = UCase(Conf2) Then
            'por si hay algun "=" en la respuesta
            RST = Conf(H).sValor
            
            'y si esta vacio!!!!
            If RST <> "" Then
                tmp = RST
            Else
                tmp = Conf(H).sDefaul
            End If
            
            Exit For
        End If
    
    Next H
    
    tErr.Anotar "c5_aab", tmp
    If tmp = "NO EXISTE" Then
        GetV = ""
    Else
        GetV = tmp
    End If
    
End Function

'define el valor predeterminado de cada config
'me obligara a definir cada elemento de la config
'yme libera de escribir el valro predeterminado cada vez que llamo a la funcion
Public Function SetDef(Conf2 As String, DefValue As String, Optional Ayuda As String)
    Dim isSet As Boolean
    isSet = False
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        If UCase(Conf(H).sNombre) = UCase(Conf2) Then
            Conf(H).sDefaul = DefValue
            Conf(H).sHelp = Ayuda
            isSet = True
            Exit For
        End If
    Next H
    
    'si no existia crear la nueva config!
    If isSet = False Then
        AddDefinition Conf2, DefValue, Ayuda
    End If
End Function

Public Function SetV(Conf2 As String, newValue As String) As String
    tErr.Anotar "c5_aac", Conf2, newValue
    
    Dim isSet As Boolean
    isSet = False
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        If UCase(Conf(H).sNombre) = UCase(Conf2) Then
            If Conf(H).sValor <> newValue Then
                HayCambiosPendientes = True
                Conf(H).sValor = newValue
            End If
            isSet = True
            Exit For
        End If
    Next H
    
    tErr.Anotar "c5_aad", isSet
    'si no existia crear la nueva config!
    If isSet = False Then
        AddCF Conf2, newValue
        HayCambiosPendientes = True
    End If
    
End Function

Public Sub AbrirCfg(Optional nf As String = "")
    
    On Local Error GoTo errLoadConf
    
    tErr.Anotar "c5_aae", nf
    HayCambiosPendientes = False
    
    If nf <> "" Then SetMainFile nf

    ReDim Conf(0)
    
    Dim TXT As String, CFG As String, RST As String
    tErr.Anotar "c5_aaf", MainArchivo
    If fso.FileExists(MainArchivo) Then
        tErr.Anotar "c5_aag"
        Set TE = fso.OpenTextFile(MainArchivo, ForReading, False)
            Dim FullConfig As String
            FullConfig = TE.ReadAll
        TE.Close
        
        tErr.Anotar "c5_aah"
        'desencriptar
        FullConfig = Encriptar(FullConfig, True)
        
        tErr.Anotar "c5_aai"
        Dim sp2() As String
        sp2 = Split(FullConfig, Chr(5)) 'cada renglon de configuracion
        
        Dim SP() As String, CantCfs As Long
        Dim H As Long
        tErr.Anotar "c5_aaj", UBound(sp2)
        For H = 0 To UBound(sp2)
            tErr.Anotar "c5_aak", H, sp2(H)
            TXT = sp2(H)
            SP = Split(TXT, "=")
            If TXT <> "" Then
                CFG = SP(0)
                RST = Mid(Trim(TXT), Len(CFG) + 2, Len(TXT) - Len(CFG) + 1)
            
                'lo agrego a la lista de disponibles
                tErr.Anotar "c5_aal", H, CFG, RST
                AddCF CFG, RST
            End If
        Next H
    End If
    
    tErr.Anotar "c5_aam"
        
    Exit Sub
    
errLoadConf:
    tErr.AppendLog "errLoadConf_897", tErr.ErrToTXT(Err)
End Sub

Public Function GrabarConfig(Optional nf As String) As Long
        
    tErr.Anotar "c5_aan", nf
    
    On Local Error GoTo errChgConf
    
    If nf <> "" Then SetMainFile nf
    
    If HayCambiosPendientes = False Then
        GrabarConfig = 0
        Exit Function
    End If
    
    'escribir un temporal desencriptado
    Dim FullConfig As String
    FullConfig = ""
    
    Dim H As Long
    tErr.Anotar "c5_aao", UBound(Conf)
    
    For H = 1 To UBound(Conf)
        tErr.Anotar "c5_aap", Conf(H).sNombre, Conf(H).sValor
        FullConfig = FullConfig + Conf(H).sNombre + "=" + Conf(H).sValor
        If H < UBound(Conf) Then FullConfig = FullConfig + Chr(5) 'SEPARADOR DE CONFIGS!!
    Next H
    
    tErr.Anotar "c5_aaq"
    'encriptar
    FullConfig = Encriptar(FullConfig, False)
    'grabar el kilombo
    tErr.Anotar "c5_aar"
    Set TE = fso.CreateTextFile(MainArchivo, True)
        TE.Write FullConfig
    TE.Close
    
    tErr.Anotar "c5_aas"
    'hacer una copia de seguridad cada vez que haya cambios
    If BackupMainArchivo <> "" Then
        fso.CopyFile MainArchivo, MainArchivo + ".BUP"
    End If
    
    tErr.Anotar "c5_aat"
    Exit Function
    
errChgConf:
    GrabarConfig = Err.Number
    tErr.AppendLog "errChgConf_897", tErr.ErrToTXT(Err)
End Function

Private Sub AddCF(sName As String, sVal As String)
    Dim c As Long
    c = UBound(Conf) + 1
    ReDim Preserve Conf(c)
    Conf(c).sNombre = sName
    Conf(c).sValor = sVal
    Conf(c).sDefaul = ""
End Sub

Private Sub AddDefinition(sName As String, sVal As String, Optional Ayuda As String = "")
    Dim c As Long
    c = UBound(Conf) + 1
    ReDim Preserve Conf(c)
    Conf(c).sNombre = sName
    'Conf(c).sValor = "" no lo toco por las dudas quwe ya este seteado
    Conf(c).sDefaul = sVal
    Conf(c).sHelp = Ayuda
End Sub

Public Function GetV2(Conf2 As String, ValDef) As String
    'leer el archivo de configuracion y devolver valor (teniendo en cuenta un valor predeterminado)
    Dim tmp As String
    
    tmp = "NO EXISTE"
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        If UCase(Conf(H).sNombre) = UCase(Conf2) Then
            'por si hay algun "=" en la respuesta
            RST = Conf(H).sValor
            
            'y si esta vacio!!!!
            If RST <> "" Then
                tmp = RST
            Else
                tmp = ValDef
            End If
            
            Exit For
        End If
    
    Next H
    
    
    If tmp = "NO EXISTE" Then
        GetV2 = ValDef
    Else
        GetV2 = tmp
    End If
    
End Function

Public Function GetHelp(Conf2 As String) As String
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        If UCase(Conf(H).sNombre) = UCase(Conf2) Then
            GetHelp = Conf(H).sHelp
            Exit Function
        End If
    Next H
    
End Function

'otras formas de leer valores como no-string
Public Function GetSingle(cf As String) As Single
    Dim SN As Single, t As String
    t = GetV(cf)
    If t = "" Then t = "0"
    SN = CSng(t)
    GetSingle = SN
End Function

Public Function GetInt(cf As String) As Integer
    Dim SN As Integer, t As String
    t = GetV(cf)
    If t = "" Then t = "0"
    SN = CInt(t)
    GetInt = SN
End Function

Public Function GetLong(cf As String) As Long
    Dim SN As Long, t As String
    t = GetV(cf)
    If t = "" Then t = "0"
    SN = CLng(t)
    GetLong = SN
End Function

Public Function GetDouble(cf As String) As Double
    Dim SN As Double, t As String
    t = GetV(cf)
    If t = "" Then t = "0"
    SN = CDbl(t)
    GetDouble = SN
End Function

Public Function GetBool(cf As String) As Boolean
    Dim SN As Boolean, t As String
    t = GetV(cf)
    If t = "" Then t = "0"
    SN = CBool(t)
    GetBool = SN
End Function

'inicializar los valores predeterminados y la configuracion
Public Function IniCFG(sf As String)
    'el archivo donde se leera es el parametro

End Function

'se llama igual y devuelve lo mismo (a su manera) para el OCX de la config pueda mostrar los valores actuales
Public Function GetPropertisList(Optional IncludeValues As Boolean = False) As String()
    
    Dim TmpO() As String 'lista final a devolver
    ReDim TmpO(0) 'por si no hay que devuelva una matriz valida
    
    'nombre y resultado de cada valor
    Dim nCFG As String
    Dim RST As String
    
    Dim H As Long
    For H = 1 To UBound(Conf)
        nCFG = Conf(H).sNombre
        RST = Conf(H).sValor
        'y si esta vacio!!!! y si se configuro intencionalmente vacio !!!!! SEGUIRAQUI
        If RST <> "" Then
            tmp = RST
        Else
            tmp = Conf(H).sDefaul
        End If
        
        ReDim Preserve TmpO(H)
        TmpO(H) = nCFG
        If IncludeValues Then TmpO(H) = TmpO(H) + "=" + RST
    
    Next H
    
    GetPropertisList = TmpO
    
End Function

Public Function GetConfigCantidad() As Long
    GetConfigCantidad = UBound(Conf)
End Function

Public Function GetNameCfgByIndex(ix As Long) As String
    If ix = 0 Or ix > UBound(Conf) Then
        'seguiraqui marcar error
        GetNameCfgByIndex = ""
    Else
        GetNameCfgByIndex = Conf(ix).sNombre
    End If
End Function

Public Function GetValueCfgByIndex(ix As Long) As String
    If ix = 0 Or ix > UBound(Conf) Then
        'seguiraqui marcar error
        GetValueCfgByIndex = ""
    Else
        GetValueCfgByIndex = Conf(ix).sValor
    End If
End Function

Public Function SetPathError(t As String)
    tErr.FileLog = t
End Function

Private Sub Class_Initialize()
    ReDim Conf(0)
    Set tErr = New tbrErrores.clsTbrERR
End Sub
