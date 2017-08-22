VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cl_tbrSoftSelectKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'permite almacenar cadenas de texto que representan un numero y _
    hacerlo configurable con un formulario incluido

Private Type tbrKeys
    Name As String 'nombre interno de la config
    lKeyCode As Long 'sale de KD e identifica realmente
    Description As String 'detalle que aparece en la autoconfig para que el tipo elija
End Type

Private KD(144) As String 'no todos se pueden usar
Private Ks() As tbrKeys
Private FSO As New Scripting.FileSystemObject
Private mArchivo As String

Public Function ADDcfg(sName As String, lKeyCode As Long, Descripcion As String) As Long
    Dim H As Long
    H = UBound(Ks) + 1
    ReDim Preserve Ks(H)
    
    Ks(H).Name = sName
    Ks(H).lKeyCode = lKeyCode
    Ks(H).Description = Descripcion
    
End Function

Public Function SaveCfg(sFile As String) As Long
    
    If sFile <> "" Then mArchivo = sFile
    
    'creo todo el texto primero
    Dim TMP As String
    Dim H As Long
    TMP = ""
    For H = 1 To UBound(Ks)
        TMP = TMP + Ks(H).Name + Chr(5) + CStr(Ks(H).lKeyCode) + Chr(5) + Ks(H).Description
        If H < UBound(Ks) Then TMP = TMP + Chr(6)
    Next H
        
    If FSO.FileExists(mArchivo) Then FSO.DeleteFile mArchivo, True
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(mArchivo, True)
        TE.Write TMP
    TE.Close
End Function

Public Function LoadCfg(sFile As String) As Long

    If sFile <> "" Then mArchivo = sFile

    If FSO.FileExists(mArchivo) = False Then
        LoadCfg = 1
        Exit Function
    End If
    
    
    Dim TE As TextStream, SP() As String, SP2() As String
    Dim TMP As String
    Set TE = FSO.OpenTextFile(mArchivo)
        TMP = TE.ReadAll
    TE.Close
    
    SP = Split(TMP, Chr(6)) 'cada renglon
    
    ReDim Ks(0)
    
    Dim H As Long, J As Long
    For H = 0 To UBound(SP)
        SP2 = Split(SP(H), Chr(5))
        For J = 0 To UBound(SP2)
            ReDim Preserve Ks(H + 1)
            Ks(H + 1).Name = SP2(0)
            Ks(H + 1).lKeyCode = CLng(SP2(1))
            Ks(H + 1).Description = SP2(2)
        Next J
    Next H
    
    LoadCfg = 0

End Function

Public Function ChgCfg(sName As String, NewKeyCode As Long) As Long
    
    Dim Encontro As Boolean
    Encontro = False
    Dim H As Long
    For H = 1 To UBound(Ks)
        If LCase(Ks(H).Name) = LCase(sName) Then
            Ks(H).lKeyCode = NewKeyCode
            Encontro = True
            Exit For
        End If
    Next H
    
    If Encontro = False Then
        ChgCfg = 1
    Else
        ChgCfg = 0
    End If
    
End Function

Public Function ReadCfgByName(sName As String) As Long
    
    Dim H As Long
    For H = 1 To UBound(Ks)
        If LCase(Ks(H).Name) = LCase(sName) Then
            ReadCfgByName = Ks(H).lKeyCode
            Exit Function
        End If
    Next H
    
    'no encontro nada
    ReadCfgByName = -1
    
End Function

Public Function GetName(sId As Long) As String
    If sId > UBound(Ks) Then
        GetName = ""
    Else
        GetName = Ks(sId).Name
    End If
End Function

Public Function GetDescr(sId As Long) As String
    If sId > UBound(Ks) Then
        GetDescr = ""
    Else
        GetDescr = Ks(sId).Description
    End If
End Function

Public Function GetKeyCode(sId As Long) As Long
    If sId > UBound(Ks) Then
        GetKeyCode = 0
    Else
        GetKeyCode = Ks(sId).lKeyCode
    End If
End Function

Public Function GetMaxCfg() As Long
    GetMaxCfg = UBound(Ks)
End Function

Public Function GetKDDesc(i As Long) As String
    GetKDDesc = KD(i)
End Function

Private Sub IniKD()
    KD(1) = "1 Botón primario del mouse"
    KD(2) = "2 Botón secundario del mouse"
    KD(3) = "3 Tecla Cancel"
    KD(4) = "4 Botón central del mouse"
    KD(8) = "8 Tecla RETROCESO"
    KD(9) = "9 Tecla TAB"
    KD(12) = "12 Tecla SUPR"
    KD(13) = "13 Tecla ENTRAR"
    KD(16) = "16 Tecla MAYÚS"
    KD(17) = "17 Tecla CTRL"
    KD(18) = "18 Tecla MENÚ"
    KD(19) = "19 Tecla PAUSA"
    KD(20) = "20 Tecla BLOQ MAYÚS"
    KD(27) = "27 Tecla ESC"
    KD(32) = "32 Tecla BARRA ESPACIADORA"
    KD(33) = "33 Tecla RE PÁG"
    KD(34) = "34 Tecla AV PÁG"
    KD(35) = "35 Tecla FIN"
    KD(36) = "36 Tecla INICIO"
    KD(37) = "37 Tecla FLECHA IZQUIERDA"
    KD(38) = "38 Tecla FLECHA ARRIBA"
    KD(39) = "39 Tecla FLECHA DERECHA"
    KD(40) = "40 Tecla FLECHA ABAJO"
    KD(41) = "41 Tecla SELECT"
    KD(42) = "42 Tecla IMPRIMIR PANTALLA"
    KD(43) = "43 Tecla EXECUTE"
    KD(44) = "44 Tecla SNAPSHOT"
    KD(45) = "45 Tecla INS"
    KD(46) = "46 Tecla SUPR"
    KD(47) = "47 Tecla AYUDA"
    KD(48) = "48 Tecla 0 (COMUN)"
    KD(49) = "49 Tecla 1 (COMUN)"
    KD(50) = "50 Tecla 2 (COMUN)"
    KD(51) = "51 Tecla 3 (COMUN)"
    KD(52) = "52 Tecla 4 (COMUN)"
    KD(53) = "53 Tecla 5 (COMUN)"
    KD(54) = "54 Tecla 6 (COMUN)"
    KD(55) = "55 Tecla 7 (COMUN)"
    KD(56) = "56 Tecla 8 (COMUN)"
    KD(57) = "57 Tecla 9 (COMUN)"
    
    KD(65) = "65 Tecla A"
    KD(66) = "66 Tecla B"
    KD(67) = "67 Tecla C"
    KD(68) = "68 Tecla D"
    KD(69) = "69 Tecla E"
    KD(70) = "70 Tecla F"
    KD(71) = "71 Tecla G"
    KD(72) = "72 Tecla H"
    KD(73) = "73 Tecla I"
    KD(74) = "74 Tecla J"
    KD(75) = "75 Tecla K"
    KD(76) = "76 Tecla L"
    KD(77) = "77 Tecla M"
    KD(78) = "78 Tecla N"
    KD(79) = "79 Tecla O"
    KD(80) = "80 Tecla P"
    KD(81) = "81 Tecla Q"
    KD(82) = "82 Tecla R"
    KD(83) = "83 Tecla S"
    KD(84) = "84 Tecla T"
    KD(85) = "85 Tecla U"
    KD(86) = "86 Tecla V"
    KD(87) = "87 Tecla W"
    KD(88) = "88 Tecla X"
    KD(89) = "89 Tecla Y"
    KD(90) = "90 Tecla Z"
    
    KD(96) = "96 Tecla 0 (TECLADO NUMERICO)"
    KD(97) = "97 Tecla 1 (TECLADO NUMERICO)"
    KD(98) = "98 Tecla 2 (TECLADO NUMERICO)"
    KD(99) = "99 Tecla 3 (TECLADO NUMERICO)"
    KD(100) = "100 Tecla 4 (TECLADO NUMERICO)"
    KD(101) = "101 Tecla 5 (TECLADO NUMERICO)"
    KD(102) = "102 Tecla 6 (TECLADO NUMERICO)"
    KD(103) = "103 Tecla 7 (TECLADO NUMERICO)"
    KD(104) = "104 Tecla 8 (TECLADO NUMERICO)"
    KD(105) = "105 Tecla 9 (TECLADO NUMERICO)"
    KD(106) = "106 Tecla * (TECLADO NUMERICO)"
    KD(107) = "107 Tecla SIGNO MÁS (+) (TECLADO NUMERICO)"
    KD(108) = "108 Tecla INTRO (TECLADO NUMERICO)"
    KD(109) = "109 Tecla SIGNO MENOS (-) (TECLADO NUMERICO)"
    KD(110) = "110 Tecla PUNTO DECIMAL (.) (TECLADO NUMERICO)"
    KD(111) = "111 Tecla SIGNO DE DIVISIÓN (/) (TECLADO NUMERICO)"
    KD(112) = "112 Tecla F1"
    KD(113) = "113 Tecla F2"
    KD(114) = "114 Tecla F3"
    KD(115) = "115 Tecla F4"
    KD(116) = "116 Tecla F5"
    KD(117) = "117 Tecla F6"
    KD(118) = "118 Tecla F7"
    KD(119) = "119 Tecla F8"
    KD(120) = "120 Tecla F9"
    KD(121) = "121 Tecla F10"
    KD(122) = "122 Tecla F11"
    KD(123) = "123 Tecla F12"
    KD(124) = "124 Tecla F13"
    KD(125) = "125 Tecla F14"
    KD(126) = "126 Tecla F15"
    KD(127) = "127 Tecla F16"
    
    KD(144) = "144 Tecla BLOQ NUM"
End Sub

Private Sub Class_Initialize()
    ReDim Ks(0)
    mArchivo = ""
    IniKD
End Sub

Public Sub ShowCFG(hTK As cl_tbrSoftSelectKey)
    FCfg.ShowCFG hTK
End Sub
