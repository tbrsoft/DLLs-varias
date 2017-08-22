VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKEYS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private fso As New Scripting.FileSystemObject
Private nFOt As New clsTODO

Private lFiles As New Load_Files 'cargador de archivos con claves e indices de licencia

Dim CC As New tbrCrypto.Crypt

'tipos de licencia que se pueden cargar
Public Enum TypeLic
    LicenciaVencida = -4
    ArchivoNoValidoComoLic = -3
    ParaOtraPC = -2 'no fue hecha para este equipo
    a_SinLeer = -1
    aSinCargar = 0 'todavia no puso nada
    BErronea = 1 'cuando ya erro 3 o mas veces
    CGratuita = 2 'primera que me puede pedir
    DMinima = 3
    EComun = 4
    FPremium = 5
    GFull = 6
    Supsabseee = 7
End Enum

Private LaLi As TypeLic 'basica de 3pm 'EX mLicencia, le cambie el nombre

'**************************************************************
'valores que esta bueno saber en todo momento una vez leido
'todos referidos a validacion por fecha o versión de licencia
Private nnFecha As Long 'fecha del archivo L37
Private nnVers As Long
Private nnVers2 As Long 'version que estoy usando
Private nnFecha2 As Long
Private difXaSoporte As Long
Private difXaVersion As Long
'**************************************************************
Private lista3(9, 5) As String 'cada tipo de licencia se diferencia por el GetNR que devuelve
'cada soft que use esta dll elige los codigos que desea usar
'el primer indice es el tipo de licencia y cada uno tiene hasta 5 codigos diferentes para acceder a esa misma licencia
'si alguno es cero no se usa
'el valor guardado es el mnr que devuleve al leer el archivo de licencia

Private Sub Class_Initialize()
    LaLi = a_SinLeer
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
End Sub

'creo un archivo de licencia en la PC actual para el soft (nombre y version) indicado
'Create File Lic
Public Function CFL(filSave As String, forSF As String, VersSF As Long, filLog As String)
    On Local Error GoTo errFL
    'generar el archivo de esta PC !
    
    nFOt.SetLog filLog
    Terr.Anotar "IC12ff"
    
    nFOt.SetSF forSF 'nuevo agosto 2007 para no mezclar con karaokes ni con programas de artime y manu
    nFOt.SetSV VersSF
    nFOt.SetFH ((Year(Date) - 2000) * 10000) + (Month(Date) * 100) + (Day(Date))
    
    CFL = nFOt.CargarFT(filSave)
    
    Exit Function
    
errFL:
    Terr.AppendLog "errFL", Terr.ErrToTXT(Err)
    CFL = Err.Number
End Function

'Insert File Lic
Public Sub IFL(sFile As String, forSF As String, VersSF As Long, Optional LogOnErr As Boolean = False)
    On Local Error GoTo ErrIF
    
    If fso.FileExists(sFile) Then
        LaLi = shek(sFile, forSF, VersSF, LogOnErr) 'este es el UNICO lugar donde reviso la licencia directo al archivo
    Else
        LaLi = aSinCargar
    End If
    
    Dim j1 As Long, j2 As Long
    Randomize
    j1 = ((CLng(Rnd * 80)) + 10)
    Randomize
    j2 = ((CLng(Rnd * 80)) + 10)
    
    'son 5 digitos 2 al azar, la licencia (lic+1 puede ser) y dos mas al azar
    If j1 > 50 Then
        Terr.AppendSinHist "tret-12:" + CStr(j1) + CStr(CLng(LaLi) + 1) + CStr(j2)
    Else
        Terr.AppendSinHist "tret-12:" + CStr(j1) + CStr(CLng(LaLi)) + CStr(j2)
    End If
    
    Exit Sub
    
ErrIF:
    Terr.AppendLog "errIF", Terr.ErrToTXT(Err)
End Sub

'Get Licencia
Public Function GL() As TypeLic
    GL = LaLi
End Function

'la siguiente funcion recibe un archivo de licencia y la indicacion de para que software o plugin
'fue hecho y en base a eso da como resultado el tipo de licencia que corresponde

Private Function shek(ArchTest As String, sForSF As String, VersSF As Long, Optional EsErrorSiNoVa As Boolean = False) As TypeLic 'ex garrala

    Terr.Anotar "setML", ArchTest, sForSF
    
    'ahora viene encriptado ! nuevo jul 09
    CC.DecryptFile eMC_Blowfish, ArchTest, AP + "h31.cda", nFOt.GetStrError

    Dim TEZ As String 'lista "FINAL:" original
    Dim mNR As Long
    mNR = nFOt.GetNR(AP + "h31.cda", TEZ)    'el tez queda en ni bosta!!!
    
    'ya pluedo borrar el desencriptado que me botonea como es el archivo
    fso.DeleteFile AP + "h31.cda", True
    
    Dim forSF As String 'saber para que soft fue hecho
    forSF = nFOt.GetSF
    Terr.Anotar "forsf", forSF
    
    If forSF <> "" And LCase(forSF) <> LCase(sForSF) Then
        ' "" van a dar las anteriores supuestamente
        ' las mantengo por compatibilidad.
        mNR = -2333
        Terr.Anotar "IC05a", forSF, sForSF
    End If
    
    Terr.Anotar "IC05", mNR
    'si da -2 es que puso un archivo que NO es de licencia
        
    Dim TmpLic As TypeLic
    'si mnr es <0 no sirve el archivo !!!
    If mNR < 0 Then
        TmpLic = ArchivoNoValidoComoLic
        If EsErrorSiNoVa Then Terr.AppendLog "IC09c"
        shek = TmpLic
        Exit Function
    End If
    
    '///////////////////////////////////////////////////////////////////////
    'verque la licencia no este vencida
    
    Dim resVersionAndTime As String
    'set 08 nuevo
    'getNR carga si lo tuviera el L37 la fecha y version de la licencia
    
    nnFecha = nFOt.GetFH
    resVersionAndTime = "FH:" + CStr(nnFecha)
    Terr.Anotar "IC09L", "FH:" + CStr(nnFecha)
    
    'VER QUE VERSION ES la que tiene el archivo de licencia!
    nnVers = nFOt.getSV
    resVersionAndTime = resVersionAndTime + vbCrLf + "SV:" + CStr(nnVers)
    Terr.Anotar "IC09m", "SV:" + CStr(nnVers)
    
    nnVers2 = VersSF 'la mandan siempre desde afuera, si la tomo de aqui sera la version de la dll y estaria MAL
    resVersionAndTime = resVersionAndTime + vbCrLf + "SV2:" + CStr(nnVers2)
    Terr.Anotar "IC09n", "SV2:" + CStr(nnVers2)
    
    
    nnFecha2 = (Year(Date) - 2000) * 10000 'año 2008 es = 80.000
    nnFecha2 = nnFecha2 + Month(Date) * 100 'setiembre de 2008 seria 80.000 + 900 = 80.900
    nnFecha2 = nnFecha2 + Day(Date) '23 set 08 = 80.923
    resVersionAndTime = resVersionAndTime + vbCrLf + "FH2:" + CStr(nnFecha2)
    Terr.Anotar "IC09o", "FH2:" + CStr(nnFecha2)
    
    resVersionAndTime = resVersionAndTime + vbCrLf + "FH3:" + CStr(nnFecha)
    Terr.Anotar "IC09p", "FH3:" + CStr(nnFecha)
    
    
    difXaSoporte = nnFecha2 - nnFecha
    resVersionAndTime = resVersionAndTime + vbCrLf + "DfSupp:" + CStr(difXaSoporte)
    Terr.Anotar "IC09q", "DfSupp:" + CStr(difXaSoporte)
    
    If difXaSoporte < 0 Then
        'el tipo tiene problemas con el reloj o me quiere cagar
        'o cuando se creo el archivo de licencia se creo con fecha del futuro por error
        Terr.AppendSinHist "SUPP0089" + vbCrLf + resVersionAndTime
    End If
    
    If difXaSoporte > 10000 Then
        'ya no tiene soporte'
        Terr.AppendSinHist "SUPP0099" + vbCrLf + resVersionAndTime
    End If
        
    difXaVersion = nnVers2 - nnVers '(imposible menor que cero!!)
    resVersionAndTime = resVersionAndTime + vbCrLf + "DfVs:" + CStr(difXaVersion)
    Terr.Anotar "IC09r", "DfVs:" + CStr(difXaVersion)
    
    If difXaVersion > 100000 Then
        'ya no tiene actualziación!!!
        'SU LICENCIA DE 3PM LO CUBRE POR ACTUALIZACIONES HASTA LA VERSION "nnvers"+100.000 y usted esta usando "nnvers2"
        TmpLic = LicenciaVencida
        If EsErrorSiNoVa Then Terr.AppendLog "IC09k"
        GoTo FIM
    End If
    
    'los limites de 3PM son 1 punto de version major (la fecha como los relojes se cambian facil la ignoro)
    
    '///////////////////////////////////////////////////////////////////////
    TEZ = Trim(UCase(TEZ))
    Terr.Anotar "IC06", TEZ 'final del L37 abierto
    
    Dim FFdeLaClave As String
    FFdeLaClave = nFOt.GetRF 'da ok el de esta misma PC
    
    FFdeLaClave = Trim(UCase(FFdeLaClave))
    Terr.Anotar "IC07", FFdeLaClave 'final de esta PC
        
    If TEZ <> FFdeLaClave Then
    
        'VER SI NO SON TAN DISTINTAS LAS 2 PCs
        Dim DF As String
        'primero esta pc y 2º para la que fue hecha la clave
        DF = nFOt.GetDiff2PC(FFdeLaClave, TEZ)
        
        Dim Ret(20) As Long
        Dim L As Long
        For L = 1 To Len(DF)
            Ret(L - 1) = CLng(Mid(DF, L, 1))
        Next L
        
        Dim PT1 As Long 'puntos para medir la igualdad de 2 equipos
        'ret(0):si es cero la PC1 tiene el minimo de indices necesarios para arrancar
        'ret(1):si es cero la PC2 tiene el minimo de indices necesarios para arrancar
        
        'ret(2):cantidad de placas de red de la PC1
        'ret(3):cantidad de placas de red de la PC2
        'ret(4):cantidad de coindicencias en placas de red
        
        'ret(5):dif en la bios si es 4 esta ok, si es 6 esta mal. NO puede ser otro resultado
        'ret(6):dif en el micro si es 2 esta ok, si es 1 esta mal. NO puede ser otro resultado
        'ret(7):dif en el micro reserved si es 5 esta ok, si es 3 esta mal. NO puede ser otro resultado
        
        'ret(8):cantidad de discos en la pc1
        'ret(9):cantidad de discos en la pc2
        'ret(10):coincidencias en los discos
        
        '00 111 4 1 3 111 0000000000
        '001921F5682A|00000000|BFEBFBFF00000F49|409000F|WD-WMAM98125093|
        '001921F5682A|00000000|BFEBFBFF00000F64|604000F|WD-WMAM98125093|
        If Ret(0) = 1 Then PT1 = 0: GoTo VALIDAR
        If Ret(1) = 1 Then PT1 = 0: GoTo VALIDAR
        
        If Ret(2) = Ret(3) Then PT1 = PT1 + 10
        PT1 = PT1 + (Ret(4) * 50)
        
        If Ret(5) = 4 Then PT1 = PT1 + 20
        If Ret(6) = 2 Then PT1 = PT1 + 20
        If Ret(7) = 5 Then PT1 = PT1 + 20
        
        If Ret(8) = Ret(9) Then PT1 = PT1 + 10
        PT1 = PT1 + (Ret(10) * 50)
        
        Dim CM As String 'COMO MOSTRAR
        CM = "DF:" + DF + vbCrLf + _
            "PT:" + CStr(PT1) + vbCrLf
            
        Dim H As Long
        Dim TJ As String 'todo junto
        TJ = "B" + TEZ + "WEX" + FFdeLaClave
        Dim Mx As Long
        Mx = Len(TJ)
        
        For H = 1 To 200
            If (H < 27) Or (H > (Mx + 27)) Then
                Randomize
                CM = CM + CStr(Int(Rnd * 7))
                CM = CM + CStr(Int(Rnd * 7))
            Else
                If (H - 27) <= Len(TJ) Then CM = CM + Mid(TJ, (H - 26), 1)
            End If
        Next H
        'que se esconda bien
        CM = Replace(CM, "|", "%")
        
        If EsErrorSiNoVa Then Terr.AppendSinHist CM
        
VALIDAR:
        If PT1 >= 100 Then
            GoTo ValeTAMBIEN 'son parecidas las pcs
        Else
            TmpLic = ParaOtraPC
            If EsErrorSiNoVa Then Terr.AppendLog "IC08" 'para que me envien el registro en cualquier caso
        End If
    Else 'el archivo fue hecho para esta PC (SON IGUALES EL LIC y el L38)
    
ValeTAMBIEN:
    
        Terr.Anotar "IC09e", mNR
        If mNR < 0 Then
            TmpLic = ArchivoNoValidoComoLic
            If EsErrorSiNoVa Then Terr.AppendLog "IC09f"
        End If
        
        Dim Encontro As Boolean, i1 As Long, i2 As Long
        Encontro = False
        
        For i1 = 2 To 7
            For i2 = 1 To 5
                'agregado nov 09 trabajando en licenciero, da erro de que no coinciden los tipos
                If lista3(i1, i2) = "" Or Not IsNumeric(lista3(i1, i2)) Then
                    lista3(i1, i2) = "0"
                End If
                If mNR = CLng(lista3(i1, i2)) Then
                    TmpLic = i1
                    Encontro = True
                    Exit For
                End If
            Next i2
        Next i1
            
        If Encontro = False Then
            TmpLic = BErronea
            If EsErrorSiNoVa Then Terr.AppendLog "IC10." + CStr(mNR)
        End If
        
    End If
    
FIM:
    '**********************
    shek = TmpLic
    '**********************
End Function

'setear los codigos de cada tipo de licencia
Public Sub setID(i1 As Long, i2 As Long, i3 As Long)
    lista3(i1, i2) = i3
    'seguir aqui que no haya duplicados !!! (en ese caso no se tomara en cuenta el primero que se haya escrito)
    
End Sub

Public Function GETnnFecha() As Long 'fecha del archivo L37
    GETnnFecha = nnFecha
End Function

Public Function GETnnFecha2() As Long 'fecha del archivo L37
    GETnnFecha2 = nnFecha2
End Function

Public Function GETnnVers() As Long 'fecha del archivo L37
    GETnnVers = nnVers
End Function

Public Function GETnnVers2() As Long 'fecha del archivo L37
    GETnnVers2 = nnVers2
End Function

Public Function GETdifXaSoporte() As Long 'fecha del archivo L37
    GETdifXaSoporte = difXaSoporte
End Function

Public Function GETdifXaVersion() As Long 'fecha del archivo L37
    GETdifXaVersion = difXaVersion
End Function

Public Sub SetLog(f As String)
    nFOt.SetLog f
End Sub

'traduce un archivo de licencia
Public Function Tasl(AR As String, AR2 As String) As Long 'ex LiToLo
    On Local Error GoTo NoEsArchDe3PM2
    Dim CC As New tbrCrypto.Crypt
    
    CC.DecryptFile eMC_Blowfish, AR, AR2, nFOt.GetStrError
    Tasl = 0
    Exit Function
NoEsArchDe3PM2:
    Tasl = 1
End Function

Public Function Desuso() As clsTODO
    Set Desuso = nFOt
End Function

Public Function Desuso2() As Load_Files
    Set Desuso2 = lFiles
End Function

'cargar los indices y claves de un archivo externo (archivos LOAD)
Public Function SetIXs(sFile As String, UseIndexs As Long)

    'sFile es el archivo a leer
    'uSeIndex indica si en este objeto de licencias uso los Indices 1 o 2 (SetId1 o SetId2)
    'en el licenciero se usan dos objetos y por eso los archivos load traen dos
    
    lFiles.Unload__D sFile
    
    '********************************
    '********************************
    'cargar los indices leidos!!!
    Dim i1 As Long, i2 As Long, i3 As Long, sI3 As String
    
    For i1 = 2 To 7
        For i2 = 1 To 5
        
            If UseIndexs = 1 Then
                sI3 = lFiles.GetID1s(i1, i2)
            Else
                sI3 = lFiles.GetID2s(i1, i2)
            End If
            
            If sI3 = "" Then
                i3 = 0
            Else
                If Not IsNumeric(sI3) Then
                    Terr.AppendLog "NNUM_i3.1." + CStr(sI3) + "." + CStr(i1) + "." + CStr(i2)
                    i3 = 0
                Else
                    i3 = CLng(sI3)
                End If
            End If
            
            setID i1, i2, i3
        Next i2
    Next i1
    
    '********************************
    '********************************
    
    'cargar tambien la clave de encriptacion
    'estas en realidad son las cadenas que se usan para encriptar todo el manejo interno
    'si se cambian todas las licencias anteriores dejaran de ser compatibles. NO TOCAR !
    
    If UseIndexs = 1 Then
        Desuso.SetStrError lFiles.LL1   'saca el LLAVE1 del archivo load que se refiere a la encriptacion
        'aqui estoy con una licencia comun
    Else
        Desuso.SetStrError lFiles.LL2   'saca el LLAVE2 del archivo load que se refiere a la encriptacion
        'aqui estoy en licencia del licenciero
    End If
    
End Function

'matar la licencia por que se mando una cagada
Public Sub UpMng()
    LaLi = BErronea
End Sub

'setear la licencia segun interfase y soft
'por ejemplo si tiene una interfase valida para el soft!
'lo pongo todo junto para que haya un solo lugar donde manejar todo el kilombo de codigos de interfases de todos los softwares
Public Sub Set2(Sf7 As String, cOd7 As String)
    'codigos de los programas encriptados
    'genesis-mp6 = 35yfsAKVETNRy2KO99AfqmrB3YO5nwko
    'e2games = no implementado por que no usa interfases
    '3pm = ZUPZP4Pq0ylE/pIFZgU24g==
    'mpRock =
    'multitechgames =
    'meridia =
    'rockotech MOSSE [interfases h3k] = XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B
    
    'INLCUIR LOS CODIGOS DE LA INTERFASE H3K
    
    'este es como un case else para todas las que no anote en 3pm
    'If cod = 3333302 And Sf7 = "3pm" Then LaLi = Supsabseee
    
    Terr.Anotar "set2", cOd7, Sf7
    
    If cOd7 = "0" Then Exit Sub
    
    'traidos desd 3PM v7 TODO cargar
    'cargo el primero para asentar com va y depues hacer los demas
    
    '41837215 'cesar pizarro BsAs activada para 3PM v7, esta vencida
    If cOd7 = dcr("mv+UQuihqouU3bLBRbEP2YI6wh4duFV1") And LCase(Sf7) = dcr("ZUPZP4Pq0ylE/pIFZgU24g==") Then LaLi = LicenciaVencida 'no se actualizo a la 8
    '45757651 'diego fontana buenos aires (mLicencia = GFull)
    '46523511 'sergio soza (fue en un paquete con 3 mas. Esta es la única con karaoke)
    '47536194 'sergio soza la mando para que le pongamos karaoke
    '49509525 'guillermo rosario USB con expender y todo
    '49547493 'Alejandro molina se lo llevo del taller
    '50072908 'rocchio
    '54144024 'ale molina
    '55923640 'ale molina
    '56634485 'ale molina
    '55287301 'cesar pizarro
    '68838655 'la trajo anzil gustavo para reparar
    '70108830 'oscar brito licencia simple
    '82453846, 86542785, 85721433, 85620164, 87381243, 89463531 'rodolfo martino (justiniano pose / bell ville)  'tenia kar y se arrepintio al final ? compro solo expendedores y un karaoke
    '85888147 'oscar brito licencia simple
    '90566933 'martino pide especial con karaoke
    '95641791 ' gruccos 14/11/08
    '90609013 'dario settimo (antes era solo t-store, ahora pago la diferencia y es con expendedor)
    '93685209 'roberto abrahan bsas
    '94939172 ' ariel schell compro interfase para un equipo que tenia lic de L37, en enero de 2010 se le activo la interfase igual
    '93549609 'creo que es del ale veiga (kar incluido)
    '93633532 ' carlos salas compro interfase para un equipo que ya tenia licencia de archivo
    '93934714 'carlos salas t-Store ! NO 3PM
    '96104800 'antes era de tStore para chicago pero la cambio por 3pm
    '9651396 'las chelaz bsas Juan Carlos Mendoza
    '96369266 'oscar britto chaco 16/01/2009
    '15 completas para victor grande de arias
    '102996647, 101296007, 119202121, 97670972, 102271795, 102872252,98369777, 99625825, 98225533, 101871311, 99924147, 99354684, 103290080,100677430, 98804175, 100154459
    '97670972 'ahora esta en manos de raul pogonza, estaba registrada como "grande de arias" ??? puede ser que luis haya hecho dos con el mismo numero ?????????????    'si luis estuvo arreglando apurado y saco algunas con numeros repetidos
    '104329645, 103389501, 107510805 'martino, solo mpRock-(genesis-MP6) [14 abr 2009]
    '90213537 'PRUEBA EXPENDEDOR, no se adonde fue
    '105929791 'gustavo lopez rosario, full chiche exp y kar
    '106218576 'ale veiga
    '109995186 'dardo paso de mpRock a 3Pm y la actualizamos aqui a full
    '109529197 'ale veiga 1/2 activacion de extras
    '110959589 ''dardo paso de mpRock a 3Pm y la actualizamos aqui a full
    '111457645 'no la probe yo pero fernando tecnico de mose me dice que es una que esta en jamaica
    '130747914 'mosse 6 botones sin retrdador (del primer grupo de 4) 05/01/2010
    '130199372 'mosse 6 botones sin retrdador (del primer grupo de 4) 05/01/2010
    '131091713 'mosse 6 botones sin retrdador (del primer grupo de 4) 05/01/2010
    '131210625 'mosse 6 botones sin retrdador (del primer grupo de 4) 05/01/2010
    
    '104329645 '1/10 dic 08 martino H2k
    If cOd7 = dcr("RiZ1bEFVeLNS4E7m0dWMwXSCOhw0D/YH") And LCase(Sf7) = LCase(dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko")) Then LaLi = Supsabseee
    '103389501 '2/10 dic 08 martino H2k
    If cOd7 = dcr("3Oiqmf6Xkir9bSvHfj18o50OPvSwwDEL") And LCase(Sf7) = LCase(dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko")) Then LaLi = Supsabseee
    '107510805 '3/10 dic 08 martino H2k
    If cOd7 = dcr("afmxbwcoczuTN6Mdf2O31J0wNquUo8E7") And LCase(Sf7) = LCase(dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko")) Then LaLi = Supsabseee
    
    '1003793314 'PRIMERA interfase ILUS para mosse
    If cOd7 = dcr("4//h5UMAgpfMwQBcuAIVhvwXye1RpapU") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1008351066 ' interfase ILUS para mosse
    If cOd7 = dcr("MbbJbG2d1z1ND/Nv8pCq0t/TrsTgmBxp") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1014563174 interfase ILUS para mosse
    If cOd7 = dcr("OJaauC5hB5BneUFeVM1D7G5phzDYkuMt") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1016076852 ' 2º interfase ILUS para mosse
    If cOd7 = dcr("mgz1XhKNs9RIelYOxJ9MUt1LHfKTVk0S") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1016837387 ' 2º interfase ILUS para mosse
    If cOd7 = dcr("r+V6OqypdPC2DvF4PEKIVyZuesoNr0zQ") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1018360390 interfase ILUS para mosse
    If cOd7 = dcr("FnKNcwZ8V5n9d3ZrfONOo6jOWKFVb+FC") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1021001897 interfase ILUS para mosse
    If cOd7 = dcr("+FBV1lyRUUhO9dqNrbjP+2/91OVYayuE") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1023049496 interfase ILUS para mosse
    If cOd7 = dcr("MVPNJgIc82/ZYdx3IbaLkTO9jzKFHh7L") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1024649665 ' 2º interfase ILUS para mosse
    If cOd7 = dcr("3fANdCUYy2KGVFEf/Twp4+AlV51+ll++") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    '1031011222 interfase ILUS para mosse
    If cOd7 = dcr("sN0eLdUJW0tFNnlt12VbBt5ci4ug77ir") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee
    
    '1031850856 interfase ILUS para vicente ezio wiskeria arroyito
    If cOd7 = dcr("IedAsPVggeQotO9e6kEUH40Wj9c8KUz8") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1037616572 gustavo anzil cambio por interface H2K que le vendimos abril 2010
    If cOd7 = dcr("vX1AznsJHMGZZf4rvqvFtIwu7Myez0qk") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1038456206 jose aguirre compra 2 juntas 12/05/2010
    If cOd7 = dcr("XoieJsVh2rmwUGY0B6vUGj21TP6A92aB") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    '1039216741 jose aguirre compra 2 juntas 12/05/2010
    If cOd7 = dcr("atyCotv5GfVWewEWbW02jGgYcwzyfs69") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1041781082 gustavo lago parana entre rios 28/05/2010
    If cOd7 = dcr("5Em22lQfhCiUreis1cS1hTYn2UW0h/gg") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1044589216 gustavo anzil 28/05/2010
    If cOd7 = dcr("UwEnSZwTs3QJyiQ/DiknZ8+w+DeRLbjq") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1044675707 gustavo anzil 28/05/2010
    If cOd7 = dcr("TyuRiTt7dkJ5Ti2iyaJxFoGQRdqOt3LP") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1057633289 el luis se lo llevom a pogonza/molina 17/06/2010
    If cOd7 = dcr("Rp2MAwKJ5jbIRiKmr6n6diT5h7TpoIO/") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1058482248 anzil devolvio 2 de cajitas negras y le damos dos de las nuevas cajas chiquitas '20/06/2010
    If cOd7 = dcr("LXiILxuW+ITbH9xyBE8DvU4NPVJpMOh8") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1065454892  el luis se lo llevom a pogonza/molina 17/06/2010
    If cOd7 = dcr("l0CHaz5ir4LyN0IykK9I7PZ1zza1Y6hC") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
    
    '1066889471 anzil devolvio 2 de cajitas negras y le damos dos de las nuevas cajas chiquitas '20/06/2010
    If cOd7 = dcr("ZvOE7QGSMHoaj38ZY2GONF59CZyn6q8+") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee

    '1039216741 AUTOM. 3PM
If cOd7 = dcr("9CVIZwH++7tAuFcAZLFX584TVLDRVGC1") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1051107038 AUTOM. 3PM
If cOd7 = dcr("EWNTHX+5NGqXNWTXzC6bQ5Wq1A+Utmxl") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1053075538 AUTOM. 3PM
If cOd7 = dcr("iSodNJYJNjoM+bDHnpoxmYo3tBueOMN8") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1053162029 AUTOM. 3PM
If cOd7 = dcr("fslqrEGVobLfYwDWXSGJDijulNK7Mvzx") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1056198710 AUTOM. 3PM
If cOd7 = dcr("Btl+KX7SDKRnmvT3R7TndLyHS+BtXixI") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1056872754 AUTOM. 3PM
If cOd7 = dcr("Yy7qLqRIgCkRXeim29UZLHoEx1arN+MW") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1068403149 AUTOM. 3PM
If cOd7 = dcr("udS2SzgxP/EIAaECQOz+ptpha3wrWHqW") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee
'la tengo probando en mi casa y quiero que tambien sirva para rockotech
If cOd7 = dcr("udS2SzgxP/EIAaECQOz+ptpha3wrWHqW") And LCase(Sf7) = LCase(dcr("XCt37n/ktRy69HWMs+qJMjdtUQ3XFR1B")) Then LaLi = Supsabseee

'1076215427 AUTOM. 3PM
If cOd7 = dcr("stBqrpBxfyyAGN3ltNybxFNX7DVg08Cs") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1083188070 AUTOM. 3PM
If cOd7 = dcr("rCsZmVJxTPLTaDGSIlzEwwWYbUfNZxAt") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1083871439 AUTOM. 3PM
If cOd7 = dcr("7Bu/b42B4ocRcFBFSUVuQhrF1A9Pg0ju") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1085471608 AUTOM. 3PM
If cOd7 = dcr("l3m7BGQjiI26SPu8rzD+5xljOm1xN5Ka") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1087668655 AUTOM. 3PM
If cOd7 = dcr("t/kEBS1FwCpShpTQmrOe3wNvtAQ4IPlD") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1089182334 AUTOM. 3PM
If cOd7 = dcr("Ekp1d8MtAfjYSk+GbLg31+1bJnn7a0Od") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1103127621 AUTOM. 3PM
If cOd7 = dcr("/sRccdZbkKPZTWECzLQ/ae+sBS7nSGwp") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1108219293 AUTOM. 3PM
If cOd7 = dcr("pEsL4bOYW4U1xhvM95XCBlqOF7XTlHUA") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1118989153 AUTOM. 3PM
If cOd7 = dcr("wtgaAkBtXCkSnDSCQmOFw/z4x/+2o3VP") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1121630660 AUTOM. 3PM
If cOd7 = dcr("apEhGJIEf0kFvjEzhq/rwP7uxPKEmFMW") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1126950879 AUTOM. 3PM
If cOd7 = dcr("LZDSqc7Psg1Qyid3FT49AJzt3Io3LUkM") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1128245335 AUTOM. 3PM
If cOd7 = dcr("IVkaZzXNIiXSLmZg6vzWMUq1TMQ3qsrz") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1133240154 AUTOM. 3PM
If cOd7 = dcr("2Zxnbp/+v/GKudqvempc/hxkar+DVPRn") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1136197736 AUTOM. 3PM
If cOd7 = dcr("tNIC4MZU8Sr6wCxoSwZSlS+xQ444gW3H") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1137037370 AUTOM. 3PM
If cOd7 = dcr("PwrE3CMZDV6kSTT/AplzbfQ7wktW6/5m") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1144693383 AUTOM. 3PM
If cOd7 = dcr("rbII9fuKNzrWTRmbCCjOsOuWUC/kDDMj") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1155987839 AUTOM. 3PM
If cOd7 = dcr("/wxVvsxNH6teg/p9HDjzsPdPtA9mazYh") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1159792447 AUTOM. 3PM
If cOd7 = dcr("NR8q+N2rbFppYuWnIJ/OTmDPppxJFqJ4") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1165549733 AUTOM. 3PM
If cOd7 = dcr("zUIHxQTjErRsPbKWLM4x/UsY/C44cXtK") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1166233102 AUTOM. 3PM
If cOd7 = dcr("4tEXcPfqQXW+2C22OLvJBoYoZgQ2VijD") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1168507316 AUTOM. 3PM
If cOd7 = dcr("LI4qbXf0688X8RShKGMmFNlpmOxoOYmJ") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1172435886 AUTOM. 3PM
If cOd7 = dcr("a0bb0IUHGBzxxJgFAKx/Act/CvnbP/0j") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1173196421 AUTOM. 3PM
If cOd7 = dcr("do3+DGq5rbT0pO/iGTreh9Bcbsv4WiJP") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1176013880 AUTOM. 3PM
If cOd7 = dcr("uf4FbFm5uxMAmDt1azQ65qPZgbqvAdEq") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1176993637 AUTOM. 3PM
If cOd7 = dcr("5zHB5gqAkrfwm27KMTMo99+SMkTJkbX+") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1193975606 AUTOM. 3PM
If cOd7 = dcr("UsEqLGkCN+6n1SNV7XDKLIZ3nK+EM/dE") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1194516919 AUTOM. 3PM
If cOd7 = dcr("DohhgTJ1MQ/+9EiBjYjtIOWWKlYlfU24") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1203231788 AUTOM. 3PM
If cOd7 = dcr("pkRTTL2dYTTuxJR/sfLsMis4tyzCtPyl") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1204745466 AUTOM. 3PM
If cOd7 = dcr("/JH6RD892LZe7UC2yNZWhm6sYbdFJnxd") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1204831957 AUTOM. 3PM
If cOd7 = dcr("Djc29jE9Frqi25RK4wnyfiiucEjx9tkh") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1205286779 AUTOM. 3PM
If cOd7 = dcr("7YX7u9yQm59m/kCFJnIvKrGRGXu5J/DG") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1213773101 AUTOM. 3PM
If cOd7 = dcr("7HujiJDPYCkCUumaDdBND/ByUZAZGo/Y") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1224542961 AUTOM. 3PM
If cOd7 = dcr("D2kt67qxvVoggl3wxGKfaVd5kdURW7Lt") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1227666132 AUTOM. 3PM
If cOd7 = dcr("uKLlQZegcq3utUGP2JgjCemyqHbulYKo") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1229179811 AUTOM. 3PM
If cOd7 = dcr("CympuHC9iIjrvmhEhwJR7D/P/YE/nygK") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1229863180 AUTOM. 3PM
If cOd7 = dcr("4+tPqnG+DkQ1gOjrq3rCzjtoS8JAjqNS") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1233344320 AUTOM. 3PM
If cOd7 = dcr("6NGq9qimiIkUo3F3sSwRPbKgIypVZICt") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1235312821 AUTOM. 3PM
If cOd7 = dcr("Fad6NMJGh1Noc5oIk110QQjNTYgNfRf+") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1239643957 AUTOM. 3PM
If cOd7 = dcr("JBS3fUbbsb/dB70g1EanZ0oHs/3bzHvq") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1239870572 AUTOM. 3PM
If cOd7 = dcr("W4wu2VzSsnMZKM4AaK7M5KJZ3YbiMS+E") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1243799142 AUTOM. 3PM
If cOd7 = dcr("ZvVCl0Yrc1Pu5+ditG/Xb4CsUfZrSMrH") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1246922314 AUTOM. 3PM
If cOd7 = dcr("7Da+6IBhWn3/I0sRzkNni5TB4MF7Fwf0") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1249205852 AUTOM. 3PM
If cOd7 = dcr("W5V0MR5wR3/+rnpFyqgqGc0plMPebYpn") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1249879896 AUTOM. 3PM
If cOd7 = dcr("PsT8QLml5Syc/RrADRLnZGbffN9DMKV+") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1251086824 AUTOM. 3PM
If cOd7 = dcr("zDt54qdUphx6hhuV5m18u1LBKyW5FVnW") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1254578327 AUTOM. 3PM
If cOd7 = dcr("NffaIF74l2LvCk4FaYbfRNPmChshIjCd") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1256178496 AUTOM. 3PM
If cOd7 = dcr("O2uAJoa2IWLoX2nQcucbLoiFgPnTA44o") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1258900139 AUTOM. 3PM
If cOd7 = dcr("xjC/nT5/WhCrEQlkBVZG/0b9FVCus7Im") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1266861865 AUTOM. 3PM
If cOd7 = dcr("nAXQX/XbJ7JkYJW+AHylA3lk+BzjzLC6") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1269669999 AUTOM. 3PM
If cOd7 = dcr("RFbn4NWJNlbWf6w1El/QZvM/6Aeyzo8S") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1273834509 AUTOM. 3PM
If cOd7 = dcr("0BXaWSKvyv8mdn7UYt+Blc71xAEBqqos") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1274595044 AUTOM. 3PM
If cOd7 = dcr("/ZanoXhalirV+J8h+obd3uo8d/bxRlBy") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1276869257 AUTOM. 3PM
If cOd7 = dcr("GXMU9L+KfBbOPVF0YDd28K6t92bCdFHs") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1277631725 AUTOM. 3PM
If cOd7 = dcr("eOeLDQ3t9NQuW7PG9duq6y3lJs3vIkat") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1278926181 AUTOM. 3PM
If cOd7 = dcr("lyDcFIDOAXCPeCOhcgqFSaFDIVTsAdDg") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1282328223 AUTOM. 3PM
If cOd7 = dcr("RkVlfzMA7C0o6dwm0Hd+gopioPZv1p+d") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1297648442 AUTOM. 3PM
If cOd7 = dcr("DyNQ7pp12GtdBIfh8BHqBNJFxNMErX0L") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1302346873 AUTOM. 3PM
If cOd7 = dcr("hDth9yuykGGaTAGM7C3N5S0PO0WFLr4a") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1319722082 AUTOM. 3PM
If cOd7 = dcr("VzMiTnHtBFEyVwyjEgNLmVuSRP39AvOg") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1326921340 AUTOM. 3PM
If cOd7 = dcr("066ByN65RlGcGDUO3vOui6xmmZWHSIBg") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1327455261 AUTOM. 3PM
If cOd7 = dcr("VXoyy3xKcovHjL8sNhHFT508jpKDYlb5") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1335416987 AUTOM. 3PM
If cOd7 = dcr("G5tIaSRkcQycZcKEv+0uXOOkeQqeUDW4") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1343140841 AUTOM. 3PM
If cOd7 = dcr("v+2ZoBXCiY17hK6FngQpE93t01QJ3ksP") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1344663844 AUTOM. 3PM
If cOd7 = dcr("3YVJB4XEAax2XQoa90tp1a5DPD9a9Y3H") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1355207090 AUTOM. 3PM
If cOd7 = dcr("3v1+CVxI8w7dpb58tP5/DJGmiB0Klw5g") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1357481303 AUTOM. 3PM
If cOd7 = dcr("ghAkhC3+WW0Q6XRKCc+DgxCU7bdl1nXN") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1359002373 AUTOM. 3PM
If cOd7 = dcr("b9LYgUCMLKg9Dm6S/q9NpmUzcEqzXem7") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1363920026 AUTOM. 3PM
If cOd7 = dcr("NfG1h4mQNuB4qu9nLw1L8ynDl+wNblOt") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1369772233 AUTOM. 3PM
If cOd7 = dcr("qsnCGc4hojC/6GjdG8LQAUPBQaEK4TLc") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1371655138 AUTOM. 3PM
If cOd7 = dcr("tCddPGY5zdB5t7+Ns3Ukc70fqYAQ3msj") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1378267880 AUTOM. 3PM
If cOd7 = dcr("Dux40SyhjOxPBryX86FVFExxDmoIol5S") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1379028415 AUTOM. 3PM
If cOd7 = dcr("IUKvhV+C2sW46yucq/kaJU8RPuBCqtiB") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1382424998 AUTOM. 3PM
If cOd7 = dcr("waABz0j4ZCFZzHYwvkoJaRe3CL5bXMVL") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1384699211 AUTOM. 3PM
If cOd7 = dcr("tFUfG/Ib0dFdrTryyXCilq5zXjlQjlaG") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1386754202 AUTOM. 3PM
If cOd7 = dcr("TsZ6yt/hajYqZsPXJuCpMMu69WEEnYTr") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1389790883 AUTOM. 3PM
If cOd7 = dcr("mJ2AScjdZrBDHn802bOiPgAl4tgGUM1P") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1389798275 AUTOM. 3PM
If cOd7 = dcr("P0jySRT4mCA9va89W43iIlZV9s52uPcw") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1393194857 AUTOM. 3PM
If cOd7 = dcr("vCGaNHXNGduZlfY/dI1+sKoOssVL8yAj") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1395469071 AUTOM. 3PM
If cOd7 = dcr("nd95LaVqio/UozujufuS3KYAZpxkvkPj") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1399047064 AUTOM. 3PM
If cOd7 = dcr("XH10OdiKJtIS18V6ZJ4CK48FtjiJnopr") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1399807599 AUTOM. 3PM
If cOd7 = dcr("t0ClTrLefzJLyaumV7i2gjm96CbFXZkS") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1406780243 AUTOM. 3PM
If cOd7 = dcr("/8EC8buqICpUX1ZZPls97FqC14biwiCv") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1416789568 AUTOM. 3PM
If cOd7 = dcr("KS4u2UlxGhT58YvmHZ+hNy2wOXHoXeqy") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1419063781 AUTOM. 3PM
If cOd7 = dcr("fxwuk6qXKUftbAcDmk76n2PVBScZHfiR") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1419826249 AUTOM. 3PM
If cOd7 = dcr("Y3S+wIpFkMOS1Jspm7z4gzfQwtvut7kL") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1421478674 AUTOM. 3PM
If cOd7 = dcr("n8pdI3LhTvgbpU2Dc5ot+w3nU00iALkD") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1423762212 AUTOM. 3PM
If cOd7 = dcr("lW1AqC6FbvAa6q/Nx0IpDxKiOAieFBiX") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1439080498 AUTOM. 3PM
If cOd7 = dcr("ldUbhJb8WHeeoxKAXmDexIC7Jrn5xoI5") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1444541397 AUTOM. 3PM
If cOd7 = dcr("BEjqfN22bHV5VpxuVWnKbNx4g1em6KZ7") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1446053142 AUTOM. 3PM
If cOd7 = dcr("MRmrOnicChq9TdIjGKr6AMj67atFNur/") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1451504716 AUTOM. 3PM
If cOd7 = dcr("4kCIMH+9QJvyWs45jvhaJYDCQcWijYJE") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1455309324 AUTOM. 3PM
If cOd7 = dcr("Mdp4cp8zsHiKtSJq2xkX7Lsa1nSirJhy") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1456824935 AUTOM. 3PM
If cOd7 = dcr("0WhnNPDw5Wq1bmD2ztR8liMsWaf2jhSz") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1476088508 AUTOM. 3PM
If cOd7 = dcr("YClwhAfArgrlshJzU+nMxgxIVffibu1y") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1478136107 AUTOM. 3PM
If cOd7 = dcr("O7lwDLSEawd+tASB2wX7IlZR6i6tt0RJ") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1479580011 AUTOM. 3PM
If cOd7 = dcr("v+NIH1Kw9BDogegMEtN96E5OdK7In/Pq") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1484574830 AUTOM. 3PM
If cOd7 = dcr("oCmuZRCmXaqFgCFkU+PgtemogT3Ng3gN") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1485337298 AUTOM. 3PM
If cOd7 = dcr("ewW3hCKev4RpHNEoMMeXh06DtQT0RRTO") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1486097833 AUTOM. 3PM
If cOd7 = dcr("0oBS5YcwyuVwcegGtZBYImmvdivV/3CV") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1487611511 AUTOM. 3PM
If cOd7 = dcr("HK+PUHuITegID0iBI4HBR+TJDLgDnh8l") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1493297091 AUTOM. 3PM
If cOd7 = dcr("hCi3eYjx2Sn7aliwoRCy0WKlQ7yHHSfF") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1497628228 AUTOM. 3PM
If cOd7 = dcr("f+EeX4f+lhxPOn4hpUNdq1LbLH91/Cp2") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1498162149 AUTOM. 3PM
If cOd7 = dcr("2WUq/5p4RIPkkzkDyY+vEJmsR1FxNdiN") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1498388763 AUTOM. 3PM
If cOd7 = dcr("qTxWMwdJ8iQSdByATjB0chgCyl+40T4i") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1499141906 AUTOM. 3PM
If cOd7 = dcr("76YxZics70Ds0LAD9bBr3tfnFoDYRMi9") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1505354015 AUTOM. 3PM
If cOd7 = dcr("Uy2rdecArI7WVt11tfKAaHgzi8OlnZ9L") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1507637553 AUTOM. 3PM
If cOd7 = dcr("dsWWJUVFesXLr4HRtaZ3F1t6GWpVE5IM") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1509606053 AUTOM. 3PM
If cOd7 = dcr("MEe6B+SDwkNnhPbE8kUNMVDEch5cCYhx") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1526901127 AUTOM. 3PM
If cOd7 = dcr("Ximtbn8y9UrdK/1hR5DgMxAl20eliNFO") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1531503742 AUTOM. 3PM
If cOd7 = dcr("6Df/OWF26eH+FSJjbwFg/AZ0dD2+oDPd") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1538424130 AUTOM. 3PM
If cOd7 = dcr("uYAARBKkHk/BdbQM2ZJ8vZK7TFIFgDs7") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1539184665 AUTOM. 3PM
If cOd7 = dcr("+10p95SWNZqr1RM3Mi5EwWbnQIt2jnCY") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1541065637 AUTOM. 3PM
If cOd7 = dcr("zDsAz+SnHbyqGeoVZ4O5uwX1/Vb07M6q") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1543883095 AUTOM. 3PM
If cOd7 = dcr("zrleMZRMifE+84l1bAx2nppfLjuNMjBs") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1544636238 AUTOM. 3PM
If cOd7 = dcr("5CMCpzE73WVec4hOoFxaEMAF6r3B9wZs") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1551608882 AUTOM. 3PM
If cOd7 = dcr("zXyDsJEerGnbSmalrVhO7tgL8N3o/pTf") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1556927168 AUTOM. 3PM
If cOd7 = dcr("1M7fyZYwGgT8tVGIF5k8wUh8yTHaKHvf") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1558135134 AUTOM. 3PM
If cOd7 = dcr("QNjwG+ivM/DDurhbWwhnvFHZ8pA989Rr") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1558984092 AUTOM. 3PM
If cOd7 = dcr("OigFjlAdFoyI5OhJTNl5mvBE6vGLeXrG") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1562378742 AUTOM. 3PM
If cOd7 = dcr("dMK4PdTSO59k4NcdYu41VU/sVVrvMWE1") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1566936493 AUTOM. 3PM
If cOd7 = dcr("pCxo77jhV5kxu5FzsDlMHuie1RWP5Efo") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1571625599 AUTOM. 3PM
If cOd7 = dcr("8OU0qrI29Qw0jRT2ORvmR29IYkeUki2F") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1573148602 AUTOM. 3PM
If cOd7 = dcr("yyh0szaaYWuv7KSMeIPdsYPf+HSOuwTx") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1578161175 AUTOM. 3PM
If cOd7 = dcr("jD+LLX2DE89wmYSQBePz0RBruGuaGxPL") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1579753952 AUTOM. 3PM
If cOd7 = dcr("WpepQ8/mNm4sBQ/dobxy4yZaPkANUmms") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1583918462 AUTOM. 3PM
If cOd7 = dcr("c46R6vVNjbMH1P7TjpihnsRFTimwW+pa") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1585973453 AUTOM. 3PM
If cOd7 = dcr("/zbexRkWyLC/EW2ZbpqK2hxCM6eWGPOr") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1614399326 AUTOM. 3PM
If cOd7 = dcr("fXrmETb5rJ0sQytsJXwTXxbbbLTNhlY3") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1621458460 AUTOM. 3PM
If cOd7 = dcr("lsucOIqEbKQ2enoUabPZnZQyx9LztrWP") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1622972138 AUTOM. 3PM
If cOd7 = dcr("fbEtgQcJ7qZkZ44N1B99Jw4CjLLIiWNr") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1648309074 AUTOM. 3PM
If cOd7 = dcr("EK44dLXR8ulqrYv5AatPxVYL5ibK339n") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1651484502 AUTOM. 3PM
If cOd7 = dcr("AWFDRlnuAeopHXQOva5EoFkr4Bhq4NYB") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1655955762 AUTOM. 3PM
If cOd7 = dcr("p+WOKsCiOHxlQ5MpQ1bpIp0JSqhN21tk") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1659752978 AUTOM. 3PM
If cOd7 = dcr("TM5mBIL/4yeUYNC5yApWUJpUKDNDXPvf") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1666725622 AUTOM. 3PM
If cOd7 = dcr("CNocKpuhp/Xe95331QQCWrXOfqrlBPT4") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1668859712 AUTOM. 3PM
If cOd7 = dcr("//THVoHK+5NWYN/fi3AmZLeUBfbSnVh+") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1669762303 AUTOM. 3PM
If cOd7 = dcr("Aointvr5sp/jn/bCt2BqrqXAz65FBE2k") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1672184588 AUTOM. 3PM
If cOd7 = dcr("IiOZYt5hUZciT8VAVYgtvy2oeI++bK33") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1677495482 AUTOM. 3PM
If cOd7 = dcr("uP/x2DnlFqAg/TG/4pYo+rrczjJAbxTj") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1681659992 AUTOM. 3PM
If cOd7 = dcr("2SSgDUn3QLkR86o2ZXJSozyu7uEbToG7") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1688885754 AUTOM. 3PM
If cOd7 = dcr("nm31gHizV3abKCpCmB4pjUYHuDSodfh9") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1692280404 AUTOM. 3PM
If cOd7 = dcr("ILPU2+SY02Hr9lGdEgIh1nfQ7rhovTNn") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1695237986 AUTOM. 3PM
If cOd7 = dcr("VxzwYo7+kHj9pifQsNERGZmNwVogVyg/") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1701536585 AUTOM. 3PM
If cOd7 = dcr("Lsq+aG27oiW1imDYMpV6TcfPhtxo1bA0") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1703050263 AUTOM. 3PM
If cOd7 = dcr("u9L/uKXPJ3CRwXV6Yue/xzQzGFemk7rh") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1703733632 AUTOM. 3PM
If cOd7 = dcr("8Pg0P0sHHNAzpAdi02j5F1tSLPNg0J2G") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1704494168 AUTOM. 3PM
If cOd7 = dcr("CfvxU9YTUH+vqqjfT7T853T30QrLj4SD") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1706847479 AUTOM. 3PM
If cOd7 = dcr("UELPm3tj1JX164stAkn0Ev/wrYocIFee") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1707530848 AUTOM. 3PM
If cOd7 = dcr("s3QRmO54VNl9jBl/eZg6wkZnkcHNQloy") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1714503492 AUTOM. 3PM
If cOd7 = dcr("pEdLn/7NhY9BLx006F8OncOWdqRIzjvh") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1716017170 AUTOM. 3PM
If cOd7 = dcr("+qBpAeAyj78/5hOisOk6MKUJZxUOGc9m") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1716777705 AUTOM. 3PM
If cOd7 = dcr("TCV93BYNYDPvtLXMFdwbRWP07HnFeSoz") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1717617339 AUTOM. 3PM
If cOd7 = dcr("mVxE3LdJhVbRPBHxqpWX3II6xTbPXiqo") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1721476136 AUTOM. 3PM
If cOd7 = dcr("q+eS/aYCErs2drW0oJFL/tBn6Ca+eERi") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1726873521 AUTOM. 3PM
If cOd7 = dcr("pKkgcgZiO1P7e/RYLb3xdRcYeT6Cys7H") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1729602556 AUTOM. 3PM
If cOd7 = dcr("lzWIj/+M0AkFV4CUUff2ymWDYB89CBzK") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1731483528 AUTOM. 3PM
If cOd7 = dcr("DrohvtY+sXtmap5pO6zTeRZ/COln6voH") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1738088878 AUTOM. 3PM
If cOd7 = dcr("uOfhH9yQhshJsEbToDmFGQrQAG1N/QPb") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1738317425 AUTOM. 3PM
If cOd7 = dcr("xCMLFSlRpddQIz5DhZl6J9McglZTykg0") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1746129703 AUTOM. 3PM
If cOd7 = dcr("DQj4P3wK1OvN8oChD1iT1+4o/wDdqXby") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1757354385 AUTOM. 3PM
If cOd7 = dcr("HAmRp0AB9HB9sVxDklh5c47qwa9Ye80B") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1758114920 AUTOM. 3PM
If cOd7 = dcr("zqok7KcAXr5ENnUctCigkz5BQQxR8kZE") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1774336353 AUTOM. 3PM
If cOd7 = dcr("PuMhQTTj5uRtZEdymoLq/GmE3GL72FJ2") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1778133570 AUTOM. 3PM
If cOd7 = dcr("/Y/mCBQ8cQdRV7JRfNqzeC5ctIXCYQIY") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1779647248 AUTOM. 3PM
If cOd7 = dcr("jpTlyuf1qcMg0JVwnUnEFK0Pf9BMoGgY") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1792839392 AUTOM. 3PM
If cOd7 = dcr("hGXV3qQBqZmfGMfriGK8BsC5MfWa1+O2") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1798150286 AUTOM. 3PM
If cOd7 = dcr("IdJWH/FTj89MIzYqNaBLszMo4FamJn1F") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1799673289 AUTOM. 3PM
If cOd7 = dcr("BbZLSsRYqEIpXq6KXBq0PSs8y67P5swS") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1802314796 AUTOM. 3PM
If cOd7 = dcr("uakbGrUZd58dbiBYFVgCKs+H6bNrKFSG") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1808167003 AUTOM. 3PM
If cOd7 = dcr("gNmUNUhnJK0W1EvEnqvpNlcULeEMwLuJ") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1809680681 AUTOM. 3PM
If cOd7 = dcr("iGaqvE7JgtsDyUHJKxzeKEMTHR+ArO0S") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1819690006 AUTOM. 3PM
If cOd7 = dcr("lZVyHqQGz5CimVdBMBj7zxuHkVtksny1") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1822331513 AUTOM. 3PM
If cOd7 = dcr("QCu9IRve4Pr9LQeMdkbYroNZDdh28AcL") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1822865434 AUTOM. 3PM
If cOd7 = dcr("zRjhZHTFary95irQ0Ljn2KhzYSAiPGZ4") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1823625969 AUTOM. 3PM
If cOd7 = dcr("itTJAhfRR+4bP5MlaDWsD5BTO5Ff1/Gu") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1828936863 AUTOM. 3PM
If cOd7 = dcr("/i4wJfi1Zm1VCH+AqneFQdddTIo8H02x") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1828946188 AUTOM. 3PM
If cOd7 = dcr("tdbPC2+ya5J8+j5wXCH6piF3Ru1rSIhu") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1863668728 AUTOM. 3PM
If cOd7 = dcr("NE4NBlbMQ8f/Kie5wfHQxzZFo2gtkaiX") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1868753007 AUTOM. 3PM
If cOd7 = dcr("ixyrqlk/bNRGNbejHIjNpu3aBee0SE+4") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1872917517 AUTOM. 3PM
If cOd7 = dcr("lYvSp6kkZiZFZ7Yf6jI6ZyZrSryVNdtA") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1875954198 AUTOM. 3PM
If cOd7 = dcr("zOjBP1O0EOg5SPodDFZzPrxTMO9OcYdg") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1882164374 AUTOM. 3PM
If cOd7 = dcr("RXbvg5y4M+M5WGqBX1Gf0XOMjDH5N5X8") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1883687377 AUTOM. 3PM
If cOd7 = dcr("eB5qZlF3R6pF1H8ytooNGpZaM9y9iMei") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1885961590 AUTOM. 3PM
If cOd7 = dcr("q2QyDdpmZ+q2FRp/bwPgcuanGeP7LHzt") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1891639778 AUTOM. 3PM
If cOd7 = dcr("R6389oGUA1hJ4kGbYgnINNtnNm53lzdl") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1894457237 AUTOM. 3PM
If cOd7 = dcr("gWcKhGvbRfmNI6Xe77Sr6gPQsb+CoZVk") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1895210380 AUTOM. 3PM
If cOd7 = dcr("1il+og0BjY07LAe8+3vHfiY0b9qqLMTp") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1905227097 AUTOM. 3PM
If cOd7 = dcr("YlmtV0XQyv1gjY1AVDiz95Yyye60oirT") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1905980240 AUTOM. 3PM
If cOd7 = dcr("7ljn++Frq2WiZcZ/RXoWQ41EWuLaycBp") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1908795766 AUTOM. 3PM
If cOd7 = dcr("l71Yt4dC58K0d3VMqIt50okllZRDMrq9") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1915236422 AUTOM. 3PM
If cOd7 = dcr("cyy9IVEtQzEyxbrsWl8KGePkwSCpk6FI") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1917510635 AUTOM. 3PM
If cOd7 = dcr("rEbVl/aMHHrT9X1XrapLRWaiT96yMJk/") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1918044556 AUTOM. 3PM
If cOd7 = dcr("EZ5fgT2A1djEd0ozkW1fBAMz76CINyu9") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1918271170 AUTOM. 3PM
If cOd7 = dcr("SDHlj14xYuqREZFq2P1H6DiFoiTnha+r") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1928814416 AUTOM. 3PM
If cOd7 = dcr("p+SRNkLZxmHX/aZ/83dFT2GBXVIshYKd") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1934492603 AUTOM. 3PM
If cOd7 = dcr("sTbyZuOLqxhK2FBdvVVX5+R9f21IjO1C") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1939584275 AUTOM. 3PM
If cOd7 = dcr("Guvn7/+SzzUJKXOgV5VaGXrhD89AnXbe") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1941465247 AUTOM. 3PM
If cOd7 = dcr("q4K3U6B44emSjSVDbd280szQmbJa0sqX") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1942986317 AUTOM. 3PM
If cOd7 = dcr("fYAq8N4ls9znHIJskkIBqw8PVCFH9L8B") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1946783533 AUTOM. 3PM
If cOd7 = dcr("CLX/2QfeAMQGr0qmOmDmCcHOzeoU36fn") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1955050633 AUTOM. 3PM
If cOd7 = dcr("WxnljYfZ+4aWpRpXgtsa/6phXJnAPQGO") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1956564311 AUTOM. 3PM
If cOd7 = dcr("0xPKMaHiwBgzd1kepXXtKyWQAnnt3Esx") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1957326779 AUTOM. 3PM
If cOd7 = dcr("c/U2kMy7415BkyxT7XwCnx/Wb/aETo7c") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1963003034 AUTOM. 3PM
If cOd7 = dcr("WAv/pn9zDpjhbVmngZd45Zv/FEfFpONv") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1965820493 AUTOM. 3PM
If cOd7 = dcr("L5ufrqnsKU6E1Kr3r8qWMezuqJLSwxkR") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1976590353 AUTOM. 3PM
If cOd7 = dcr("fRSX4rT7vB7nx9Kz7S+yhErGaiwgLpX7") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1978866499 AUTOM. 3PM
If cOd7 = dcr("UUphEpDuE8+5MSBKg1WGu9FmRPXzAWor") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1985076675 AUTOM. 3PM
If cOd7 = dcr("IWyAAq05R62itedahEgoPvsMRd1mo9Yv") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1985839143 AUTOM. 3PM
If cOd7 = dcr("9wtwE3ATf2/+tboQ4Pc6w1ioh/2eh8kG") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'1995846535 AUTOM. 3PM
If cOd7 = dcr("arHz1WCZe+nndQaxa8nOMEww4gLcqbeH") And LCase(Sf7) = LCase(dcr("4LTC29s72z02y+fnCryAXQ==")) Then LaLi = Supsabseee

'2005855860 AUTOM. 3PM
If cOd7 = dcr("MB5hpUlOWp4GcRsPQCZtmy8VVvTORjA1") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2012294583 AUTOM. 3PM
If cOd7 = dcr("VXx7lUy1XGIS5zdF3i0Hx8kad1q4s663") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2015112041 AUTOM. 3PM
If cOd7 = dcr("/N/lHTPVlmqvUSphOspYDZ1DgbnFt26s") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2026635044 AUTOM. 3PM
If cOd7 = dcr("Jo6kbWhgIsYe+akfSE4ahnH6PCD5PXZn") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2028916649 AUTOM. 3PM
If cOd7 = dcr("pgeGxYc3qYf1RkUXN/KNVK/7viNMrQUA") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2032311299 AUTOM. 3PM
If cOd7 = dcr("fIvTvPIt3f9qzgqPULNei/pZX2OZEGuh") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2032320624 AUTOM. 3PM
If cOd7 = dcr("jv1fe76+ecboxeDp569dYVVKWnQAcDvT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2043081159 AUTOM. 3PM
If cOd7 = dcr("hhZhi80YLag1WvvvjznPQrmSfZ/9iAI8") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2045898618 AUTOM. 3PM
If cOd7 = dcr("JGy1BjBrl5t6V9A8X64ZVnY7JBIKXYjr") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2052031628 AUTOM. 3PM
If cOd7 = dcr("HdDyYvp2SXqgNhuCLJcNJahPMlH1U81Q") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2053624405 AUTOM. 3PM
If cOd7 = dcr("glI4Te0++JkC0fXzRyl+3zCkyj14Uu5K") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2057429013 AUTOM. 3PM
If cOd7 = dcr("wa+HncGmaC7pGLmBjFS1GPu4jyy8QqMc") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2059476612 AUTOM. 3PM
If cOd7 = dcr("RxydIY1cqr87yEKm9FHP3EVS2SvBfIh2") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2063641122 AUTOM. 3PM
If cOd7 = dcr("EECAoNN8iqT6XmI751njekWqTRSrWx/R") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2069712551 AUTOM. 3PM
If cOd7 = dcr("xBwMAdvXF0I6KTm1XtKSOXAjSdaKemis") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2074410982 AUTOM. 3PM
If cOd7 = dcr("fqh8cCfKerDBacKOA+yb6sYFyPM+mjp0") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2078968733 AUTOM. 3PM
If cOd7 = dcr("0b+I9yWhPxnt1gOGfwpIdCy0kzYz27ls") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2079416163 AUTOM. 3PM
If cOd7 = dcr("Tj/RH3JA+eUp1E0RqcDGRfIywFPWkGKR") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2079502654 AUTOM. 3PM
If cOd7 = dcr("DRO3z/hH4bzoEnwYFvFU0byBOsfqhZ7g") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2079721876 AUTOM. 3PM
If cOd7 = dcr("8iuLn3D3I2o2qhmBbgQgLVINXDGjSE+e") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2085180842 AUTOM. 3PM
If cOd7 = dcr("P5353px9asDoT/oSGovJy+Q8T07ztLTf") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2085714762 AUTOM. 3PM
If cOd7 = dcr("5XQkCvTA0jVyflb9041mLB86eoXsbnqa") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2087988976 AUTOM. 3PM
If cOd7 = dcr("GvlHy+a47R/vYzN9v2BnYxelLCWlrYkz") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2091392950 AUTOM. 3PM
If cOd7 = dcr("FZZu395TxVLZ1iuWxdAWiB0Njl/E1HPS") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2096711236 AUTOM. 3PM
If cOd7 = dcr("sCRWmPm7y/X5/d2NPF7HMwuSsTFwLFdc") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2102836854 AUTOM. 3PM
If cOd7 = dcr("1Zebr5YIV0P4WtH8u8Jo1qwxJRbu/1G5") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2107928526 AUTOM. 3PM
If cOd7 = dcr("KUO900rbMBgPdVxHT8XukmdNhD0azjrX") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2108015017 AUTOM. 3PM
If cOd7 = dcr("DmScNU6yaIH837HDx4CUs+8y/vi7W8WO") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2115740804 AUTOM. 3PM
If cOd7 = dcr("l+36CtnCQgUYhmOLNHggc29f738B14q7") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2124463065 AUTOM. 3PM
If cOd7 = dcr("0x4bDg3YTjfo+8/j6Zx6YX5N41Y380cZ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2126431565 AUTOM. 3PM
If cOd7 = dcr("U+6wrrRqa71tGPhssTcyLgqvvXYbcuzh") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2126737278 AUTOM. 3PM
If cOd7 = dcr("2/LfRRNT69uDhismfVVt37k7y3BXQQqj") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2127271199 AUTOM. 3PM
If cOd7 = dcr("H5uw6hrJj/v6qyiSYF3zU94VAkm/zgDu") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2131349218 AUTOM. 3PM
If cOd7 = dcr("E9r3RNMyLZ12fhi342jUFw+3Y7/2ycTc") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2135232925 AUTOM. 3PM
If cOd7 = dcr("jvlThIIcaC03LO09WI0Y44ZlhactXcD5") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2136746603 AUTOM. 3PM
If cOd7 = dcr("WNuIcPsNPdYnK8hUC3wX3G/uLv31VFlr") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2141445034 AUTOM. 3PM
If cOd7 = dcr("zNT6YSzedCeduGUKXc2GSlVfZ2y19lqa") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2145993460 AUTOM. 3PM
If cOd7 = dcr("RVWdg65mjMxr2vjWD6puItJCe53iO/Ah") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2146002785 AUTOM. 3PM
If cOd7 = dcr("/jdHxrXgObC2NLLm/FWEzubSEn7ZsdDT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2156457607 AUTOM. 3PM
If cOd7 = dcr("TbYDU/kwiLx+RcPa7QTXnX5ZBjtLoDjd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2158820244 AUTOM. 3PM
If cOd7 = dcr("Onr6BxDMGQZ2FMeFxKL58Lgdfl+PcUip") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2162669716 AUTOM. 3PM
If cOd7 = dcr("lktnq/Jf2GEogh+jQ8k2++H8snLO5mQT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2182695757 AUTOM. 3PM
If cOd7 = dcr("UOWG8sOOIP3OdyPh4J/qN+8UubqRa8Oj") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2183456292 AUTOM. 3PM
If cOd7 = dcr("pyaMSVVMZ0/jEjMWlahWrNMQbRQTcWJR") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2186492973 AUTOM. 3PM
If cOd7 = dcr("7HItBRedRhwDGsGmTc3aXI3boE1Biq2h") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2192782248 AUTOM. 3PM
If cOd7 = dcr("tZfHkwB+lym/H30dCKHC9H/ILCz3tLdd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2203552108 AUTOM. 3PM
If cOd7 = dcr("L0s7amBTggrB6/2jnWz6wfpXSAe74OmI") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2209546371 AUTOM. 3PM
If cOd7 = dcr("333wESxCT+rMSt7+qC0FVf/xPzymFW2+") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2212808290 AUTOM. 3PM
If cOd7 = dcr("TpN6qmsIrP/y6Ye9X++IfSqVziHagidH") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2213482334 AUTOM. 3PM
If cOd7 = dcr("vp1POjw1VS7WYgQwjHj3Z9U87cUPsqD6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2218802553 AUTOM. 3PM
If cOd7 = dcr("7d2YcHPtapymM3lOy0yzMeUxRsaZwR9b") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2221968656 AUTOM. 3PM
If cOd7 = dcr("eWNXDOEWNSxfG7pCu+77OSLrcJ2cZW4A") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2225091828 AUTOM. 3PM
If cOd7 = dcr("KSrtwQ5bPIbbd1UWIMX7A6MRg5WTRJKK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2229563088 AUTOM. 3PM
If cOd7 = dcr("JqViN8sAKNmmufqq7EHeIzd89kvF5AQK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2229572413 AUTOM. 3PM
If cOd7 = dcr("zk7JX9jDmipMaVGsKHYwfHQgEDdhqVGU") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2231444060 AUTOM. 3PM
If cOd7 = dcr("OuBbDdrLw+G0MQ+Kkg0FbXEu3FNmbpBQ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2242747841 AUTOM. 3PM
If cOd7 = dcr("LEcSy/W92TqWbZVz+uYcutmEmtG2UX+p") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2247839513 AUTOM. 3PM
If cOd7 = dcr("U3yKYew21Akd9N8QGoZoXG56cJd6rHW6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2253517700 AUTOM. 3PM
If cOd7 = dcr("WCnEaTMIqiGu+qhZbbgz1OCFM0MQmq9+") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2256561773 AUTOM. 3PM
If cOd7 = dcr("+OIcuDM6S3sSHnBtP/Ui22XJTvkhfJ1U") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2267182185 AUTOM. 3PM
If cOd7 = dcr("s1vpkryEN2iAJTp7cHsrTzWLXIYBCyGP") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2276351876 AUTOM. 3PM
If cOd7 = dcr("2JNgF96XpgZx7Tk1J/1tkcUUXuy6d2cJ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2285608058 AUTOM. 3PM
If cOd7 = dcr("jvaid83pCKzuqAno9OtNxaxImUaMWBx+") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2287882271 AUTOM. 3PM
If cOd7 = dcr("bl8qoQnrkzIdO70/5PZANakKDdC5SVVl") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2296377918 AUTOM. 3PM
If cOd7 = dcr("VUhbYfJwHV5BqabPePVBOjeo5wcVn+Yi") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2299412666 AUTOM. 3PM
If cOd7 = dcr("cIdVhWel9mjknt6l0fcxbybSSRZlxQL+") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2299491765 AUTOM. 3PM
If cOd7 = dcr("Zq2jzCAl0HBUrICdszebB0uSURWIriEE") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2301293638 AUTOM. 3PM
If cOd7 = dcr("6025ixyTp0RlvheeXprUxvcmg2iNzehw") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2301372737 AUTOM. 3PM
If cOd7 = dcr("hxPtQC7UEGjFL4Zi6YsnDCg7GDZqLnqP") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2302816641 AUTOM. 3PM
If cOd7 = dcr("CezTfdyKhbWBDmD9SPYBXAht4vY+JnCh") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2303569784 AUTOM. 3PM
If cOd7 = dcr("JfJTzs1qnWERRX8o7/0GTOEmI7HneuxV") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2304864240 AUTOM. 3PM
If cOd7 = dcr("e6Xzkw5PYNUG2OSnheDbDl4hEYKAM10b") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2305169953 AUTOM. 3PM
If cOd7 = dcr("ujkLIrKEubhsurtF5dPHqxzztKMeJICN") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2312142597 AUTOM. 3PM
If cOd7 = dcr("fWGE78qtJvzbp3bZKNfGgG8EFcjpOCnA") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2315634100 AUTOM. 3PM
If cOd7 = dcr("K5h5DgaF00UTeBA2FL9rVuKHGL/N6aXB") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2317155170 AUTOM. 3PM
If cOd7 = dcr("SXhXNHrPP+S23+KhnugV8tnHIrlSBS4L") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2322833358 AUTOM. 3PM
If cOd7 = dcr("pZvSFK4NhvDgft/u7LB8wWkHCdOCsNj6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2329438708 AUTOM. 3PM
If cOd7 = dcr("R30kWF1rszSFVu9NJemoNiQcU4Csz1RC") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2339448032 AUTOM. 3PM
If cOd7 = dcr("oXvK36mELUJNbumDYvbvmDNxZ86yxy+s") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2342852007 AUTOM. 3PM
If cOd7 = dcr("R/nWhD3by1XGiIUd6Ab26vIMDlWGCULW") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2347181211 AUTOM. 3PM
If cOd7 = dcr("N6D6WAruB3dvMgxaDOSpNqxV99ercAli") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2353708358 AUTOM. 3PM
If cOd7 = dcr("TmTnCcSoJMEGYw4/23Bh+L6fN9Hh1Ta7") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2363402645 AUTOM. 3PM
If cOd7 = dcr("4DVtGpEqzzGeLIafi9OIXgZbaPu+b2mx") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2367199861 AUTOM. 3PM
If cOd7 = dcr("JnCe7P9U+GJ2wLwKc0H0rYuKsEXT2YLT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2369483399 AUTOM. 3PM
If cOd7 = dcr("juTV6zm0jYsv/Lp15MMgQvmfnpWN+LQj") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2372964540 AUTOM. 3PM
If cOd7 = dcr("fLm+bvbgBuvBVi1HaKTyQy9GLJbj3Jw8") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2377216578 AUTOM. 3PM
If cOd7 = dcr("nTOTUzEv3cacMtYW6nY80X6IF8WB1ES9") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2388520359 AUTOM. 3PM
If cOd7 = dcr("svY4NqyVeAP/ky2Ck/6myDlttDWijeFK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2397233295 AUTOM. 3PM
If cOd7 = dcr("z8M/88tpQOfC+aNBJ5aplv2rMAt7uKtc") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2404205939 AUTOM. 3PM
If cOd7 = dcr("ZDj0pmVkeeCBXpwIXUcZuLWta2LCFfTS") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2408003155 AUTOM. 3PM
If cOd7 = dcr("UuiHYuHe4XUpLLtdyHker1KxfPJWyeFz") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2418546400 AUTOM. 3PM
If cOd7 = dcr("KBwIth9m9pMpQuA6L9SLXWjpAPAX3KOe") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2421948442 AUTOM. 3PM
If cOd7 = dcr("tkAT9Al4aDLJj/nnh/t92U9i1EqaOtdC") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2428782339 AUTOM. 3PM
If cOd7 = dcr("lJNcz/om4Yon2lWgXnlzXUT/LnHKkDj7") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2441204624 AUTOM. 3PM
If cOd7 = dcr("F/IoIjQKukYYyCCmwti0GuutGjP68X2F") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2442727627 AUTOM. 3PM
If cOd7 = dcr("SCTajj6p+mexb3Pb+57SBxlIoUIMw1LS") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2447819299 AUTOM. 3PM
If cOd7 = dcr("HITeqPyQnHsgmjDCsVsnImnVt1RjWBgN") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2448799056 AUTOM. 3PM
If cOd7 = dcr("tFqJX84uD9tOb/0xOpaEAJgQ5iVbAQdk") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2449332977 AUTOM. 3PM
If cOd7 = dcr("AmHBJn+kIuBoSkrOiEirqfBgOBE5rdBi") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2451974484 AUTOM. 3PM
If cOd7 = dcr("lRqcpBi7q1n+OYaS4ZRXmAZDuJMetWz6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2454250630 AUTOM. 3PM
If cOd7 = dcr("LGhlQTA4upu1dH5ORjINMHUPrIlNxpH0") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2455011165 AUTOM. 3PM
If cOd7 = dcr("YIgRzQf/I8D1qlrOPABdwQEr7X6jX4cK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2457294703 AUTOM. 3PM
If cOd7 = dcr("1h/bjteDFoq6uniijKRq+fsTcYxh46Tg") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2458055238 AUTOM. 3PM
If cOd7 = dcr("CNjISXK0nh3mXxPJzWpHQPz7rNCbWPfr") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2458808381 AUTOM. 3PM
If cOd7 = dcr("tERiRVDRaSW2gwwiPonX+NwVeyebWs40") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2461230666 AUTOM. 3PM
If cOd7 = dcr("nS4JHgYC9MlyAwhUMinKPs0yAVQaV9CW") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2461449888 AUTOM. 3PM
If cOd7 = dcr("lBZ+RGCxCk6w7A38L0gEc6+4S11G8J5k") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2463504879 AUTOM. 3PM
If cOd7 = dcr("xA7E4PoXtLw3ctBt2Qh6hdpJdpmCR052") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2477845341 AUTOM. 3PM
If cOd7 = dcr("EBlamXX5w+nnPNFzHMdh511jPOSg8QpY") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2479594958 AUTOM. 3PM
If cOd7 = dcr("e8gf5vbONQLOCx9h6kGgc1uEI4iGM5T/") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2481249315 AUTOM. 3PM
If cOd7 = dcr("liuCqDBujtjT9KwuJ74R/G1Ai9f3I4Ki") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2482770385 AUTOM. 3PM
If cOd7 = dcr("KsavAJJUW9oc86OY+n7g3Ityh2DuMuiT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2483523528 AUTOM. 3PM
If cOd7 = dcr("OIb1/nFbHMEh+w4DJ6ZNYYi0mW7xrT3T") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2498622592 AUTOM. 3PM
If cOd7 = dcr("zr5zwFwK+4k5kXt3gb+MieTmd3d+1VDR") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2505063248 AUTOM. 3PM
If cOd7 = dcr("oR3Y+Iy2y+d1lq6yg2T+pue5E0xqLhYo") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2513558895 AUTOM. 3PM
If cOd7 = dcr("1BM+ifpO8ER4cc+4zQMHSbe3DKiow7W7") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2518648634 AUTOM. 3PM
If cOd7 = dcr("MGmI47MnbeGacdlWANufL98yP/pOvfpb") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2518650567 AUTOM. 3PM
If cOd7 = dcr("Phl1b9wzDefXM6I1tjWiw3bbsNH37f26") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2522045217 AUTOM. 3PM
If cOd7 = dcr("w9DSt77nfoJMAhNKnDMv6fD33C4YUl+9") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2527136889 AUTOM. 3PM
If cOd7 = dcr("0+v6sBWb2EYHr36dr/UoL+Mf4svNBHG6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2531301399 AUTOM. 3PM
If cOd7 = dcr("jRiBtKZ17IHOJcQdW7nnIFjfVD4r87Lg") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2538667284 AUTOM. 3PM
If cOd7 = dcr("88wEg1SJxEU6M3vzTv7/++5mr/1XsKDd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2544116925 AUTOM. 3PM
If cOd7 = dcr("GbhXhsixjI64mkGcVGMsWAGJ9ylaeNtd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2554126250 AUTOM. 3PM
If cOd7 = dcr("Mz9dYnUfI1ClBu+Vmkxk7P9OjRPvolCP") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2554352864 AUTOM. 3PM
If cOd7 = dcr("nw1Fa1sGVEaXoRbOE+9pRrRRSGEd+6Fu") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2558684001 AUTOM. 3PM
If cOd7 = dcr("DhzhpUZbek/Gs1d9sSgbVxXqsl1Jw36h") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2562542797 AUTOM. 3PM
If cOd7 = dcr("ck0o8Fc51AfhgToyILnxKE8i43sX5PVe") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2564142966 AUTOM. 3PM
If cOd7 = dcr("iF2nQnALgN00/zVPvtxtVn3+HP1/MWRQ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2568700717 AUTOM. 3PM
If cOd7 = dcr("SA7Myt5qnpsBwHt1sclXrYvKG6JY2IoX") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2581885470 AUTOM. 3PM
If cOd7 = dcr("Emg6J0rWv4todsQRBXpiZYJnsUbK9jXk") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2582104692 AUTOM. 3PM
If cOd7 = dcr("QO0oNuUWWokC3PbWP5dUOQqLvLj67mvM") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2592121409 AUTOM. 3PM
If cOd7 = dcr("tVGpk6QHeMtARIqEW0CAJ9VF17+f15aW") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2592568839 AUTOM. 3PM
If cOd7 = dcr("1AKSN6GoNulhYOUvfwAQmQq0ghcDj9He") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2598500145 AUTOM. 3PM
If cOd7 = dcr("BhIoh7KKyUiyb3kxtZfPRKPO8Nksi436") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2602130734 AUTOM. 3PM
If cOd7 = dcr("nS4aaEZU27JklMmwNurGkHDssgSD7sr/") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2607756327 AUTOM. 3PM
If cOd7 = dcr("axE06PwvCcYQn4L+i0Slfd8HkvLjCfh1") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2607975549 AUTOM. 3PM
If cOd7 = dcr("asN34Ctz/ysOFIvd9cjefwHNUoMHruS+") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2611911512 AUTOM. 3PM
If cOd7 = dcr("gKZzYQu+CxvsmtnZIFeXIwkPlsXnZCHw") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2618526187 AUTOM. 3PM
If cOd7 = dcr("6mQusIazuXSATJ0U+pZrlJrONOPXXD2Q") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2619200231 AUTOM. 3PM
If cOd7 = dcr("OAz3dkj0HSSCI3ww3rigXlpdqpNnx/ly") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2623661129 AUTOM. 3PM
If cOd7 = dcr("yVCFDF1Wb7OFCxgtmVRRVYzkLC/m2pYJ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2627012509 AUTOM. 3PM
If cOd7 = dcr("LCkItZz0tJKyTMC1+TFTNjzSX9aCDn/h") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2627239123 AUTOM. 3PM
If cOd7 = dcr("kQjqdvTotsQoKT6sCqufJ58LSLJLABqF") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2633451231 AUTOM. 3PM
If cOd7 = dcr("1clW90qZBfxOUV+FWrZv4GB65V2BCyNh") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2638456413 AUTOM. 3PM
If cOd7 = dcr("hUhbFN19ebdAx1i24WjdlB2P7FDPtov7") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2638762126 AUTOM. 3PM
If cOd7 = dcr("jjCUEtxbEa6415Q9HngoFGCVQ+hzsOuT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2641177018 AUTOM. 3PM
If cOd7 = dcr("imk5Q3xgzK1O8gAUKsM20TSrwvkaIqJQ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2646495305 AUTOM. 3PM
If cOd7 = dcr("KIy5GKV/zu/thGoosk3ukF3DpZfXfIpp") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2649226273 AUTOM. 3PM
If cOd7 = dcr("kjkNBJVqkavUiKnAE5KaZAg03cFAmzP6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2657038550 AUTOM. 3PM
If cOd7 = dcr("cK5x+SsFMHn42DC6CbtrdcTeWjs12gEf") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2657799085 AUTOM. 3PM
If cOd7 = dcr("FlrXPE5p3WPBCMsXZe5sLcEGWUDtt6sU") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2666968776 AUTOM. 3PM
If cOd7 = dcr("uomPRxVirnQPzM8yWMe3estsrpRSeSo4") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2667055267 AUTOM. 3PM
If cOd7 = dcr("ER67wimfcGEWlUa/h4xv7NEJyyy3rLQu") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2674247133 AUTOM. 3PM
If cOd7 = dcr("FJ7/YuhhNBG8zJSj1Eg42S7mlsdO+eZ4") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2675770136 AUTOM. 3PM
If cOd7 = dcr("ZLzXjMAPNyjy2lvSVeJ/KG3QXYqo21tq") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2678578270 AUTOM. 3PM
If cOd7 = dcr("johJpCXlFmIibQ/tZcb6px4x3sen1B+3") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2678804884 AUTOM. 3PM
If cOd7 = dcr("p3hC3HH3ub2hD0CcploqfIYp89228A6i") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2686994818 AUTOM. 3PM
If cOd7 = dcr("MhpGZvsEbdQAUuxGaIGVndPlsx60RhVL") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2701919863 AUTOM. 3PM
If cOd7 = dcr("RKARSZ7/1N3q7D0fco+HTHXgKlYWSrYz") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2712768821 AUTOM. 3PM
If cOd7 = dcr("DZONmsrlJZxQixV+j1k1tGDd0B+8tznk") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2725593672 AUTOM. 3PM
If cOd7 = dcr("WBQmDb4nnPxS8/KTeC45T86vnkBfXcg1") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2727790719 AUTOM. 3PM
If cOd7 = dcr("ouVH1FTVuPVh/IptDdwO2pGN0JDauWPp") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2748647070 AUTOM. 3PM
If cOd7 = dcr("54QyUd8kT0O86IL2AIuPcRljLmHgg6Tm") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2756303083 AUTOM. 3PM
If cOd7 = dcr("MVvf4Shg5LzcdmYo04ZmGevasUYJSoO3") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2761297902 AUTOM. 3PM
If cOd7 = dcr("ybbjibOZwFOSCauFVQyHKIzjbTLImpXp") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2767833478 AUTOM. 3PM
If cOd7 = dcr("wUkChqTwVijavk0ziWmVMVgNdBB7ZsaR") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2776319800 AUTOM. 3PM
If cOd7 = dcr("rrieYgFkbUIfYzjAGmKI+gYCc7bL7+AZ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2777842803 AUTOM. 3PM
If cOd7 = dcr("JNBl2BuMEdJDdA5SAlj4qa6Lu7AtNHf/") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2789223750 AUTOM. 3PM
If cOd7 = dcr("qTBxJQD3sTYF8oUoOChOus4ITLd2zdVy") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2789373198 AUTOM. 3PM
If cOd7 = dcr("GYNRZ3DJ1c6R18CWTNpUZQi22tjm3JFq") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2791778766 AUTOM. 3PM
If cOd7 = dcr("tYyU6YSxt+6Df3KqGsR7rUM/L7tlCMyw") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2798479931 AUTOM. 3PM
If cOd7 = dcr("hnN1dAU7StIvn5HM5MRJBPftSt5Mshex") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2799382522 AUTOM. 3PM
If cOd7 = dcr("6EQNMjcvFh8UAY73dM9i3tuUGpvRxkkX") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2802548625 AUTOM. 3PM
If cOd7 = dcr("XbtuTY8HM6NlcSV8SyZouwFml3Es48Z1") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2804071628 AUTOM. 3PM
If cOd7 = dcr("SBgwTbrGCRXLs+Fqe6IxyGLMi4+TB4fH") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2809249791 AUTOM. 3PM
If cOd7 = dcr("4nufGI4v7HPS0WTFpT9E7CdgGSIQXCOO") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2816441657 AUTOM. 3PM
If cOd7 = dcr("3o7MfCzbva5uvRf1ZGJ/x1+J0SNF5YGm") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2828419482 AUTOM. 3PM
If cOd7 = dcr("2CEoAcZx/3q0HPVgHtIuK+giIxrLbnke") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2828505973 AUTOM. 3PM
If cOd7 = dcr("FsbySMUI3Xu9mGYXB3tWTVuh0p/dtLnT") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2829180017 AUTOM. 3PM
If cOd7 = dcr("4BUjkSK6SvOujj5dl5zB700VitlVyjZu") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2836381208 AUTOM. 3PM
If cOd7 = dcr("nfsU1Jc/XwQTdNZtDEb01D86m3oci4uv") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2838655421 AUTOM. 3PM
If cOd7 = dcr("4xUfWO6Qssqj3dXH/c1H2aSwuVj7QH6a") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2843353852 AUTOM. 3PM
If cOd7 = dcr("OzyLSGPrDrivyP0CaUC1dd1eJymYLIXq") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2844954021 AUTOM. 3PM
If cOd7 = dcr("HXYdK0ZtdC9AJSORn+2n/MSgwfg1wpWx") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2847151068 AUTOM. 3PM
If cOd7 = dcr("TgqitPv/SfKxc3nFCP06yn04+QjIfpJd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2847762155 AUTOM. 3PM
If cOd7 = dcr("lkkkcIvBvVvOcIiHNbxvpD5i/7VbsTy1") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2849206059 AUTOM. 3PM
If cOd7 = dcr("F7yXrZsLqbwoqCtlayiAam4YHvD6OJIS") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2851840174 AUTOM. 3PM
If cOd7 = dcr("dpjZqyH4tqNiAphhpOcR6p0HAbjsW4kK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2861096356 AUTOM. 3PM
If cOd7 = dcr("QmSRUJUOk8SZtBRevhVu8MRiGzDrC0z4") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2864893572 AUTOM. 3PM
If cOd7 = dcr("20eX6UNFj3gtWBizwVTtOO+Q0tKFi2ng") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2869301874 AUTOM. 3PM
If cOd7 = dcr("8gynqJNXKaM7O9X0JEuEi7i/ijr4vpt9") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2876423967 AUTOM. 3PM
If cOd7 = dcr("MNfsEbMXO8O6fF1WlXB8+A5YvEh0xFIk") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2877718422 AUTOM. 3PM
If cOd7 = dcr("MeyCsCqlX7AyF0kSacL3/Cl+VXEQtmQB") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2883930531 AUTOM. 3PM
If cOd7 = dcr("tSWfOaEWGPez79KGw7tq39x43jFBQjEQ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2885749922 AUTOM. 3PM
If cOd7 = dcr("oAHYgswPECqHSGHnAkkZTR3W14PHys1S") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2891208888 AUTOM. 3PM
If cOd7 = dcr("7mD8tZWTPiWLRmuxo45qP5M12Mt8NApZ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2912748608 AUTOM. 3PM
If cOd7 = dcr("6iFy+QC6aDdhRz8neEKJmteeIyZTVJ6s") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2914717108 AUTOM. 3PM
If cOd7 = dcr("z8fuIMnyY7wxXg5+SRKKPJ7gFihcGEAB") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2915477643 AUTOM. 3PM
If cOd7 = dcr("ToyjlL77yJy1ZaNTx54XHvCb2FMdaC0J") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2923203430 AUTOM. 3PM
If cOd7 = dcr("bN1oZlfo+R3F9cSNE3KVx82/BWPHF7g5") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2932678834 AUTOM. 3PM
If cOd7 = dcr("ZMoZXo3UfABvP7GZIsTI7qvmdvttYstv") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2937009971 AUTOM. 3PM
If cOd7 = dcr("PcTBzqJytFljlR0DIBBRhenE3p8cwcB6") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2942695551 AUTOM. 3PM
If cOd7 = dcr("ISnSBCy9kbWc4YCB6rigBn2fWKc4Uemo") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2943229472 AUTOM. 3PM
If cOd7 = dcr("7UbMg8Wv48bUT1Mn8xAmETLMBOeueUrZ") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2944743150 AUTOM. 3PM
If cOd7 = dcr("33aA9j9F/cgkn5xSVNUpo8qzG8Ac5wSt") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2945503685 AUTOM. 3PM
If cOd7 = dcr("0V6WKDuw9LbRa4ksBlPN/az0gQk9xvpo") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2947026688 AUTOM. 3PM
If cOd7 = dcr("B+ZmGu3xsabPqCtire9RlTr1Qk6t5Tti") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2951715793 AUTOM. 3PM
If cOd7 = dcr("Uj8nda580DW1+BjMIL43r2E0+znHSscw") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2953465411 AUTOM. 3PM
If cOd7 = dcr("uKjq1pghsTR5jTCWXCNKsdhWya0K77a5") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2954752474 AUTOM. 3PM
If cOd7 = dcr("ai9DukTQzrsym/eZh8vkg8fnRPinrRjh") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2958557082 AUTOM. 3PM
If cOd7 = dcr("QsKxSnRp7GsQa3vYEh6SUVboIv4/JBjC") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2959844146 AUTOM. 3PM
If cOd7 = dcr("Dp2s8AJZzUeX3+EWmpa15slwOZ7mxkFd") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2967043404 AUTOM. 3PM
If cOd7 = dcr("vpkcNQwpDWSdL50Lv3boBTAhrE9ThJed") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2968566407 AUTOM. 3PM
If cOd7 = dcr("KKWkneRTTBoUzNOiYC2yj/b+dvvv/cxq") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2971741835 AUTOM. 3PM
If cOd7 = dcr("TxdmxEKY+05HUp16/twtqZUQgTxFNgi4") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2975539051 AUTOM. 3PM
If cOd7 = dcr("wmIMS28LPKDxQghXFuCRDJ15v/GHKe8B") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2983264838 AUTOM. 3PM
If cOd7 = dcr("oiyn1NfCojJf6hZZfqnxDQUB3GYYowJK") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2983491452 AUTOM. 3PM
If cOd7 = dcr("YaNhvgQwNLDlezcidyE/X+XB3fLad8Gc") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2985548376 AUTOM. 3PM
If cOd7 = dcr("DTONf5ZEcvr9bkn7rzZAYrwuBSCe7ZNC") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee

'2994795233 AUTOM. 3PM
If cOd7 = dcr("Qac3Jd2xNHRF7ZxAfnZ5pn53XiGCgrS4") And LCase(Sf7) = LCase(dcr("sbKfvrSD72Kk1uzZOsl45Q==")) Then LaLi = Supsabseee


End Sub
