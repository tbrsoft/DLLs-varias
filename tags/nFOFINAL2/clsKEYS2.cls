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
'el primer indice es el tipo de licenciay cada uno tiene hasta 5 codigos
'si alguno es cero no se usa

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
                If mNR = lista3(i1, i2) Then
                    TmpLic = i1
                    Encontro = True
                    Exit For
                End If
            Next i2
        Next i1
            
        If Encontro = False Then
            TmpLic = BErronea
            If EsErrorSiNoVa Then Terr.AppendLog "IC10"
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
    'seguir aqui que no haya duplicados !!!
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

'matar la licencia por que se mando una cagada
Public Sub UpMng()
    LaLi = BErronea
End Sub

'setear la licencia segun interfase y soft
'por ejemplo si tiene una interfase valida para el soft!
'lo pongo todo junto para que haya un solo lugar donde manejar todo el kilombo de codigos
Public Sub Set2(Sf7 As String, cOd7 As String)
    'codigos de los programas encriptados
    'genesis-mp6 = 35yfsAKVETNRy2KO99AfqmrB3YO5nwko
    'e2games = no implementado por que no usa interfases
    '3pm = todavia no traje los codigos para aca, acerlo cuando este listo 3PM 8
    
    If cOd7 = 0 Then Exit Sub
    '104329645
    If cOd7 = dcr("RiZ1bEFVeLNS4E7m0dWMwXSCOhw0D/YH") And LCase(Sf7) = dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko") Then LaLi = Supsabseee '1/10 dic 08 martino
    '103389501
    If cOd7 = dcr("3Oiqmf6Xkir9bSvHfj18o50OPvSwwDEL") And LCase(Sf7) = dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko") Then LaLi = Supsabseee '2/10 dic 08 martino
    '107510805
    If cOd7 = dcr("afmxbwcoczuTN6Mdf2O31J0wNquUo8E7") And LCase(Sf7) = dcr("35yfsAKVETNRy2KO99AfqmrB3YO5nwko") Then LaLi = Supsabseee '3/10 dic 08 martino
    
    'este es como un case else para todas las que no anote en 3pm
    'If cod > 0 And Sf7 = "3pm" Then LaLi = Supsabseee
            

End Sub
