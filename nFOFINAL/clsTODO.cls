VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ACUMTTT As String
Dim FSO As New Scripting.FileSystemObject
Dim ResuFinal As String
Private InternalSF As String
Private InternalFH As Long 'fecha de habilitacion
Private InternalSV As Long 'version de cuando se habilita

Dim CC As New tbrCrypto.Crypt
Private PthApp As String

Public toSetKey As String

'Dim LB() As String

'sacado de 3PM con sistema de licencia anterior
Private Function SimpleUnEncript(TXT As String) As String
    'agregar un caracter cada un caracter del original
    'solo para que no se pueda buscar por texto
    Dim Largo As Long, LT As String, TMP As String
    Largo = Len(TXT)
    TMP = "": Dim A As Long
    For A = 1 To Largo
        LT = Mid(TXT, A, 1)
        'solo los caracteres impares valen
        If A / 2 <> A \ 2 Then
            'y los acumulo
            TMP = TMP + LT
        End If
    Next
    SimpleUnEncript = TMP
End Function

Public Function GetRF() As String
    GetRF = ResuFinal
End Function

Public Sub SetSF(defSF As String)
    'define para que soft se esta creando el archivo de datos
    InternalSF = defSF
End Sub

'///////////////////////////////////////////////////////
'CAMBIOS SET08 la fecha de creacion de la licencia y la version del sistema se guardan
'para cobrar por actualizaciones
Public Sub SetFH(newFH As Long)  'setear fecha de habilitacion
    InternalFH = newFH
    Terr.Anotar "ayaa", newFH
End Sub

Public Sub SetSV(newSV As Long)  'setear la version del sistema
    InternalSV = newSV
    Terr.Anotar "ayab", newSV
End Sub

Public Function GetFH() As Long 'los date no son un tipo de dato agradable
    GetFH = InternalFH
End Function

Public Function getSV() As Long '(major x 100.000) + (minor x 1.000) + revision
    getSV = InternalSV
End Function
'///////////////////////////////////////////////////////

Public Function GetSF() As String
    'obtenerlo del archivo que se esta leyendo
    GetSF = InternalSF
End Function

'obtiene el archivo LIC de datos de la PC y lo graba
Public Function DoNow(DestFinal As String) As Long

    On Local Error GoTo ErrDoNow
    
    Terr.Anotar "aac"
    
    Dim tmpDN As Long
    '0 todo ok
    tmpDN = 1
    Dim TI As New clstbrInfoPC
    
    tmpDN = 2
    ACUMTTT = TI.GetIM   'leo TODOs los datos
    Terr.Anotar "aad"
    
    tmpDN = 3
    ResuFinal = TI.ResFinal  'tomo solo lo que no cambia
    
    'buscar claves anteriores de 3PM!!!
    'en system32+dcilib22.dll esta encriptada muy feamente!
    '(uno sirve, uno no, etc)
    Dim FF As String, CL44 As String
    Terr.Anotar "aae"
    '///////////////////////////////////////////////////////
    'leer la licencia anterior es diferente segun el soft
    If LCase(InternalSF) = "3pm" Then
        FF = FSO.GetSpecialFolder(SystemFolder) + "\dcilib22.dll"
    End If
    
    If LCase(InternalSF) = "empirok" Then
        FF = FSO.GetSpecialFolder(SystemFolder) + "\wBrumiaz.dll"
    End If
    '///////////////////////////////////////////////////////
            
    If FSO.FileExists(FF) Then
        Dim TE2 As TextStream
        Set TE2 = FSO.OpenTextFile(FF, ForReading, False)
            CL44 = SimpleUnEncript(TE2.ReadLine)
        TE2.Close
    Else
        CL44 = "nox11"
    End If
    
    tmpDN = 4
    Terr.Anotar "aaf"
    
    ACUMTTT = ACUMTTT + vbCrLf + "EX3: " + CL44
    
    'nuevo agosto 2007!
    ACUMTTT = ACUMTTT + vbCrLf + "SFSFSF:" + InternalSF
    Terr.Anotar "aag", InternalSF
    
    'nuevo set 08!
    ACUMTTT = ACUMTTT + vbCrLf + "SV066:" + CStr(InternalSV)
    Terr.Anotar "aag44", InternalSV
    ACUMTTT = ACUMTTT + vbCrLf + "FH067:" + CStr(InternalFH)
    Terr.Anotar "aag45", InternalFH
    
    FF = FSO.GetSpecialFolder(SystemFolder) + "\nad3.cd"
        
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(FF, True)
        TE.Write ACUMTTT
    TE.Close
    Terr.Anotar "aah"
    tmpDN = 5
    'encriptarlo!!!
    
    CC.EncryptFile eMC_Blowfish, FF, DestFinal, KKY
    
    Terr.Anotar "aai"
    tmpDN = 6
    'borrar el original!
    FSO.DeleteFile FF, True
    Terr.Anotar "aaj"
    DoNow = 0
    Exit Function
    
ErrDoNow:
    Terr.AppendLog "AAA1-" + CStr(tmpDN), Terr.ErrToTXT(Err)
    MsgBox "Error en -DoNow- (" + CStr(tmpDN) + ")"
    DoNow = tmpDN
End Function

Public Function GetNR(AR As String, ByRef FF As String) As Long
    'obtiene el numero de respuesta NR que tiene algun archivo de licencia creado "ar"
    'devuelve en FF al final al que corresponde !!
    'ademas obtiene el fFINAL desencriptando todo para ver si corresponde a la misma
    'PC que trata de cargarlo!!!
    
    Dim tmpNR As Long
    Terr.Anotar "bsaa"
    'abrirlo y leerlo
    Dim FSO As New Scripting.FileSystemObject
    
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile(AR, ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    'al principio son X valores separados por puntos!
    Dim cFINAL() As Long 'clave final
    Dim A As Long, SP() As String
    If InStrRev(Txe, ".") = 0 Then
        tmpNR = -2 'no es aparentemente un archivo de licencia!
        GoTo FIN
    End If
    Terr.Anotar "bsab"
    
    SP = Split(Txe, ".")
    For A = 0 To UBound(SP) - 1 'el ultimo es "" y da error
        ReDim Preserve cFINAL(A)
        If Not IsNumeric(SP(A)) Then
            tmpNR = -2 'no es aparentemente un archivo de licencia!
            GoTo FIN
        End If
        cFINAL(A) = CLng(SP(A)) 'cargo uno por uno los valores
    Next A
    
    Terr.Anotar "bsac"
    
    'reconstruyo fFinal2
    Dim Letra As String, fFinal2 As String
    fFinal2 = ""
    Dim n As Long, testN As Long
    testN = -3 'bandera de que no leyo ninguno
    
    'no uso for por que no son ciclos iguales
    
    Dim V1 As Long 'contador de letras empieza en 1 cuando se genera el archivo
    V1 = 1
    Dim ix As Long 'valor dentro del split
    ix = 0
    Do
        If ix >= UBound(SP) Then Exit Do
        A = cFINAL(ix)
        Select Case A
            Case 60 To 69
            
                If testN = -3 Then testN = cFINAL(ix + 2) - V1
                n = cFINAL(ix + 2) - V1
                
                If testN <> n Then
                    tmpNR = (-V1 * 1000) - ix 'para saber donde da el error doy un numero de donde fue el problema
                    GoTo FIN
                End If
                
                fFinal2 = fFinal2 + Chr(cFINAL(ix + 1))
                
                ix = ix + 7
                V1 = V1 + 1
                
            Case 70 To 89
            
                If testN = -3 Then testN = cFINAL(ix + 2)
                n = cFINAL(ix + 2)
                
                If testN <> n Then
                    tmpNR = (-V1 * 1000) - ix 'para saber donde da el error doy un numero de donde fue el problema
                    GoTo FIN
                End If
                
                fFinal2 = fFinal2 + Chr(cFINAL(ix + 1))
                
                V1 = V1 + 1
                ix = ix + 4
                
            Case 90 To 170
            
                If testN = -3 Then testN = cFINAL(ix + 1) - (cFINAL(ix + 5))
                n = cFINAL(ix + 1) - (cFINAL(ix + 5))
                
                If testN <> n Then
                    tmpNR = (-V1 * 1000) - ix 'para saber donde da el error doy un numero de donde fue el problema
                    GoTo FIN
                End If
                
                fFinal2 = fFinal2 + Chr(cFINAL(ix + 5))
                
                V1 = V1 + 1
                ix = ix + 6
'            Case Else 'NO VAA
'                V1 = V1 + 1
        End Select
    Loop
    'salido de aqui solo llega si el nr es igual en todos lados
    'en ese caso fFinal2 tiene todo el texto
    Terr.Anotar "bsae"
    '****************
    'si llego hasta aqui son todos iguales!
    tmpNR = testN
    '****************
    
    Dim fFinal As String
    fFinal = "" 'traduzco en esta variable todo el contenido
    For A = 1 To Len(fFinal2)
        fFinal = fFinal + chgL2(Mid(fFinal2, A, 1))
    Next A
    'ahora es fFinal el que tiene el texto en espa�ol de todo
    
    'le saco el "final:"
    FF = Mid(fFinal, 7, Len(fFinal) - 6)
    Terr.Anotar "bsaf", FF
    
    'le saco el det de soft si lo tuviera
    'le saco tambien el "sv" y el "fh"
    Dim USF As Long 'ubicacion del sf si hubiera
    USF = InStr(1, FF, "SFSFSF:", vbTextCompare)
    
    Dim USF2 As Long 'ubicacion del sv si hubiera
    USF2 = InStr(1, FF, "SV066:", vbTextCompare)
    
    Dim USF3 As Long 'ubicacion del fh si hubiera
    USF3 = InStr(1, FF, "FH067:", vbTextCompare)
    
    Terr.Anotar "bsag", USF, USF2, USF3
    'definir interfal SF si estaba
    If USF > 0 Then
        If USF2 > 0 Then 'si es una version con "sv" si o si es version con "fh" (se agregaron las dos juntas en set08)
            'asi esta el original
            'fFinal = fFinal + "SFSFSF:" + InternalSF + "SV066:" + InternalSV + "FH067:" + InternalFH
            'por ejemplo 123451234533|mother?|mother?|micro?|disco|SFSFSF:3pmSV066:742677FH067:81031
            'usf=43  :  usf2= 53  :  usf3=65
            InternalSF = Mid(FF, USF + 7, USF2 - USF - 7)        'le resto los 6 del titulo que sigue y 1 del mid
            Dim tmp99 As String
            tmp99 = Mid(FF, USF2 + 6, USF3 - USF2 - 6)
            InternalSV = CLng(tmp99) 'le resto los 6 del titulo que sigue y 1 del mid
            tmp99 = Mid(FF, USF3 + 6)
            InternalFH = CLng(tmp99)
        Else
            'el ultimo parametro si no esta va hasta el final (p�ede ser una version sin SV que se agrego recien en set08)
            InternalSF = Mid(FF, USF + 7)
            InternalFH = 80923 'valores predeterminados (2008 set 23)
            InternalSV = 6     'valores predeterminados (0.00.006)
        End If
        'saco tambien la "|"
        FF = Left(FF, USF - 1)
    Else
        InternalSF = ""
        InternalFH = 80923 'valores predeterminados
        InternalSV = 6     'valores predeterminados
    End If
    
    Terr.Anotar "bsah", InternalSF, InternalSV, InternalFH
    
    If UCase(Left(fFinal, 6)) <> "FINAL:" Then
        tmpNR = -1 'no valido!!
    End If
    
FIN:
    Terr.Anotar "bsai", tmpNR
    GetNR = tmpNR
End Function

Public Function Get0406(AR As String) As String
    'puede dar error al desencriptar
    On Local Error GoTo NoEsArchDe3PM
        
    CC.DecryptFile eMC_Blowfish, AR, "c:\h.cda", KKY
    
    'abrirlo y desencriptarlo
    Dim FSO As New Scripting.FileSystemObject
    
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile("c:\h.cda", ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    FSO.DeleteFile "c:\h.cda", True
    
    Dim SP() As String
    SP = Split(Txe, vbCrLf)
    Dim A As Long
    'buscar solo los numeros que tomo como referencia
    Dim S4_6 As String
    
    For A = 0 To UBound(SP)
        'ver tambien si tenia clave
        If Left(SP(A), 3) = "EX3" Then
            Dim SP2() As String
            SP2 = Split(SP(A), ":")
            S4_6 = Trim(SP2(1))
        End If
    Next A
    Get0406 = S4_6
    
    Exit Function
NoEsArchDe3PM:
    Get0406 = "NO"
End Function

Public Function LiToLo(AR As String, AR2 As String) As Long
    On Local Error GoTo NoEsArchDe3PM2
    
    CC.DecryptFile eMC_Blowfish, AR, AR2, KKY
    LiToLo = 0
    Exit Function
NoEsArchDe3PM2:
    LiToLo = 1
End Function

Public Function GFlic(AR As String) As String
    'obtener solo el final de un .LIC
    Terr.Anotar "eaa"
    CC.DecryptFile eMC_Blowfish, AR, "c:\h2.cda", KKY
    
    'abrirlo y desencriptarlo
    Dim FSO As New Scripting.FileSystemObject
    
    Terr.Anotar "eab"
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile("c:\h2.cda", ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    Terr.Anotar "eac"
    FSO.DeleteFile "c:\h2.cda", True
    
    Dim SP() As String, fFinal As String: fFinal = "NiBosta"
    SP = Split(Txe, vbCrLf)
    Dim A As Long
    'buscar solo los numeros que tomo como referencia
    Terr.Anotar "ead"
    
    For A = 0 To UBound(SP)
'        'lo hago en una funcion aparte
'        'ver tambien si tenia clave
'        If Left(SP(A), 3) = "EX3" Then
'            Dim SP2() As String
'            SP2 = Split(SP(A), ":")
'            S4_6 = Trim(SP2(1))
'        End If
        If UCase(Left(SP(A), 6)) = "FINAL:" Then
            fFinal = Trim(Mid(SP(A), 7, Len(SP(A)) - 6)) 'con la palabra final y todo (de paso me sirve para validar _
                que es un codigo valido!)
            Exit For
        End If
    Next A
    
    GFlic = fFinal
End Function

Public Function CK(AR As String, NR As Long, dest As String) As Long
    '***************************************
    'crear un archivo de licencia dado un archivo de datos (AR)
    'puede recibir como par�metro AR el ("FINAL"+"SFSFSF") si no existe el .LIC original
    '***************************************
    On Local Error GoTo NoEsArchDe3PM
    Dim fFinal As String
    fFinal = "NiBosta" 'por las dudas que no entre
    
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(AR) Then
        'el archivo de licencia aplicado al archivo de datos dados se hara para
        'que devuelva el numero (NR) solicitado
        'este ultimo lo uso para saber que tipo de licecnia es
        
        'El archivo AR esta encriptado
        CK = 1
        Terr.Anotar "daa"
        CC.DecryptFile eMC_Blowfish, AR, "c:\h.cda", KKY
        
        'abrirlo y desencriptarlo
        CK = 2
        Terr.Anotar "dab"
        Dim TE As TextStream
        Dim Txe As String 'todo el texto
        Set TE = FSO.OpenTextFile("c:\h.cda", ForReading, False)
            Txe = TE.ReadAll
        TE.Close
        
        CK = 3
        Terr.Anotar "dac"
        FSO.DeleteFile "c:\h.cda", True
        
        Dim SP() As String
        SP = Split(Txe, vbCrLf)
        Dim A As Long
        'buscar solo los numeros que tomo como referencia
        Terr.Anotar "dad"
        
        InternalSF = "" 'si no lo lee sera alguno viejo que no lo tenia a�n
        InternalFH = 80923 'valores predeterminados son los 3
        InternalSV = 6
        
        For A = 0 To UBound(SP)
    '        'lo hago en una funcion aparte
    '        'ver tambien si tenia clave
    '        If Left(SP(A), 3) = "EX3" Then
    '            Dim SP2() As String
    '            SP2 = Split(SP(A), ":")
    '            S4_6 = Trim(SP2(1))
    '        End If
            If UCase(Left(SP(A), 6)) = "FINAL:" Then
                fFinal = SP(A) 'con la palabra final y todo (de paso me sirve para validar _
                    que es un codigo valido!)
            End If
            
            'leo para que software fue hecho!
            If UCase(Left(SP(A), 7)) = "SFSFSF:" Then
                'le saco el prefijo sfsfsf:
                InternalSF = Mid(SP(A), 8)
                Terr.Anotar "dae", InternalSF
            End If
            
            If UCase(Left(SP(A), 6)) = "SV066:" Then
                InternalSV = CLng(Mid(SP(A), 7))
                Terr.Anotar "dae22", InternalSV
            End If
            
            If UCase(Left(SP(A), 6)) = "FH067:" Then
                'le saco el prefijo sfsfsf:
                InternalFH = CLng(Mid(SP(A), 7))
                Terr.Anotar "dae23", InternalFH
                'DEJO MI FECHA, ANADA SABER COMO ANDA EL RELOJ DEL USUARIO!!! ... bueno espero que el de tbr ande bien
                'ult digito a�o x 10.000 + mes x 100 + dia
                InternalFH = ((Year(Date) - 2000) * 10000) + (Month(Date) * 100) + (Day(Date))
                Terr.Anotar "dae23b", InternalFH
            End If
        Next A
        
        'nuevo.. ahora el soft tambien indica diferencias
        fFinal = fFinal + "SFSFSF:" + InternalSF + "SV066:" + CStr(InternalSV) + "FH067:" + CStr(InternalFH)
    Else 'es posible que venga no de un archivo LIC si no de los parametros sueltos
        'cuando viene de los parametros sueltos tiene que venir
        'nofile+chr(5)+final+chr(5)+sfsfsf
        '=> sp(0)="nofile"
        '=> sp(1)=final
        '=> sp(2)=sfsfsf ' para que software fue hecho
        Dim SP71() As String
        SP71 = Split(AR, Chr(5))
        If UBound(SP71) < 4 Or LCase(SP71(0)) <> "nofile" Then
            'quiso pasar un archivo que no existe!
            CK = 15
            Exit Function
        End If
        
        fFinal = "FINAL:" + SP71(1) + _
                 "SFSFSF:" + SP71(2) + _
                 "SV066:" + SP71(3) + _
                 "FH067:" + SP71(4)
    End If
    
    CK = 4
    
    Terr.Anotar "daf"
    Set CC = Nothing
    
    'fFINAL tiene el codigo a partir del cual se generara
    'primero hago reemplazos varios
    Dim Letra As String, fFinal2 As String
    fFinal2 = ""
    'ReDim Preserve LB(0)
    For A = 1 To Len(fFinal) 'entre 50 y 150 aprox
        Letra = LCase(Mid(fFinal, A, 1))
        Letra = chgL(Letra)
    '    ReDim Preserve LB(A)
    '    LB(A) = LCase(Mid(fFinal, A, 1)) + " " + Letra
        'If Letra = "/" Then MsgBox Mid(fFinal, A, 1)
        fFinal2 = fFinal2 + Letra
    Next A
    
    CK = 5
    
    'aplico otros cambios
    Dim supF As String: supF = ""
    For A = 1 To Len(fFinal2) 'entre 50 y 150 aprox
        Dim p4 As String 'cuatro caracteres por cada uno
        Letra = Mid(fFinal2, A, 1)
        p4 = GetP4(Letra, A, NR)
        supF = supF + p4
    Next A
    
    CK = 6
    
    'escribo en el destino indicado
    Dim TE4 As TextStream
    Set TE4 = FSO.CreateTextFile(dest, True)
        TE4.Write supF
    TE4.Close
    
    CK = 0
    
    Exit Function
NoEsArchDe3PM:
    'sale con el CK en el error, solo cero esta ok
End Function

'Public Function GetLBs() As String()
'    GetLBs = LB
'End Function

Private Function P4to1(Cuatro As String, vA As Long, ByRef NR As Long) As String
    'recibe cuatro letras y devuelve una letra. Funcion inversa de GetP4
    'segun GetP4 estos representan:
    'llamamos v1 a mid(cuatro,1,1), llamamos v2 a mid(cuatro,2,1), etc etc
    'v1 = chr(abs(v2-v1) _
        donde v1 es el indice desde 1 hasta len(fFinal2) _
        donde v2 es el NR
    '=> v1 = chr(abs(NR-A))
    
    'el valor A es un dato puesto como vA
    
    'NR devuelve el NR de cada cuarteto. TODOS DEBEN SER IGUALES
    
    '**************************************************************
    'los valores que necesigto son:
        'A indice MID del fFinal 2
        'NR numero de respuesta que se quiso colocar en el archivo
        'Caracter de fFinal 2 para armar el paquete
    Dim Tmp_A As Long: Tmp_A = vA
    Dim Tmp_NR As Long
    Dim Tmp_CHR As String
    '**************************************************************
    
    Dim T1 As Integer, T2 As Integer, T3 As Integer, T4 As Integer
    T1 = Asc(Mid(Cuatro, 1, 1))
    T2 = Asc(Mid(Cuatro, 2, 1))
    T3 = Asc(Mid(Cuatro, 3, 1))
    T4 = Asc(Mid(Cuatro, 4, 1))
    
    'ver si es abs(NR-A) es NR-A o A-NR (saber cuale s mayor)
    If T2 >= 128 Then 'NR>A
        Tmp_NR = Tmp_A - T1
    Else 'A>NR
        Tmp_NR = Tmp_A + T1
    End If
    
    If T4 >= 100 Then 'NR>=tmp_chr
        Tmp_CHR = Chr(Tmp_NR - T3)
    Else 'tmp_chr>NR
        Tmp_CHR = Chr(Tmp_NR + T3)
    End If
    
    NR = Tmp_NR
    P4to1 = Tmp_CHR
    
End Function

Private Function GetP4(s As String, V1 As Long, v2 As Long) As String
    'devulelve 4 letras en base al caracter y dos valores variables
    'a del for y el NR esperado
    
    'pongo los 3 como numeros
    Dim v0 As Long
    v0 = Asc(s)
    
    'v0 es el valor de la letra actual. de 0 a 255
    'v1 es el a del for o sea va de 1 a 50-150)
    'v2 es el nr que como lo defino yo va de 1 a 100
    
    'debe ser una funcion reversible !
    
    'para los valores previos temporales numericos
    Dim TP(6) As Long
    
    Dim fFIN As String 'si o si empeiza en numero y termina en punto
    
    'el primero de los 4 es al azar
    Randomize
    TP(0) = CLng(Rnd * 100) + 60 'da entre 60 y 160
    
    If TP(0) < 70 Then
        TP(1) = v0
        TP(2) = V1 + v2
        TP(3) = v2 + v0
        TP(4) = TP(0) + 1 'si o si distinto de cero
        TP(5) = TP(0) - 40 'si o si distinto de cero
        TP(6) = TP(0) + V1 'si o si distinto de cero
        fFIN = CStr(TP(0)) + "." + _
               CStr(TP(1)) + "." + _
               CStr(TP(2)) + "." + _
               CStr(TP(3)) + "." + _
               CStr(TP(4)) + "." + _
               CStr(TP(5)) + "." + _
               CStr(TP(6)) + "."
        
    Else
        If TP(0) < 90 Then 'de 71 a 89
            TP(1) = v0
            TP(2) = v2
            TP(3) = 6
            TP(4) = 0
            TP(5) = 0
            TP(6) = 0
            
            fFIN = CStr(TP(0)) + "." + _
               CStr(TP(1)) + "." + _
               CStr(TP(2)) + "." + _
               CStr(TP(3)) + "."
        Else
            'entre 91 y 160
            TP(1) = v0 + v2
            TP(2) = V1 * 2
            TP(3) = V1 + 19
            TP(4) = TP(1) 'si o si distinto de cero
            TP(5) = v0 'si o si distinto de cero
            TP(6) = 0
            
            fFIN = CStr(TP(0)) + "." + _
               CStr(TP(1)) + "." + _
               CStr(TP(2)) + "." + _
               CStr(TP(3)) + "." + _
               CStr(TP(4)) + "." + _
               CStr(TP(5)) + "."
        End If
    End If
    
    GetP4 = fFIN
End Function

Public Function GetDiff2PC(ID1 As String, ID2 As String) As String
    'devuelve la diferencia de hardware entre 2 equipos
    Dim Ret(20) As Long 'cada una de los valores con informacion
    ' de 0 a 6 son diferencias en las placas de red existentes
    
    'NO SE INICIALIZO!!!!!
    Dim M As Long
    For M = 0 To 20
        Ret(M) = 0
    Next M
    On Local Error GoTo ErrDIFF
    
    Dim SP1() As String, SP2() As String
    
    SP1 = Split(ID1, "|")
    SP2 = Split(ID2, "|")
    
    'si alguno de los dos no tiene la cantidad de indices deseados salgo IA!
    '000E7B034604.0012F0323ADC|          |00000000000006D8|D080006|454X0713S      |
    '000BCD56747A.000D9D46EC17|CNF3230D95|BFEBF9FF00000F27|207000F|6G2722         |
    '00138F9B3DD2             |00000000  |BFEBFBFF00000F49|409000F|WD-WMAP91962322|
    
    'donde
    'SP(0) = Lista de MACS separadas por .
    'SP(1) = BIOS
    'SP(2) = Micro
    'SP(3) = Micro Reserved
    'SP(4) = DISCO 1
    'SP(5) = DISCO 2
    'SP(6) = DISCO 3
    'SP(7) = DISCO 4
    
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
    'ret(11):
    'ret(12):
    'ret(13):
    'ret(14):
    'ret(15):
    'ret(16):
    'ret(17):
    'ret(18):
    'ret(19):
    
    'son por lo menos 5 (el (5) vacio si hay solo un disco)
    
    If UBound(SP1) < 5 Then
        Ret(0) = 1
        Exit Function
    End If
    
    If UBound(SP2) < 5 Then
        Ret(1) = 1
        Exit Function
    End If
    
    'sublistas de MACs
    Dim SP1B() As String
    Dim SP2B() As String
    
    SP1B = Split(SP1(0), ".")
    SP2B = Split(SP2(0), ".")
    
    Ret(2) = UBound(SP1B) + 1 'cantidad de placas de la PC1
    Ret(3) = UBound(SP2B) + 1 'cantidad de placas de la PC1
    Ret(4) = 0 'cantida de coincidencias
    'Ya tengo todo lo que necesito para comparar
    Dim A As Long, B As Long
    For A = 0 To UBound(SP1B)
        For B = 0 To UBound(SP2B)
            If LCase(SP1B(A)) = LCase(SP2B(B)) Then
                Ret(4) = Ret(4) + 1 'cantidad de coincidencias
                GoTo nNext
            End If
        Next B
nNext:
    Next A
    
    'DIF en LA BIOS
    If Trim(LCase(SP1(1))) <> Trim(LCase(SP2(1))) Then
        Ret(5) = 6
    Else
        Ret(5) = 4
    End If
    
    'DIF en el micro
    If Trim(LCase(SP1(2))) <> Trim(LCase(SP2(2))) Then
        Ret(6) = 1
    Else
        Ret(6) = 2
    End If
    
    'DIF en reserved
    If Trim(LCase(SP1(3))) <> Trim(LCase(SP2(3))) Then
        Ret(7) = 3
    Else
        Ret(7) = 5
    End If
    
    'DIF en discos
    'ver la cantidad
    Ret(8) = UBound(SP1) - 4 'cantidad de discos duros en SP1
    Ret(9) = UBound(SP2) - 4 'cantidad de discos duros en SP1
    Ret(10) = 0 'cantidad de coincidencias
    'pasa lo que pase comparo todos con todos
    Dim J As Long, k As Long
    For J = 4 To UBound(SP1) - 1
        For k = 4 To UBound(SP2) - 1
            If Trim(LCase(SP1(J))) = Trim(LCase(SP2(k))) Then
                Ret(10) = Ret(10) + 1
                GoTo MEXT
            End If
        Next k
MEXT:
    Next J
    
    
    'pasar todo a string
    Dim SS As String
    SS = ""
    For J = 0 To UBound(Ret)
        SS = SS + CStr(Ret(J))
    Next J
    
    GetDiff2PC = SS
    Exit Function
    
ErrDIFF:
    Ret(20) = Ret(20) + 1
    Resume Next
End Function

Public Function GetPtosDiff(ID1 As String, ID2 As String) As Long 'devuelve los puntos

    Dim DF As String
    DF = GetDiff2PC(ID1, ID2)
    
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
    
VALIDAR:
    GetPtosDiff = PT1
End Function

Private Function chgL(s As String) As String
    Select Case s
        Case "a": chgL = "a"
        Case "b": chgL = "f"
        Case "c": chgL = "b"
        Case "d": chgL = "g"
        Case "e": chgL = "k"
        Case "f": chgL = "c"
        Case "g": chgL = "h"
        Case "h": chgL = "i"
        Case "i": chgL = "j"
        Case "j": chgL = "d"
        Case "k": chgL = "e"
        Case "l": chgL = "s"
        Case "m": chgL = "t"
        Case "n": chgL = "u"
        Case "o": chgL = "r"
        Case "p": chgL = "v"
        Case "q": chgL = "w"
        Case "r": chgL = "x"
        Case "s": chgL = "y"
        Case "t": chgL = "l"
        Case "u": chgL = "m"
        Case "v": chgL = "0"
        Case "w": chgL = "1"
        Case "x": chgL = "z"
        Case "y": chgL = "n"
        Case "z": chgL = "2"
        Case "0": chgL = "3"
        Case "1": chgL = "o"
        Case "2": chgL = "4"
        Case "3": chgL = "9"
        Case "4": chgL = "8"
        Case "5": chgL = "p"
        Case "6": chgL = "7"
        Case "7": chgL = "6"
        Case "8": chgL = "5"
        Case "9": chgL = "q"
        Case " ": chgL = "�"
        Case "-": chgL = "("
        Case ".": chgL = ")"
        Case "+": chgL = "."
        Case "(": chgL = "+"
        Case ")": chgL = "-"
        Case "�": chgL = " "
        Case ":": chgL = "|"
        Case "|": chgL = ":"
        Case Else: chgL = "/"
    End Select
End Function

Private Function chgL2(s As String) As String 'inversa de la anterior
    Select Case s
        Case "a": chgL2 = "a"
        Case "f": chgL2 = "b"
        Case "b": chgL2 = "c"
        Case "g": chgL2 = "d"
        Case "k": chgL2 = "e"
        Case "c": chgL2 = "f"
        Case "h": chgL2 = "g"
        Case "i": chgL2 = "h"
        Case "j": chgL2 = "i"
        Case "d": chgL2 = "j"
        Case "e": chgL2 = "k"
        Case "s": chgL2 = "l"
        Case "t": chgL2 = "m"
        Case "u": chgL2 = "n"
        Case "r": chgL2 = "o"
        Case "v": chgL2 = "p"
        Case "w": chgL2 = "q"
        Case "x": chgL2 = "r"
        Case "y": chgL2 = "s"
        Case "l": chgL2 = "t"
        Case "m": chgL2 = "u"
        Case "0": chgL2 = "v"
        Case "1": chgL2 = "w"
        Case "z": chgL2 = "x"
        Case "n": chgL2 = "y"
        Case "2": chgL2 = "z"
        Case "3": chgL2 = "0"
        Case "o": chgL2 = "1"
        Case "4": chgL2 = "2"
        Case "9": chgL2 = "3"
        Case "8": chgL2 = "4"
        Case "p": chgL2 = "5"
        Case "7": chgL2 = "6"
        Case "6": chgL2 = "7"
        Case "5": chgL2 = "8"
        Case "q": chgL2 = "9"
        Case "�": chgL2 = " "
        Case "(": chgL2 = "-"
        Case ")": chgL2 = "."
        Case ".": chgL2 = "+"
        Case "+": chgL2 = ")"
        Case "-": chgL2 = "("
        Case " ": chgL2 = "�"
        Case ":": chgL2 = "|"
        Case "|": chgL2 = ":"
        Case Else: chgL2 = "/"
    End Select
End Function

Public Sub SetLog(sDir As String)
    Terr.FileLog = sDir
    Terr.Anotar "aab"
    
    Dim GF As String
    GF = FSO.GetBaseName(sDir) + "." + FSO.GetExtensionName(sDir)
    'le saco el nombre que tenga y me quedo con la carpeta
    GF = Mid(sDir, 1, Len(sDir) - Len(GF))
    If Right(GF, 1) <> "\" Then GF = GF + "\"
    
    PthApp = GF
    
    CC.SetLog PthApp + "regCR.log"
    
End Sub

Public Sub kkyNew(k As String)
    KKY = k
End Sub

Private Sub Class_Initialize()
    
    KKY = "siboldi98" 'dic2010 para que no ande suelto
    
    InternalSF = "" 'valor predeterminado compatible con versiones anteriores
    InternalFH = 80923 'fecha de habilitacion 23 set 2008 inicio de este cambio = _
        ult digito a�o x 10.000 + mes x 100 + dia = _
        80.000 + 900 + 23 = 80923
    InternalSV = 6 'version de cuando se habilita

    Terr.FileLog = App.path + "\regKI.log"
    Terr.Set_ADN CStr(App.Major * 100000 + App.Minor * 1000 + App.Revision)
    
    Terr.LargoAcumula = 600
    Terr.Anotar "aaa"
    
    CC.SetLog App.path + "\regCR.log"
    
End Sub
