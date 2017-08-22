VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMMTypeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mmTMP() As clsMMType

Public Function AddMM() As clsMMType
    Dim j As Long
    j = UBound(mmTMP) + 1
    ReDim Preserve mmTMP(j)
    Set AddMM = mmTMP(j)
End Function


Public Function ObtenerArchMM(Carpeta As String, _
    Optional ordenarABC As Boolean = False, Optional buscarPerfil As Long = 0) As String()
    
    'devuelve "Carpeta + NombreArchivo + "#" + NombreArchivo"
    'devuelve PathFull,SoloNombre
    
    'buscarPerfil se agrega 14/08/2008 para tener en cuenta discos de ringtones, wallpapers y java
    'es un valor de etrada y de salida tambien
    'cuando entra puede ser 0 (cero) para trabajar normal como siempre, solo discos de musica y videos
    'cuando entra en 1 analiza automáticamente perfiles
    'si entra en 100 + X interpreto que esta forzando el disco a ser del perfil X (no programado aun xxxx)

    ' = 1 basico de multimedia
    ' = 2 disco de ringtones
    ' = 3 disco de wallpapers
    ' = 4 discos de java
    ' = 5 disco de imagenes iso / nrg / etc
    ' = 6 disco de videos 3gp
    ' = 7 disco de temas para celular

    'ADEMÁS DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    On Error GoTo ErrObtMM
    tERR.Anotar "001-0214", Carpeta, ordenarABC, buscarPerfil
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    tERR.Anotar "001-0215", Carpeta
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    Dim NombreArchivo As String, ContadorArch As Long, newName As String
    
    Dim EEXX() As String
    
    Select Case buscarPerfil
        Case 0, 101 'no deberia llamarse como 101 pero por prolijidad esta. esta eleccion es fonola base
            ReDim EEXX(9) 'solo lo de siemrpe basico del 3PM
            EEXX(0) = "mp3"
            EEXX(1) = "wma"
            EEXX(2) = "mpg"
            EEXX(3) = "mpeg"
            EEXX(4) = "avi"
            EEXX(5) = "vob"
            EEXX(6) = "mn0"
            EEXX(7) = "mn1"
            EEXX(8) = "dat"
            EEXX(9) = "wmv"
        Case 2, 102 'ringtones solo en MP3
            ReDim EEXX(0)
            EEXX(0) = "mp3"
        Case 3, 103 'wallpapers
            ReDim EEXX(3)
            EEXX(0) = "jpg" 'NO CONFUNDIR con tapas de discos!
            EEXX(1) = "jpeg"
            EEXX(2) = "bmp"
            EEXX(3) = "gif"
        Case 4, 104 'java
            ReDim EEXX(1)
            EEXX(0) = "jar"
            EEXX(1) = "jad"
        Case 5, 105 'mm91 imagenes iso/nero
            ReDim EEXX(13)
            EEXX(0) = "iso"
            EEXX(1) = "nrg"
            EEXX(2) = "nr3"
            EEXX(3) = "nra"
            EEXX(4) = "nrb"
            EEXX(5) = "nrc"
            EEXX(6) = "nrd"
            EEXX(7) = "nre"
            EEXX(8) = "nrh"
            EEXX(9) = "nri"
            EEXX(10) = "nrm"
            EEXX(11) = "nru"
            EEXX(12) = "nrv"
            EEXX(13) = "nrw"
        Case 6, 106 'mm91 videos 3gp
            ReDim EEXX(0)
            EEXX(0) = "3gp"
        Case 7, 107
            ReDim EEXX(1)
            EEXX(0) = "thm"
            EEXX(1) = "nth"
        Case 1 'deteccion automática
            ReDim EEXX(35) 'ampliado a cualquiera de los perfiles mm91
            EEXX(0) = "mp3"
            EEXX(1) = "wma"
            EEXX(2) = "mpg"
            EEXX(3) = "mpeg"
            EEXX(4) = "avi"
            EEXX(5) = "vob"
            EEXX(6) = "mn0"
            EEXX(7) = "mn1"
            EEXX(8) = "dat"
            EEXX(9) = "wmv"
            EEXX(10) = "" 'reservado para futuros archivos multimedia wav, mp4, midi
            EEXX(11) = ""
            EEXX(12) = ""
            'rigtones "mp3" que ya se leen de todas formas
            EEXX(13) = "" 'deberia identificarlos por el largo NO SE USA CON ESTA EXTENCION, ES SOLO PARA DEMOSTRAR QUE LOS BUSCO
            'wallpapers "jpg", "jpeg", "bmp", "gif"
            EEXX(14) = "jpg" 'NO CONFUNDIR con tapas de discos!
            EEXX(15) = "jpeg"
            EEXX(16) = "bmp"
            EEXX(17) = "gif"
            EEXX(18) = "jar" 'aplicaciones o juegos java
            EEXX(19) = "iso" 'imaganes iso/nero     'mm91
            EEXX(20) = "nrg"
            EEXX(21) = "nr3"
            EEXX(22) = "nra"
            EEXX(23) = "nrb"
            EEXX(24) = "nrc"
            EEXX(25) = "nrd"
            EEXX(26) = "nre"
            EEXX(27) = "nrh"
            EEXX(28) = "nri"
            EEXX(29) = "nrm"
            EEXX(30) = "nru"
            EEXX(31) = "nrv"
            EEXX(32) = "nrw"
            EEXX(33) = "3gp" 'videos para movil  'mm91
            EEXX(34) = "thm" 'temas para movil 'mm91
            EEXX(35) = "nth"
    End Select
    
    Dim ArchMMBase As Long 'cantidad de archivos de musica y videos
    Dim ArchJava As Long 'cantidad de archivos de musica y videos
    Dim ArchImagen As Long 'cantidad de imagenes
    Dim ArchMMRingtone As Long 'cantidad de mp3s o wmas de menos de 1,5 mb (hacer configurable)
    Dim ArchKaraoke As Long 'cantidad de mp3s o wmas de menos de 1,5 mb (hacer configurable)
    Dim ArchTotales As Long 'para saber proporciones de cada uno
    Dim ArchISO As Long 'imagenes ISO o de nero 'mm91
    Dim Arch3GP As Long  'videos para movil 'mm91
    Dim ArchThemes As Long 'temas para movil
    
    'una vez cargado esto por fuera se define un perfil del disco para ver de que tipo es
    'tambie estaría bueno definir un origen ya con características de tipo de disco
    'por que no solo son disco multimedia si no que hay
    'discos de aplicaciones JAVA
    'discos de ringtones (ideal para confundirse con de mp3)
    'discos de wallapers
    
    'deberia dar un perfil a cada disco, si por ejemplo hay muchos mp3 de _
        mas de 2 minutos y aparece un mp3 de 30 segundos y 2 o 3 imagenes _
        deberia darse cuenta que es un disco comun _
        Si la mayoria son imagenes es una carpeta de wallpapers _
        Si la mayoria son mp3 de menos de 2 minutos es una carpeta de ringtones _
        Si hay 10 jar y 10 jpg debo interpretar que es una carpeta de java con _
        sus screenshots correspondientes
        
    'para esto hay un dato que sirve, una cancion MP3 con calidad comun ocupa 1 MB por cada minuto
    'POR ejemplo en el disco "En vivo en Cemento (10-10-1998) - A morir !!! de catupecu _
        hay 4 canciones de menos de 1.5 MB pero hay 18 MP3s de mas de 1.5 MB _
        el sistema debería identificar este disco como de musica !
    
    ArchMMBase = 0
    ArchJava = 0
    ArchImagen = 0
    ArchMMRingtone = 0
    ArchKaraoke = 0
    ArchTotales = 0
    ArchISO = 0 'mm91
    Arch3GP = 0 'mm91
    ArchThemes = 0
    
    Dim H As Long
    For H = 0 To UBound(EEXX)
        If EEXX(H) = "" Then GoTo sigEstaVacio
        
        NombreArchivo = Dir$(Carpeta + "*." + EEXX(H))
        Do While Len(NombreArchivo)
            tERR.Anotar "001-0217", NombreArchivo, EEXX(H)
            'corregir el nombre del tema
            newName = Replace(NombreArchivo, ",", "")
            newName = Replace(newName, "#", "")
            If NombreArchivo <> newName Then
                'no se puede corregir si es un CD. Solo corrige si es disco duro
                'esta funcion se usa para leer CDs debo prevenir
                tERR.Anotar "001-0220", newName
                If FSo.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                    tERR.Anotar "001-0221"
                    'ver si existe lo que se esta por escribir
                    'si es asi elimino el actual
                    If FSo.FileExists(Carpeta + newName) Then
                        FSo.DeleteFile Carpeta + NombreArchivo, True
                    Else
                        FSo.MoveFile Carpeta + NombreArchivo, Carpeta + newName
                    End If
                    tERR.Anotar "001-0222"
                    NombreArchivo = newName
                End If
            End If
            ContadorArch = ContadorArch + 1
            
            'el hecho de diferenciar los tipos de archivos de una carpeta es exclusivo
            'de cuado se le pide a esta funcion que defina el perfil automáticamente
            If buscarPerfil = 1 Then
                Select Case H 'segun este indice es un tipo de archivo diferente
                    Case 0 'mp3
                        'ver si es arch grande o chico!
                        Dim SzeFil As Long
                        SzeFil = FileLen(Carpeta + NombreArchivo)
                        If SzeFil > CLng(1572864) Then ' <<1.5 * (1024 * 1024)>> 1,5 MB es mi base
                            ArchMMBase = ArchMMBase + 1
                        Else
                            ArchMMRingtone = ArchMMRingtone + 1
                        End If
                    Case 1, 2, 3, 4, 5, 8, 9 'wma, mpg, mpeg, avi, vob, dat, wmv
                        ArchMMBase = ArchMMBase + 1
                    Case 6, 7 'mn0,mn1
                        ArchKaraoke = ArchKaraoke + 1
                    Case 14, 15, 16, 17
                        ArchImagen = ArchImagen + 1
                    Case 18
                        ArchJava = ArchJava + 1
                    Case 19 To 32 'mm91
                        ArchISO = ArchISO + 1
                    Case 33 'mm91
                        Arch3GP = Arch3GP + 1
                    Case 34, 35 'mm91
                        ArchThemes = ArchThemes + 1
                End Select
                
                ArchTotales = ArchTotales + 1
            End If
            
            ReDim Preserve TMPmatriz(ContadorArch)
            tERR.Anotar "001-0225", ContadorArch
            TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
            NombreArchivo = Dir$
        Loop
sigEstaVacio:
    Next H

    tERR.Anotar "001-0256", buscarPerfil, ArchMMBase, ArchKaraoke, ArchTotales
    tERR.Anotar "001-0256-b", ArchMMRingtone, ArchImagen, ArchJava, ArchISO
    '//////////////////////////////////////////////
    'definir el perfil del disco
    Dim PerfilFINAL As Long 'identificador del perfil
    ' = 1 basico de multimedia
    ' = 2 disco de ringtones
    ' = 3 disco de wallpapers
    ' = 4 discos de java
    
    PerfilFINAL = 0 'quiere decir que es 100% basico (por ahora)
    'pero si pide otra cosa ...
    If buscarPerfil > 0 Then
        'si puso 1 es que quiere que hagamos esto de buscarlo nosotros
        If buscarPerfil = 1 And ArchTotales > 0 Then
            'PERFIL BASICO DE 3PM, MULTIMEDIA PARA ESCUCHAR
            Dim PROP As Single
            PROP = (ArchMMBase + ArchKaraoke) / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 1
            End If
            'perfil de ringtones
            PROP = ArchMMRingtone / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 2
            End If
            'perfil de wallpapers
            PROP = ArchImagen / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 3
            End If
            'perfil de java
            PROP = ArchJava / ArchTotales
            If PROP > 0.4 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 4
            End If
            PROP = ArchISO / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 5
            End If
            PROP = Arch3GP / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 6
            End If
            PROP = ArchThemes / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 7
            End If
            
            'si hay pocos archivos no me pongo a renegar, es un multimedia
            If ArchTotales < 3 Then PerfilFINAL = 1
            
            '//////////////////////////////////////////////
            'si no entro a ninguno dejo el predeterminado
            If PerfilFINAL = 0 Then PerfilFINAL = 1
            '//////////////////////////////////////////////
            
            
            'ahora que defini el perfil quitar todos los archivo que no correspondan estar en esta lista
            'segun el perfil que se ha determinado correcto.
            'por ejemplo si determinamos que el perfil es java saco de esta lista los archivo de karaokes
            Dim mm As Long, Cancion As String, totEliminados As Long, cadaExtencion As String
            totEliminados = 0
            For mm = 1 To UBound(TMPmatriz) 'esta en base 1
                Cancion = txtInLista(TMPmatriz(mm), 1, "#")
                cadaExtencion = LCase(FSo.GetExtensionName(Cancion)) 'minuscula por si quiero comparar con EEXX
                Select Case PerfilFINAL
                    Case 1 'base comun
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If H >= 13 Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 2 'rigtones
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If H > 0 Then 'solo los mp3 y naaaada mas va aqui
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 3 'es de wallpapers
                        'la tapa de los wallpapers no se debe mostrar
                        If LCase(Cancion) = "tapa.jpg" Then TMPmatriz(mm) = ""
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H > 17) And (H < 14) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 4 'es de java
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H <> 18) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 5 'es una imagen de disco 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H < 19) Or (H > 32) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 6 'es video para movil 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H <> 33) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 7 'temas para movil 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H < 34) And (H > 35) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
finFOR:
                End Select
            Next mm 'fin de poner en "" todos los que no van
            
            'quitar de la matriz los que no van
            limpiarMtxVacios TMPmatriz
            'XXXXXXXXXXXXXXXXXXXXXXXXXX
            'XXXX queda revisar tooooooodas las llamadas al obtenerArchMM
            'y probar esta funcion que no esta probada !!!
            'ver que detecte los perfiles joiaaaaaaa
            'hacer que muestre cada perfil como corresponde en 3PM
            'XXXXXXXXXXXXXXXXXXXXXXXXXX
        End If
        
        'los otros casos son mas directos
        If buscarPerfil > 100 Then
            PerfilFINAL = buscarPerfil - 100
        End If
        
        
    End If
    '//////////////////////////////////////////////
    
    
    '//////////////////////////////////////////////
    'devuelvo el resultado
    buscarPerfil = PerfilFINAL
    '//////////////////////////////////////////////
    
    'XXXX
    'este ordenar lee las matrices desde 1 hasta ubound
    'y mi matriz usa hasta el cero
    
    'por otra parte no lee todos los discos de wallapers al iniciar
    'igual no puse perfil automático de 3pm al iniciar
    'ya que alli hay una revision de que discos se incluyen y cuales
    'no y se base en que tengan archivos mp3s
    
    If ordenarABC Then
        Dim TMP2Matriz() As String
        ReDim TMP2Matriz(0)
        Dim OKs As Long
        
        Dim k As Long, L As Long
        Dim CPR As String 'comparador
        CPR = "ZZZ"
        Dim Min As Long 'indice del minimo
        Min = 0
        For L = 1 To UBound(TMPmatriz) 'esta en base 1
            CPR = "ZZZZZZZZZ"
            For k = 1 To UBound(TMPmatriz) 'esta en base 1
                If TMPmatriz(k) < CPR Then
                    CPR = TMPmatriz(k)
                    Min = k
                End If
            Next k
            
            OKs = OKs + 1
            ReDim Preserve TMP2Matriz(OKs)
            TMP2Matriz(OKs) = CPR
            TMPmatriz(Min) = "ZZZZZZZZZ"
        Next L
        
        ObtenerArchMM = TMP2Matriz
        
    Else
        ObtenerArchMM = TMPmatriz
    End If
    
    Exit Function
ErrObtMM:
    tERR.AppendLog tERR.ErrToTXT(Err), "Archivos.bas" + ".acpk4"
    Resume Next
    
End Function

'dupli sirve en obtenerarchmm
Private Function limpiarMtxVacios(ByRef mtx() As String)
    Dim H As Long, Listo As Boolean
    Listo = False
    Do While Listo = False
        Listo = True
        For H = 1 To UBound(mtx)
            If mtx(H) = "" Then
                quitarElemMatriz mtx, H
                Listo = False 'lo hace quedarse una vuelta mas por las dudas
                Exit For
            End If
        Next H
    Loop
End Function

'dupli sirve en limpiarmtxvacios
Private Function quitarElemMatriz(ByRef mtx() As String, Index As Long) As Long
    If Index > UBound(mtx) Then
        quitarElemMatriz = -1
        Exit Function
    End If
    
    If (Index = UBound(mtx)) Then
        If (Index > 0) Then
            ReDim Preserve mtx(Index - 1)
        Else
            ReDim mtx(0) 'mejor la dejo asi, suena a menos errores
            'Erase mtx 'la deja en situacion de error !
        End If
    Else
        Dim H As Long
        For H = Index To (UBound(mtx) - 1)
            mtx(H) = mtx(H + 1)
        Next H
        ReDim Preserve mtx(UBound(mtx) - 1)
    End If
    
    
End Function


Private Sub Class_Initialize()
    ReDim mmTMP(0)
End Sub
