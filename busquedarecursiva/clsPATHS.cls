VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPATHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private SubL() As String
Private Tams() As Single 'tamaños de los que son directorios y de los archivos
Private FSO As New Scripting.FileSystemObject
Private tErr As New tbrErrores.clsTbrERR
Public Event Leyendo(accion As String, detalle As String)

Public Function GetLista() As String()
    GetLista = SubL
End Function

Public Function setPathLog(p As String)
    tErr.FileLog = p
End Function

Public Function GetTamanoDirectorioMB(sFOL As String) As Single
    
    tErr.Anotar "aaaa", sFOL
    Dim tamano As Single
    
    'buscar que numero tiene en la lista
    Dim J As Long
    For J = 1 To UBound(SubL)
        RaiseEvent Leyendo("TamDir", "Buscando directorio" + vbCrLf + sFOL + "[" + CStr(J) + "]")
        tErr.Anotar "aaab", J
        If LCase(sFOL) = LCase(SubL(J)) Then
            'si pertenece !
            tamano = Tams(J)
            tErr.Anotar "aaac", tamano
            Exit For
        End If
    Next J
    tErr.Anotar "aaad"
    
    GetTamanoDirectorioMB = Round(tamano / 1048576, 2)
End Function

Public Sub LeerTodo(ByVal Ruta As String, _
    Solodirectorios As Boolean, Solo1Nivel As Boolean, _
    Optional sFilterFiles As String = "*.*")
    
    On Local Error GoTo errPaths
    RaiseEvent Leyendo("LeerTodo", "Leyendo todo" + vbCrLf + Ruta)
    
    tErr.Anotar "aaae", Ruta, Solodirectorios, Solo1Nivel, sFilterFiles
    
    Dim xxx As Long
        
    Dim nombres() As String, I As Long
    ' asegurarse de que existe una barra invertida inicial
    If Right(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    
    'SI ES EL PRIMERO GRABO LA RUTA (puede entrar mas de una vez por que es recursiva)
    If UBound(SubL) = 0 Then SubL(0) = Ruta
    
    tErr.Anotar "aaaf"
    If Solodirectorios = False Then
    '---------------LEE ARCHIVOS-------------------------
        nombres() = ObtenerSoloArchivos(Ruta, sFilterFiles)
        ' cargar los resultados parciales en el cuadro de lista
        For I = 1 To UBound(nombres)
            RaiseEvent Leyendo("LeerTodo", "[" + CStr(I) + "]-" + nombres(I))
            tErr.Anotar "aaag", I
            'DoEvents 'sacado por incompatibilidad con usb 5/11/2010!!!
            xxx = UBound(SubL) + 1
            ReDim Preserve SubL(xxx)
            ReDim Preserve Tams(xxx)
            
            Dim EsteFil As String
            EsteFil = Ruta & nombres(I)
            
            'a veces pasa que no existe !!!!
            SubL(xxx) = EsteFil
            If FSO.FileExists(EsteFil) Then
                Tams(xxx) = FileLen(EsteFil)
                tErr.Anotar "aaah", EsteFil, Tams(xxx)
            Else
                Tams(xxx) = 0 'ES UN ERROR QUE SE DEBE SOLUCIONAR!!!
                tErr.Anotar "aaai"
            End If
        Next I
    End If
    tErr.Anotar "aaaj"
    '---------------LEE CARPETAS-------------------------
    ' obtener la lista de subdirectorios, incluyendo los ocultos
    ' y ejecutar recursivamente esta rutina en todos ellos.
    nombres() = ObtenerSoloSubDirs(Ruta, vbHidden)
    For I = 1 To UBound(nombres)
        RaiseEvent Leyendo("LeerTodo", "[" + CStr(I) + "]-" + nombres(I))
        tErr.Anotar "aaak", I
        'DoEvents 'sacado por incompatibilidad con usb 5/11/2010!!!
        
        xxx = UBound(SubL) + 1
        ReDim Preserve SubL(xxx)
        ReDim Preserve Tams(xxx)
        
        Tams(xxx) = -1 'bandera para que se le cargue su tamaño
        
        SubL(xxx) = Ruta & nombres(I)
        tErr.Anotar "aaal", SubL(xxx)
        'para identificar carpetas
        If Right(SubL(xxx), 1) <> "\" Then SubL(xxx) = SubL(xxx) & "\"
        
        If Solo1Nivel = False Then
            tErr.Anotar "aaam", Solo1Nivel
            LeerTodo Ruta & nombres(I), Solodirectorios, False, sFilterFiles
        End If
    Next
    
    tErr.Anotar "aaan"
    
    Exit Sub
    
errPaths:
    tErr.AppendLog "Error al leer todo" + vbCrLf + tErr.ErrToTXT(Err)
    Resume Next
End Sub

Public Function UpdateFolderSize() As Long
    'ponerle a cada directorio su tamaño
    
    On Local Error GoTo errPaths
    tErr.Anotar "aaao"
    'podría leer todos los tams=-1 pero se complicaria con los subdirectorios
    'mejor leo cada archivo y le sumo a todas las carpetas que le corresponden
    
    Dim H As Long, FS As New Scripting.FileSystemObject
    
    For H = 1 To UBound(Tams)
        
        tErr.Anotar "aaap", H
        'SI ES UN ARCHIVO ==>
        If Right(SubL(H), 1) <> "\" Then
            
            'todos suman al path padre de todo
            Tams(0) = Tams(0) + Tams(H)
            RaiseEvent Leyendo("upFolSize", "[" + CStr(H) + "-" + CStr(Tams(H)) + "]")
            tErr.Anotar "aaaq", Tams(0), Tams(H)
            'ver si de algun modo la contiene
            Dim MaxFolderFile As String 'todo el path hasta la carpeta del archivo
            MaxFolderFile = FS.GetParentFolderName(SubL(H))
            If Right(MaxFolderFile, 1) <> "\" Then MaxFolderFile = MaxFolderFile + "\"
            tErr.Anotar "aaar", SubL(H)
            Dim J As Long
            For J = 1 To UBound(Tams)
                tErr.Anotar "aaas", J
                'SI ES UNA CARPETA ==>
                If Right(SubL(J), 1) = "\" Then
                    If Left(MaxFolderFile, Len(SubL(J))) = SubL(J) Then
                        'si pertenece !
                        Tams(J) = Tams(J) + Tams(H)
                    End If
                End If
                
                'no salgo, puede pertenecer a varias carpetas segun el nivel
            Next J
            tErr.Anotar "aaat"
        End If
    Next H
    tErr.Anotar "aaau"
    
    Exit Function
errPaths:
    tErr.AppendLog "Error al upfolsz" + vbCrLf + tErr.ErrToTXT(Err)
    Resume Next
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta que coincide con los atributos de búsqueda
' opcionalmente, devuelve la ruta completa.

Public Function ObtenerSoloSubDirs(Ruta As String, Optional Atributos As _
    VbFileAttribute, Optional IncluirRuta As Boolean) As String()
    
        On Local Error Resume Next
        tErr.Anotar "aaav", Ruta, IncluirRuta
    
        Dim Resultado() As String
        Dim nombredir As String, contador As Long, ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim Resultado(ALLOC_CHUNK) As String
        ' genera el nombre de ruta + barra invertida
        ruta2 = Ruta
        If Right$(ruta2, 1) <> "\" Then ruta2 = ruta2 & "\"
        nombredir = Dir$(ruta2 & "*.*", vbDirectory Or Atributos)
        
        Do While Len(nombredir)
            tErr.Anotar "aaaw", nombredir
            RaiseEvent Leyendo("getSubDir", nombredir)
            If nombredir = "." Or nombredir = ".." Then  'SEGUIRAQUI no deberia no existir pero pasa !!!!
                ' excluir las entradas "." y ".."
            ElseIf (GetAttr(ruta2 & nombredir) And vbDirectory) = 0 Then
                ' este es un archivo normal
            Else
                ' es un directorio
                contador = contador + 1
                If contador > UBound(Resultado) Then
                    ' cambia el tamaño del array resultante, si
                    ' en necesario
                    ReDim Preserve Resultado(contador + ALLOC_CHUNK) As String
                End If
                ' incluir la ruta si se pide
                If IncluirRuta Then nombredir = ruta2 & nombredir
                Resultado(contador) = nombredir
            End If
            nombredir = Dir$
        Loop
        tErr.Anotar "aaax"
        ' proporciona el array resultante
        ReDim Preserve Resultado(contador) As String
        ObtenerSoloSubDirs = Resultado
        tErr.Anotar "aaay"
End Function

' proporciona un array de cadenas que almacenan todos los nombres de archivo que
' coinciden con una especificación de archivo dada y unos atributos de búsqueda.

Public Function ObtenerSoloArchivos( _
    Ruta As String, _
    Optional sFilterFiles As String = "*.*", _
    Optional Atributos As VbFileAttribute) As String()
    
        On Local Error GoTo errPaths
        tErr.Anotar "aaaz", Ruta, sFilterFiles
    
        If Right(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    
        Dim Resultado() As String
        Dim nombreArchivo As String, contador As Long, ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim Resultado(0 To ALLOC_CHUNK) As String
        nombreArchivo = Dir(Ruta + sFilterFiles, Atributos)
        
        Do While Len(nombreArchivo)
            contador = contador + 1
            RaiseEvent Leyendo("getFiles", "[" + CStr(contador) + "] " + nombreArchivo)
            tErr.Anotar "aaba", contador
            If contador > UBound(Resultado) Then
                ' cambia el tamaño del array resultado, si es necesario
                ReDim Preserve Resultado(0 To contador + ALLOC_CHUNK) As String
            End If
            Resultado(contador) = nombreArchivo
            ' queda preparado para la siguiente iteración
            nombreArchivo = Dir
            tErr.Anotar "aabb", nombreArchivo
        Loop
        
        ' devuelve el array resultado
        ReDim Preserve Resultado(0 To contador) As String
        ObtenerSoloArchivos = Resultado
        tErr.Anotar "aabc"
        
    Exit Function
errPaths:
    tErr.AppendLog "Error al obtSoloAch" + vbCrLf + tErr.ErrToTXT(Err)
    Resume Next
End Function

Private Sub Class_Initialize()
    ReDim SubL(0)
    tErr.FileLog = App.Path + "\logRecursivo.log"
    tErr.LargoAcumula = 800
    Dim Vrs As String
    Vrs = CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    tErr.Set_ADN "tbrPaths v" + Vrs
End Sub
