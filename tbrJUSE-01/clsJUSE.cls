VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJUSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mLargoFAT As Long 'largo de la cadena FAT
'debe sumarse a cada inicio de archivos para saber el byte de inicio real!!!!
Private mArchivo As String

Private fsoJS As New Scripting.FileSystemObject
Private mArchivosFN() As String 'lista de los archivos incluidos con _
    FullPath para poder UNIR!!!
Private mArchivosFT() As String 'lista de los archivos incluidos con _
    FileTitle solo para dejar grabado no ncesito el path!!!
Private mInis() As Long 'inicio de los archivos en bytes
Private mLargos() As Long 'largos de los archivos en bytes

Public Property Get CantArchs() As Long
    'devuelve la matriz de los archivos incluidos!!
    'lo saco de la matiz de la ultima vez que se unio!!!
    'osea que si acabo de agregar se caga, no esta!
    CantArchs = UBound(mArchivosFT)
End Property

Public Property Get LargoFAT() As Long
    LargoFAT = mLargoFAT
End Property

Public Function GetListFiles(Index As Long, WithPath As Boolean) As String
    If WithPath Then
        GetListFiles = mArchivosFN(Index)
    Else
        GetListFiles = mArchivosFT(Index)
    End If
End Function

Public Sub Quitar(Index As Long)
    'quitar de la lista y NO UNIR
    mArchivosFN(Index) = ""
    mArchivosFT(Index) = ""
End Sub

Public Function AddFile(Arch As String) As Long
    'devuelve 0 si salio todo OK
    'devuelve 1 si no hay archivo a donde agregarlo
    'devuelve 2 si no eligio archivo
    AddFile = 0
    If mArchivo = "" Then
        AddFile = 1 'no hay archivo donde empaquetarlo!
        Exit Function
    End If
    If Arch = "" Then
        AddFile = 2 'no existe el archivo a incluir
        Exit Function
    Else
        If fsoJS.FileExists(Arch) = False Then
            AddFile = 3
            Exit Function
        End If
        
        Dim NewInd As Long
        NewInd = UBound(mArchivosFT) + 1
        ReDim Preserve mArchivosFN(NewInd)
        ReDim Preserve mArchivosFT(NewInd)
        mArchivosFN(NewInd) = Arch
        'filetitle solo guarda el nombre del archivo sin el path
        'que no es necesario para grabar
        mArchivosFT(NewInd) = fsoJS.GetFileName(Arch)
    End If
    'NO UNIR!!!!
    
End Function

Public Sub AddFileShowin()
    Dim CM As New CommonDialog
    Dim F As String
    
    CM.ShowOpen
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    AddFile F
    
End Sub

Private Sub Class_Initialize()
    clearAll
End Sub

Public Sub clearAll()
    mArchivo = ""
    ReDim Preserve mArchivosFT(0)
    ReDim Preserve mArchivosFN(0)
    ReDim Preserve mInis(0)
    ReDim Preserve mLargos(0)
End Sub

Private Sub Class_Terminate()
    Set fsoJS = Nothing
End Sub

Public Function Unir(Optional EsErrorQueExista As Boolean = False) As Long
    
    'EsErrorQueExista es por si se considera error que exista _
        el archivo que se va a unir. Cuando estoy trabjando sobre un mismo _
        archivo es normal que se reemplaze, en otros casos si se considera error
    
    'devuelve 0 si salio todo OK
    'devuelve 1 si no hay archivo a donde agregarlo
    'devuelve 2 si existe el origen y no se puede!!!
    'devuelve 3 por error al unir
    If mArchivo = "" Then
        Unir = 1
        Exit Function
    End If
    
    On Error GoTo FalloMal
    'si es error que exista sale y si no lo es borra el archivo anterior
    If EsErrorQueExista Then
        If Dir(mArchivo) <> "" Then
            Unir = 2
            Exit Function
        Else
            fsoJS.DeleteFile mArchivo, True
        End If
    End If
        
    'y si ponemos una FAT???? (SIN NIVELES DE CARPETAS POR AHORA!!)
    'con eso despues no debería leer de a 100 u otra cantidad buscando
    'donde empiezan y donde terminan los archivos internos
    'además sería mucho mas facil extraer la lista de archivos que contiene!!!
    'El archivo sería:
    '(FAT*) = inicio del archivo (sirve para comprobar)
    '(iARH) = inicio de nombre de archivo
    '(aINI) = inicio de byte de inicio del file (fin de archivo size)
    '(aSZE) = inicio de archivo size en bytes (fin de nombre de archivo)
    '(fARH) = final de la lista de nombres de archivo e inicio de los datos
    '(iFIL) = separador de archivos (ini y fin) no es necesario, _
        solo sirve para comprobar a la hora de leeer un archivo
    'ver entocnes el tamaño de cada uno
    
    Dim W As Long
    Dim Largo As Long 'el largo en bytes
    
    Dim ContArchSi As Long 'archivos que si estan y se uniran
    'no se cuentan los que se borraron o los que no existen
    ContArchSi = 0
    'armar el mapa del archivo!
    'genero una matriz temporal para cargar solo los archivos _
        que quedan en la union
    Dim mtxNew() As String
    ReDim Preserve mtxNew(0)
    For W = 1 To UBound(mArchivosFT)
        'ver que este!!!!
        'cuando se elimina de la lista alguno directamente lo saco de
        'la matriz y no acomodo los indices!!!
        If mArchivosFT(W) = "" Then GoTo sig
        'si ya no existe!!!
        If fsoJS.FileExists(mArchivosFN(W)) = False Then GoTo sig
        'este archivo si esta y vale
        ContArchSi = ContArchSi + 1
        Largo = FileLen(mArchivosFN(W)) 'largo en bytes
        ReDim Preserve mLargos(ContArchSi)
        mLargos(ContArchSi) = Largo
        
        'en la nueva matriz solo quedan los archivos que van
        ReDim Preserve mtxNew(ContArchSi)
        mtxNew(ContArchSi) = mArchivosFN(W)
        'una vez que sepa todos los nombres y los largos de todos los archivos
        'recien voy a saber a donde empieza cada uno
sig:
    Next W
    'ahora ya puedo redefinir la matriz mArchivos. Ahora se que archivos
    'realmenet se van a unir
    'por eso borro la matriz y la cargo correctamente
    ReDim Preserve mArchivosFN(0)
    ReDim Preserve mArchivosFT(0)
    Dim A As Long
    For A = 1 To ContArchSi
        ReDim Preserve mArchivosFN(A)
        ReDim Preserve mArchivosFT(A)
        mArchivosFN(A) = mtxNew(A)
        mArchivosFT(A) = fsoJS.GetFileName(mtxNew(A))
    Next A
    
    'ahora debo conocer los puntos de inicio de cada archivo!
    'una vez oinocideo esto ya puedo empezar a escribir el archivo
    'tener en cuenta los largos de los nombres las etiquetas indicadoras, etc.
            
    '(FAT*) = inicio del archivo (sirve para comprobar)
    '(iARH) = inicio de nombre de archivo
    '(aINI) = inicio de byte de inicio del file (fin de archivo size)
    '(aSZE) = inicio de archivo size en bytes (fin de nombre de archivo)
    '(fARH) = final de la lista de nombres de archivo e inicio de los datos
    '(iFIL) = separador de archivos (ini y fin) no es necesario, _
        solo sirve para comprobar a la hora de leeer un archivo
    
    
    'hacer un dibujo virtual del archivo
    'esta dbujo virtual primero escribe los datos de los archivos
    'estos definen los inicios RELATIVOS (no absolutos) de cada archivo
    'una vez escrita la FAT se sabra cual es el final de la FAT (cual BYTE)
    'y cual el inicio de los archivos
    
    'APROVECHO PARA IR CONTANDO LOS BYTES USADOS PARA CONOCER LAS POSICIONES DE INICIO
    Dim ContBytes As Long
    ContBytes = 0
    'el primer archivo esta en cero siempre
    'los inicios son donde esta el iFIL no el archivo especificamente
    For A = 1 To UBound(mArchivosFT)
        ReDim Preserve mInis(A)
        'queda corrido uno
        mInis(A) = ContBytes
        'no entiendo porque pero es asi!!!
        mInis(A) = mInis(A) + A
        'veo el largo de cada archivo y los 6 de la etiqueta
        ContBytes = ContBytes + 6 + mLargos(A)
    Next A
    'quedaron los inicios definidos!!!
    'una vez que genere la FAT le sumo a cada inicio el total _
        de largo de la FAT!!!!
    Dim FAT As String
    
    FAT = "(FAT*)"
    For A = 1 To UBound(mArchivosFT)
        'nombre del archivo
        FAT = FAT + "(iARH)" + mArchivosFT(A)
        'inicio
        FAT = FAT + "(aINI)" + CStr(mInis(A))
        'tamaño en bytes
        FAT = FAT + "(aSZE)" + CStr(mLargos(A))
    Next A
    'marcar el fin de la lista de archivos
    FAT = FAT + "(fARH)"
    'debo guardar este numero para cuando lea los archivos
    'sepa que a cada inicio debo sumarle esto!!!!!
    mLargoFAT = Len(FAT)
    
    'ahora ya se exactamente que tengo que escribir
    Dim PosToWrite As Long 'Posicion Para escribir
    
    Open mArchivo For Binary As #1
        'escribo todo de una
        Put #1, 1, FAT
        'ya defino la primera posicion
        PosToWrite = mLargoFAT + 1
        Dim fullDATOS As String
        For W = 1 To UBound(mArchivosFT)
            Open mArchivosFN(W) For Binary As #2
                'obtener el archivo
                'para decirle cuanto va a leeer le doy el tamaño en espacios
                fullDATOS = Space(mLargos(W))
                Get #2, 1, fullDATOS
                'poner lo obtenido en el archivo acumulado
                Put #1, PosToWrite, "(iFIL)" + fullDATOS
                PosToWrite = PosToWrite + Len("(iFIL)" + fullDATOS) + 1
            Close #2
        Next W
    Close #1
    Unir = 0
    'MsgBox "OPERACION TERMINADA !", vbInformation, "=)"
    Exit Function
FalloMal:
    Unir = 3
    MsgBox "Se Produjo un error tratando de unir los archivos. " + vbCrLf + _
        Err.Description + " (" + CStr(Err.Number) + ")"
End Function

Public Property Get Archivo() As String
    Archivo = mArchivo
End Property

Public Property Let Archivo(NewArchivo As String)
    mArchivo = NewArchivo
End Property

Public Sub ExecuteOrOpenFile(Arch As String)
    'shellexecute............
End Sub

Public Function ReadFile(Archivo As String) As Long
    'ReadFile= 0 en OK
    'ReadFile= 1 si no es del tipo


    'agarrar un archivo que no tengo puta idea de que se trata
    'cargar la matriz mArchivos, el mLargoFAT, etc, etc, etc
    mArchivo = Archivo
    Open mArchivo For Binary As #1
        'obtener los indicadores de: nombre del archivo - inicio - len in Bytes
        'para eso leo toda la fat
        Dim GetSPC As String
        'si cono cel mLargoFat lo usa.
        'Si recien abro la DLL y busco un archivo hecho ayer no tiene que saberla
        
        'cada archivo ocupa en el indice...
        'aproximado
        '6 por cada indicador = 6x3 =18
        'nombre del archivo         =15
        'byte de inicio             =5
        'bytes de largo             =5
        'Total por archivo          =43
        'aprox de archivos que habra=10
        GetSPC = Space(430)
ReRead:
        Get #1, 1, GetSPC
        'ahora ver si esta el indicador del fin de la FAT (fARH) _
            e inicio del primer archivo (iFIL)
        Dim A As Long, B As Long
        'si algun archivo escribio por dentro "(fARH)(iFIL)" recontra cage!!!
        B = InStr(GetSPC, "(fARH)(iFIL)")
        'si no lo encuentra volver a algo más grande
        If B = 0 Then
            'duplicar
            GetSPC = GetSPC + GetSPC
            'capaz que no eso de los mios!!!
            If Len(GetSPC) > 14000 Then
                ReadFile = 1
                Exit Function
            End If
            GoTo ReRead
        End If
        'ok ya encontro la FAT. Ahora recortar
        GetSPC = Left(GetSPC, B + 5)
        mLargoFAT = B + 5
        
        'empezar a cargar la matriz mArchivos()
        'FAT = FAT + "(iARH)" + mArchivos(A)
        'FAT = FAT + "(aINI)" + CStr(mInis(A))
        'FAT = FAT + "(aSZE)" + CStr(mLargos(A))
        'mas de 10000 archivos seguro que no hay!!!
                
        ReDim Preserve mArchivosFN(0)
        ReDim Preserve mArchivosFT(0)
        ReDim Preserve mInis(0)
        ReDim Preserve mLargos(0)
        
        Dim c As Long, D As Long
        'pongo uno para que mArchivos empieze en uno
        B = 1
        For A = 1 To 10000
            'solamente al cargar el primero es que necesito esto
            If A = 1 Then B = InStr(B, GetSPC, "(iARH)")
            'si ya no hay salir
            If B = 0 Then Exit For
            c = InStr(B, GetSPC, "(aINI)")
            D = InStr(c, GetSPC, "(aSZE)")
            ReDim Preserve mArchivosFT(A)
            'aca el FN se completa para que este pero vacio, ya no se usa
            ReDim Preserve mArchivosFN(A)
            mArchivosFN(A) = ""
            
            B = B + 6
            mArchivosFT(A) = Mid(GetSPC, B, c - B)
            
            c = c + 6
            ReDim Preserve mInis(A)
            mInis(A) = Mid(GetSPC, c, D - c)
            'no entiendo porque pero es asi!!!
            'mInis(A) = mInis(A) + A - 1
            D = D + 6
            'b se tiene que cargar de nuevo
            B = InStr(D, GetSPC, "(iARH)")
            ReDim Preserve mLargos(A)
            'si b es cero no hay más
            If B = 0 Then
                'entonces toma hasta el final quitando los últimos 6 que son (fARH)
                mLargos(A) = Mid(GetSPC, D, (Len(GetSPC) - 5) - D)
                Exit For
            Else
                mLargos(A) = Mid(GetSPC, D, B - D)
            End If
        Next A
        'termino de cargar las matrices
    Close #1
End Function

Public Function Extract(ExtracToFolder As String, IndexFileTo As Long) As Long
    'devuelve 0 si todo OK
    'devuelve 1 si no existe el indice
    'devuelve 2 si no eligio destino o no existe
    If Right(ExtracToFolder, 1) <> "\" Then ExtracToFolder = ExtracToFolder + "\"
    
    'ver si existe el índice
    If UBound(mArchivosFT) < IndexFileTo Then
        Extract = 1
        'MsgBox "No existe el índice que se intenta extraer"
    End If
    
    'si no elige nada salir
    Dim FSO As New Scripting.FileSystemObject
    
    'con las carpetas o unidades de red DIR no anda bien (dentro de una maquina virtual)
    If FSO.FolderExists(ExtracToFolder) = False Then
        Extract = 2
        Exit Function
    End If
    
    'If ExtracToFolder = "" Or Dir(ExtracToFolder, vbDirectory) = "" Then
    '    Extract = 2
    '    Exit Function
    'End If
    
    Dim FR As Long
    FR = FreeFile
    Open mArchivo For Binary As #FR
        'ya tengo todo
        Dim getFileSPC As String
        getFileSPC = Space(mLargos(IndexFileTo))
        'hay que correrse el largo de la fat!!!!!!!!!!!!
        'y hay que corresrse el indicador de 6 caracteres
        Get #FR, mLargoFAT + mInis(IndexFileTo) + 6, getFileSPC
    Close #FR
    'tirarlo en un archivo en la carpeta especificada con el mismo nombre que tenía
    FR = FreeFile
    
    'si existe el destino borrarlo
    If Dir(ExtracToFolder + mArchivosFT(IndexFileTo)) <> "" Then
        Kill ExtracToFolder + mArchivosFT(IndexFileTo)
    End If
    
    Open ExtracToFolder + mArchivosFT(IndexFileTo) For Binary As #FR
        Put #FR, 1, getFileSPC
    Close #FR
        
    Extract = 0
    
End Function

Public Function AddFiles(sFolder As String, Extension As String) As Long
    'devuleve la cantidad de agregados
    Dim FSO As New Scripting.FileSystemObject
    Dim Fl As Scripting.Folder
    Set Fl = FSO.GetFolder(sFolder)
    Dim F2 As Scripting.File
    Dim E1 As String, E2 As String
    For Each F2 In Fl.Files
        E1 = LCase(Right(F2.Name, Len(Extension)))
        E2 = LCase(Extension)
        If E1 = E2 Or Extension = "*" Then 'compatibilidad con comodin "*"
            AddFile F2.path
        End If
    Next
End Function
