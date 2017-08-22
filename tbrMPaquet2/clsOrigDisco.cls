VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrigDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'maneja un origen de discos
'debe por lo tanto ser a la vez un mPaquet manager (todavia no lo es)


Private mPath As String 'carpeta con \ del origen
Private FSo As New Scripting.FileSystemObject

Private mFoldersToAdd() As String 'cada carpeta que se desea agregar al origen
Private mMBaCopiar As Long 'mientras se va listando carpetas a copiar en este origen vamos contando cuantop pede agregar
Private ixUni As Long 'indice en la matriz de unidades de OrigsManager

Public Event CopyngPorc(Porc As Single, descr As String)
Private mPorcCopiyng As Single 'solo cuando esta copiando

'nuevo version 2, antes no era un manager de discos
Private MPs() As clsMPaquet

Public Function AddMPaquet(sFolder As String, Optional ByVal ModoIn As ModoInEnum = 0)
    Dim k As Long
    k = UBound(MPs) + 1
    ReDim Preserve MPs(k)
    MPs(k).CheckAndCreate sFolder, ModoIn
End Function

Public Function AddFullMPaquets()
    'seguiraqui
    'agregar todas las carpetas en este origen que cumplan con lo de tipo multimedia
    
End Function

Public Function GetmPaquetById(sID As Long) As clsMPaquet
    If sID = 0 Or sID > UBound(MPs) Then
        GetmPaquetById = MPs(0) 'vacio!
    Else
        GetmPaquetById = MPs(sID)
    End If
End Function

Public Property Get Path() As String
    Path = mPath
End Property

Public Function PathBaseName() As String
    getPathBaseName = FSo.GetBaseName(mPath)
End Function

Public Property Let Path(newP As String)
    mPath = newP
    If Right(mPath, 1) <> "\" Then mPath = mPath + "\"
End Property

Public Function GetMBaCopiar() As Long
    GetMBaCopiar = mMBaCopiar
End Function

Public Function AddFolderToCopy(newF As String)
    Dim m As Long
    m = UBound(mFoldersToAdd) + 1
    ReDim Preserve mFoldersToAdd(m)
    If Right(newF, 1) <> "\" Then newF = newF + "\"
    mFoldersToAdd(m) = newF
    
    mMBaCopiar = mMBaCopiar + CLng((FSo.GetFolder(newF).Size / 1024) / 1024)
End Function

Public Function GetDriveLetter() As String
    GetDriveLetter = FSo.GetDriveName(FSo.GetAbsolutePathName(mPath))
End Function

Public Function GetPorcCopy() As Single
    GetPorcCopy = mPorcCopiyng
End Function

Public Sub StartCopy() 'copiar sin preguntar demasiado ni revisar los 3H.DT, es una opcion
    mPorcCopiyng = 0
    Dim k As Long, des As String
    RaiseEvent CopyngPorc(1, "Iniciando copia") 'para que se note que empezo
    For k = 1 To UBound(mFoldersToAdd)
        myCopyFolder mFoldersToAdd(k), mPath, True 'EN COPYFOLDER NO DEBE HABER BARRAS AL FINAL!!!
        des = FSo.GetBaseName(mfoldertoadd) + " -> " + FSo.GetBaseName(mPath)
        mPorcCopiyng = Round(k / UBound(mFoldersToAdd), 2)
        RaiseEvent CopyngPorc(mPorcCopiyng, des)
    Next k
    mPorcCopiyng = 1
End Sub

'en vez de copiar de a chorro mando una lista de las cosas que hay que copiar carpeta por carpeta
'esto permite mejores barras de progreso y modificar/crear los 3H.DT
Public Function GetListFolderToCopy() As String
    Dim resFINAL As String
    resFINAL = ""
    Dim k As Long
    
    For k = 1 To UBound(mFoldersToAdd) 'cada folder to add es un origen del pendrive, debo verle a cada uno todas sus carpetas
    
        Dim F1 As Folder, F2 As Folder
        Set F1 = FSo.GetFolder(mFoldersToAdd(k))
        For Each F2 In F1.SubFolders
            resFINAL = resFINAL + F2.Path + Chr(5) + mPath + F2.Name + Chr(6) 'NO QUITAR EL CHR6 del final ya que sirve de separardor al usar esta funcion y acoplarse con otros resultados
            'por ejemplo "f:\musica\rock\la renga" (5) "d:\mm\rock\"+"la renga"
        Next
        
    Next k
    GetListFolderToCopy = resFINAL
End Function

Private Sub Class_Initialize()
    mMBaCopiar = 0
    ReDim mFoldersToAdd(0)
    mPorcCopiyng = 0
    ReDim MPs(0)
End Sub

Public Property Get iUni() As Long
    iUni = ixUni
End Property

Public Property Let iUni(newUni As Long)
    ixUni = newUni
End Property
