VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFullImgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'empaquetadorm de archivos de uso general
'cada clase del manu se graba como un paquete ya que puede tener cada uno mas de un archivos
'asi mismo el skin general como la suma de todos estos paquetes tambien es un paquete de este tipo


Private mFiles() As clsFullImgs_FILE  'todos los arhcivos usados
Dim refFILE As String 'referencia entre nosmbres de archivos e ids
    
Public Function GetFileCantidad() As Long
    GetFileCantidad = UBound(mFiles)
End Function
Public Function AddFileByPath(path As String, Optional mType As TypeFileSkin = 0) As clsFullImgs_FILE
    
    terr.Anotar "skin012", path, mType
    
    Dim A As Long
    A = UBound(mFiles) + 1
    ReDim Preserve mFiles(A)
    mFiles(A).path = path 'se carga el tipo segun extension y el base name
    
    'si quiere meter de pecho el tipo lo hace
    If mType > 0 Then mFiles(A).TypeFile = mType
    
    Set AddFileByPath = mFiles(A)
    
End Function

Public Function AddFileByFImg(sFI As clsFullImgs_FILE) As clsFullImgs_FILE
    
    terr.Anotar "skin013", sFI.path
    
    Dim A As Long
    A = UBound(mFiles) + 1
    ReDim Preserve mFiles(A)
    Set mFiles(A) = sFI
    
    Set AddFileByFImg = mFiles(A)
    
End Function

Public Function SaveFile(Dest As String) As Long
    'empaquetar todo en un archivo
    
    On Local Error GoTo errUnir
    terr.Anotar "skin008", Dest
    
    Dim JS As New tbrjuse2.clsJUSE
    JS.Archivo = Dest
    
    Dim A As Long
    For A = 1 To UBound(mFiles)
        terr.Anotar "skin009", mFiles(A).path
        JS.AddFile mFiles(A).path
    Next A
    
    'agregar el archivo de registros de cada archivo y su id
    CreateRefFile
    
    JS.AddFile refFILE
    
    terr.Anotar "skin010"
    JS.Unir False
    
    'eliminar el archivo de registro para que no se pise con los proximos
    fso.DeleteFile refFILE, True
    
    
    SaveFile = 0
    
errUnir:
    SaveFile = 1
    
End Function

'crear el archivo de referencia de cada archivo
Private Function CreateRefFile() As Long
    
    terr.Anotar "skin014"
    
    'si grabo un vinculo sID-->archivo
    Dim refIDs As String 'lista de los archivos y los IDs
    refIDs = ""
    
    Dim A As Long
    For A = 1 To UBound(mFiles)
        terr.Anotar "skin023", mFiles(A).path
        
        'registrar los archivos y sus IDs
        refIDs = refIDs + mFiles(A).BaseName + ">" + mFiles(A).sID + ">" + CStr(mFiles(A).TypeFile) 'los tipos los guardo por si hay algun .TXT que quiero que sea imagen por ejemplo
        If A < UBound(mFiles) Then refIDs = refIDs + "|"
    Next A
    
    terr.Anotar "skin015", A
    
    'agregar el archivo de registros de cada archivo y su id
    Dim TE As TextStream
    Set TE = fso.CreateTextFile(refFILE, True)
        TE.Write refIDs
    TE.Close
    
    terr.Anotar "skin016"
    CreateRefFile = 0
    
End Function

Private Function loadRefFile(sFlie As String)

    terr.Anotar "skin017", sFile

    Dim TE As TextStream, TX As String
    Set TE = fso.OpenTextFile(sFile, ForReading, False)
        TX = TE.ReadAll
    TE.Close
    
    Dim SP() As String, J As Long
    SP = Split(TX, "|")
    
    terr.Anotar "skin018"
    
    Dim SP2() As String
    For J = 0 To UBound(SP)
        
        SP2 = Split(SP(J), ">")
        terr.Anotar "skin019", J, SP2(0)
        
        Dim F As clsFullImgs_FILE
        Set F = GetFileByName(SP2(0))
        F.sID = SP2(1)
        F.TypeFile = CLng(SP2(2))
        
    Next J
    
    terr.Anotar "skin020"
    
End Function

Public Function OpenSkinFile(pthSkin As String, DestFolder As String) As Long

    On Local Error GoTo ErrOP

    'abrir un paquete juse en la carpeta que se indique
    
    terr.Anotar "skin003", pthSkin, DestFolder
    Clear 'empezar de cero los archivos
    
    If Right(DestFolder, 1) <> "\" Then DestFolder = DestFolder + "\"

    If fso.FileExists(pthSkin) = False Then
        terr.AppendLog "NoJus"
        OpenSkinFile = 1
        Exit Function
    End If

    'SEGUIRAQUI darle seguridad a este archivo
    'desencriptarlo
    'Dim CAES As New tbrCrypto.Crypt
    'terr.Anotar "qao181"
    'CAES.DecryptFile eMC_Blowfish, jsFile, AP + "Config.tmp", "Inserte moneda"
    'terr.Anotar "qao182"
    'If fso.FileExists(jsFile) = False Then 'saber si se desencripto ok
    '    terr.AppendLog "NoJus2"
    '    End
    'End If
    
    terr.Anotar "skin004"
    OpenSkinFile = 2
    Dim JS As New tbrjuse2.clsJUSE
    Dim res As Long
    res = JS.ReadFile(pthSkin)
    
    terr.Anotar "skin005", res
    If res = 1 Then
        terr.AppendLog "JusBroke"
        OpenSkinFile = 2
        Exit Function
    End If
    
    terr.Anotar "skin001", JS.CantArchs

    OpenSkinFile = 3
    Dim Cada As String, A2 As Long
    For A2 = 1 To JS.CantArchs

        Cada = JS.GetListFiles(A2, False)
        terr.Anotar "skin002", A2, Cada

        If Cada = "" Then
            terr.AppendLog "FailJUS"
            OpenSkinFile = 4
            Exit Function
        End If

        JS.Extract DestFolder, A2
        
        Dim esteArch As String, esteRef As String
        esteArch = DestFolder + Cada
        
        'el archivo de referencia no se debe agregar, solo usarlo para cargar todos los IDs
        If fso.GetBaseName(refFILE) = fso.GetBaseName(esteArch) Then
            'anotar para usar al final (este no se agrega ya que se crea solo al grabar)
            esteRef = esteArch
        Else
            'ahora agregarlo a este manejador
            AddFileByPath esteArch
        End If
        
    Next A2

    terr.Anotar "skin022"
    OpenSkinFile = 4
    'cargar el refFile
    loadRefFile esteRef

    terr.Anotar "skin006"
    OpenSkinFile = 0
    
    Exit Function
    
ErrOP:
    OpenSkinFile = 1
    terr.AppendLog "errOP_SKN", terr.ErrToTXT(Err)
    
End Function

Private Sub Class_Initialize()
    Clear
    refFILE = tmpFolder + "RefFILE.txt"
End Sub

Public Sub Clear()
    ReDim mFiles(0)
End Sub

Public Function GetFileBysID(searchID As String) As clsFullImgs_FILE
    Dim J As Long
    For J = 1 To UBound(mFiles)
        If LCase(mFiles(J).sID) = LCase(searchID) Then
            GetFileBysID = mFiles(J)
            Exit Function
        End If
    
    Next J
    
    'error!
    terr.AppendLog "NoFileBySiD", searchID
    GetFileBysID = mFiles(0)
End Function

Public Function GetFileById(Ix As Long) As clsFullImgs_FILE
    Set GetFileById = mFiles(Ix)
End Function

'los archivos me los pide por el nombre
Public Function GetFileByName(sName As String) As clsFullImgs_FILE
    
    terr.Anotar "skin023", sName
    Dim A As Long
    For A = 1 To UBound(mFiles)
        'se puede usar el "*" al final para buscar archivos que "empiezan por ..."
        If Right(sName, 1) = "*" Then
            If LCase(Left(mFiles(A).BaseName, Len(sName) - 1)) = LCase(Left(sName, Len(sName) - 1)) Then
                Set GetFileByName = mFiles(A)
                Exit Function
            End If
        Else
            If LCase(mFiles(A).BaseName) = LCase(sName) Then
                Set GetFileByName = mFiles(A)
                Exit Function
            End If
        End If
    Next A
    
    terr.AppendLog "skin024"
    Set GetFileByName = mFiles(0)
End Function


'el objFGull crea el FIMG y algunas clases hacen el suyo propio con las imagenes que usan
'0no para grabar sino para entregarle al oficial que crea objFull
Public Sub AddFilesFromOtherFIMG(otherFIMG As clsFullImgs)
    
    terr.Anotar "skin025"
    Dim J As Long
    For J = 1 To otherFIMG.GetFileCantidad
        terr.Anotar "skin026", J
        Dim F As clsFullImgs_FILE
        Set F = otherFIMG.GetFileById(J)
        AddFileByFImg F
    
    Next J
    
    terr.Anotar "skin027"
    
End Sub
