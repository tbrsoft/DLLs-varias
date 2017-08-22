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


Public Function AddFileByPath(path As String, Optional mType As TypeFileSkin = 0) As clsFullImgs_FILE
    Dim A As Long
    A = UBound(mFiles) + 1
    ReDim Preserve mFiles(A)
    mFiles(A).path = path 'se carga el tipo segun extension y el base name
    
    'si quiere meter de pecho el tipo lo hace
    If mType > 0 Then mFiles(A).TypeFile = mType
    
    Set AddFileByPath = mFiles(A)
    
End Function


Public Function AddFileByFImg(sFI As clsFullImgs_FILE) As clsFullImgs_FILE
    Dim A As Long
    A = UBound(mFiles) + 1
    ReDim Preserve mFiles(A)
    Set mFiles(A) = sFI
    
    Set AddFileByFImg = mFiles(A)
    
End Function

Public Function SaveFile(Dest As String) As Long
    'empaquetar todo en un archivo
    'no se graban detalles como reinterpretaciones de tipos de archivos, si fuera importante hacerlo
    
    On Local Error GoTo errUnir
    terr.Anotar "skin008", Dest
    
    Dim JS As New tbrJUSE2.clsJUSE
    JS.Archivo = Dest
    
    Dim A As Long
    For A = 1 To UBound(mFiles)
        terr.Anotar "skin009", mFiles(A).path
        JS.AddFile mFiles(A).path
    Next A
    
    terr.Anotar "skin010"
    JS.Unir False
    
    SaveFile = 0
    
errUnir:
    SaveFile = 1
    
End Function

'los archivos me los pide por el nombre
Public Function getFileByName(sName As String) As clsFullImgs_FILE
    Dim A As Long
    For A = 1 To UBound(mFiles)
        If LCase(mFiles(A).BaseName) = LCase(sName) Then
            Set getFileByName = mFiles(A)
            Exit Function
        End If
    Next A
    
    Set getFileByName = mFiles(0)
End Function

Public Function OpenSkinFile(pthSkin As String, DestFolder As String) As Long
    'abrir un paquete juse en ñla carpeta que se indique
    terr.Anotar "skin003", pthSkin, DestFolder
    
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
    Dim JS As New tbrJUSE2.clsJUSE
    Dim res As Long
    res = JS.ReadFile(pthSkin)
    
    terr.Anotar "skin005", res
    If res = 1 Then
        terr.AppendLog "JusBroke"
        OpenSkinFile = 2
        Exit Function
    End If
    
    terr.Anotar "skin001", JS.CantArchs

    If JS.CantArchs < 3 Then
        terr.AppendLog "NoJusFull"
        OpenSkinFile = 3
        Exit Function
    End If

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
        'ahora agregarlo a este manejador
        AddFileByPath DestFolder + Cada
        
    Next A2

    terr.Anotar "skin006"
    OpenSkinFile = 0
    
End Function

Private Sub Class_Initialize()
    ReDim mFiles(0)
End Sub
