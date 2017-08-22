VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuardarLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function GuardarLicencia(PathLicencia As String, IdUsuario As Long, NombreArchivo As String, FileBuffer As String) As String
    Dim PathFinal As String
    Dim Archivo As String
    
    GuardarLicencia = "-1"
    PathFinal = PathLicencia + CStr(IdUsuario) + "\"
    Archivo = PathFinal + NombreArchivo
    
    'Existe PathLicencia?
    If Dir(PathLicencia, vbDirectory) = "" Then
        'si no existe, lo creo
        MkDir PathLicencia
    End If
    
    'existe PathFinal?
    If Dir(PathFinal, vbDirectory) = "" Then
        'si no existe, lo creo
        MkDir PathFinal
    End If
    
    'existe Archivo?
    If Dir(Archivo) <> "" Then
        'si existe ya ese archivo, busco nombre alternativo
        Archivo = GetNombreAlternativo(Archivo)
    End If
    
    Open Archivo For Binary As #1
        Put #1, 1, FileBuffer
    Close #1
    
    'devuelvo el nombre del archivo guardado
    GuardarLicencia = Archivo
End Function

Private Function GetNombreAlternativo(qNombre As String) As String
    Dim Nombre As String
    Dim ext As String
    Dim tmp_NombreAlternativo As String
    
    Dim NumeroAlternativo As Long
    
    Dim aux As Long 'donde esta el primer "." de derecha a izquierda
    aux = InStrRev(qNombre, ".")
    
    'si no tiene "." no lo hago
    If aux > 0 Then
        Nombre = Mid(qNombre, 1, aux - 1)
        ext = Mid(qNombre, aux)
    Else
        Nombre = qNombre
        ext = ""
    End If
    
    NumeroAlternativo = 0
    tmp_NombreAlternativo = ""
    
    'Bucle hasta encontrar un nombre de archivo que no exista
    While Dir(tmp_NombreAlternativo) <> ""
        tmp_NombreAlternativo = Nombre + "_" + Format(NumeroAlternativo, "000") + ext
        NumeroAlternativo = NumeroAlternativo + 1
    Wend
    
    GetNombreAlternativo = tmp_NombreAlternativo
End Function
