VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnviarLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Esta funcion crea el buffer de los datos a enviar
'cuando se tranfiere la licencia, desde el servidor hasta el cliente.

'Si devuelve -1 es por que no se pudo abrir/encontrar el archivo
Public Function CrearBufferLic(PathLicencia As String, IdUsuario As Long, NombreArchivo As String) As String
    Dim BufferFinal As String
    Dim aux As String
    Dim Largo_archivo As Long
    
    If Dir(PathLicencia) = "" Then
        CrearBufferLic = "-1"
        Exit Function
    End If
    Largo_archivo = FileLen(PathLicencia)
    aux = Space(Largo_archivo)
    
    Open PathLicencia For Binary As #1
        Get #1, 1, aux
    Close #1
    
    Dim Separador As String
    Separador = "//"
    
    BufferFinal = "001" + Separador
    BufferFinal = BufferFinal + CStr(IdUsuario) + Separador
    BufferFinal = BufferFinal + NombreArchivo + Separador
    BufferFinal = BufferFinal + CStr(Largo_archivo) + "::"
    BufferFinal = BufferFinal + aux
    
    CrearBufferLic = BufferFinal
End Function

