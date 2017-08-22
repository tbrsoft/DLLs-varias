VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadInData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Estructura de un Mensaje de Envio de Licencia
'//001//IdUsuario//NombreArchivo//LargoArchivo::BufferArchivoTotal
Private Type DatosRecib
    IdUsuario As Long
    LargoArchivo As Long
    NombreArchivo As String
End Type


Dim Buffer As String
Dim qDatos As DatosRecib
Dim Recibiendo As Boolean



Public Event ArchivoRecibido(IdUsuario As Long, NombreArchivo As String, FileBuffer As String)
Public Event ArchivoEnviado()
'DNI = Datos No Identificados
Public Event DNI(DNI_Buffer As String)

Public Sub Inicia()
    Recibiendo = False
    Buffer = ""
End Sub

Public Sub ReadData(qData As String)
    Dim aux As String
    
    'Estoy en medio de un tramite de recibir un archivo?
    If Recibiendo = False Then
        
        'Si no estoy recibiendo un archivo leo el protocolo
        'los 3 primeros bytes son el codigo del protocolo
        aux = Mid(qData, 1, 3)
        Protocolo = Val(aux)
        Select Case Protocolo
            Case tbrSLWP.EnvioDeArchivo
                
                Recibiendo = True
                Buffer = ""
                
                'Aqui se carga IdUsuario, LargoArchivo
                'Y LA PRIMER PARTE DEL BUFFER DEL ARCHIVO
                CargarDatos qData
                'Si el archivo es muy pequeño llega en una sola tirada
                CheckArchivoCompleto

            Case tbrSLWP.ArchivoRecibido
                RaiseEvent ArchivoEnviado
                
            Case Else
                'Me llegaron DatosNoIdentificados. O sea, algo fuera del protocolo
                Buffer = qData
                RaiseEvent DNI(Buffer)
        End Select
    
    Else
        'Si el archivo es "grande" llega en una en varias tiradas de ReadData
        
        'Estoy recibiendo un archivo
        Buffer = Buffer + qData
        'Ya envio el archivo completo?
        CheckArchivoCompleto
    End If
End Sub

'Ya llego el largo del Archivo que se envia?
Private Sub CheckArchivoCompleto()
    'Es mayor el largo del buffer que estoy cargando que
    'el largo de archivo que me dijo el 'Cliente'
    If Len(Buffer) >= qDatos.LargoArchivo Then
        Recibiendo = False
        
        '(!) Es Posible que me llegue un Buffer mas grande de lo esperado
        'por eso lo corto
        Buffer = Mid(Buffer, 1, qDatos.LargoArchivo)
        
        'El archivo se recibio exitosamente
        RaiseEvent ArchivoRecibido(qDatos.IdUsuario, qDatos.NombreArchivo, Buffer)
    End If
End Sub

Private Function CargarDatos(DataIn As String)
    Dim aux As String
    Dim str() As String
    'cargo los datos desde el byte 1 hasta ::
    aux = Mid(DataIn, 1, (InStr(DataIn, "::")))
    'los separo
    str = Split(aux, "//")
    
    qDatos.IdUsuario = Val(str(1))
    qDatos.NombreArchivo = str(2)
    qDatos.LargoArchivo = Val(str(3))
    
    'cargo los datos de :: en adelante
    Buffer = Mid(DataIn, (InStr(DataIn, "::") + 2))
End Function
