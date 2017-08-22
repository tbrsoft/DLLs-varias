VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSLW"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SaveLic As New clsGuardarLic
Dim SendLic As New clsEnviarLic

Dim WithEvents ReadData As clsReadInData
Attribute ReadData.VB_VarHelpID = -1

Dim WS As Object
Dim qPuerto As Long

'cada licencia se graba en \qPL\ID USUARIO ej. \Licencias\241\licencia.l37
'El usuario de la DLL me da qPL y yo creo un nombre de carpeta
'con el IdUsuario
Dim qPL As String 'Path donde se guardan las licencias


Private Enum CodigoSuceso
    Iniciado = 1
    Terminado = 2
    Conexion = 3
    DatoNoIdentificado = 4
    ArchivoRecibido = 5
    LicenciaActivada = 6
    LicenciaActivadaEnviada = 7
End Enum


Public Event ArchivoRecibido(IdUsuario As Long, PathFile As String)
Public Event Suceso(Suceso As String)

Public Function InicializarPath(PathLicencias As String) As Long
    On Local Error GoTo ProblemasPath
    If Dir(PathLicencias, vbDirectory) = "" Then
        'No existe el directorio
        'Intento crearlo
        MkDir PathLicencias
    End If
    qPL = PathLicencias
    InicializarPath = 0
    Exit Function

ProblemasPath:
    'El directorio no existe NI SE PUEDE CREAR
    InicializarPath = 1
End Function

Public Sub InicializarSocket(WinsockIn As Object)
    Set WS = WinsockIn
    
    Set ReadData = New clsReadInData
    ReadData.Inicia
    
End Sub

Public Sub ComenzarServicio(Puerto As Long)
    qPuerto = Puerto
    WS.Close
    WS.LocalPort = Puerto
    WS.Listen
    
    EnviarSuceso "Servidor Encendido. Puerto:" + CStr(Puerto), Iniciado
End Sub

Public Sub DetenerServicio()
    WS.Close
    EnviarSuceso "Servicio Detenido", Terminado

End Sub

Public Function GetStrEstado() As String
    Dim nEstado As Long
    nEstado = WS.State
    
    Select Case nEstado
        Case 0
            GetStrEstado = "Cerrado"
        Case 1
            GetStrEstado = "Abierto"
        Case 2
            GetStrEstado = "Escuchando"
        Case 3
            GetStrEstado = "Conexión pendiente"
        Case 4
            GetStrEstado = "Resolviendo host"
        Case 5
            GetStrEstado = "Host resuelto"
        Case 6
            GetStrEstado = "Conectando"
        Case 7
            GetStrEstado = "Conectado"
        Case 8
            GetStrEstado = "El equipo está cerrando la conexión"
            'En este caso reinicio la conexion
            WS.Close
            WS.Listen
        Case 9
            GetStrEstado = "Error"
    End Select
End Function

Private Sub EnviarSuceso(Suceso As String, Codigo As CodigoSuceso)
    Dim aux As String
    aux = "[" + CStr(Date) + "]"
    aux = aux + "[" + CStr(Time) + "]"
    aux = aux + "[" + Format(Codigo, "000") + "]"
    aux = aux + ":"
    aux = aux + Suceso
    
    RaiseEvent Suceso(aux)
End Sub

'========================================================
'WinSock
'========================================================
'Private Sub WS_ConnectionRequest(ByVal requestID As Long)
'    WS.Accept requestID
'    EnviarSuceso "Conexion Iniciada con " + CStr(WS.RemoteHostIP), Conexion
'End Sub
'
'Private Sub WS_DataArrival(ByVal bytesTotal As Long)
'    Dim aux As String
'    WS.GetData aux
'    ReadData.ReadData aux
'End Sub

Public Sub ConnectionRequest(ByVal requestID As Long)
    WS.Close
    WS.Accept requestID
    EnviarSuceso "Conexion Iniciada con " + CStr(WS.RemoteHostIP), Conexion
End Sub

Public Sub DataArrival(ByVal bytesTotal As Long)
    Dim aux As String
    WS.GetData aux
    ReadData.ReadData aux
End Sub


Private Sub ReadData_ArchivoEnviado()
    EnviarSuceso "Se envió la Licencia Activada correctamente.", LicenciaActivadaEnviada
End Sub

'========================================================
'Eventos de los datos recibidos
'========================================================

Private Sub ReadData_ArchivoRecibido(IdUsuario As Long, NombreArchivo As String, FileBuffer As String)
    Dim ArchivoLic As String
    'Guardo el Archivo Recibido
    ArchivoLic = SaveLic.GuardarLicencia(qPL, IdUsuario, NombreArchivo, FileBuffer)
    
    'Confirmo al Cliente que el Archivo fue recibido
    WS.SendData Format(tbrSLWP.ArchivoRecibido, "000")
    
    'Aviso al Programa (log)
    EnviarSuceso "Archivo Recibido: " + ArchivoLic, ArchivoRecibido
    RaiseEvent ArchivoRecibido(IdUsuario, ArchivoLic)
End Sub
Public Sub EnviarLicenciaActivada(PathLicencia As String, IdUsuario As Long, NombreArchivo As String)
    
    'Envio la Respuesta
    Dim BufferSend As String
    BufferSend = SendLic.CrearBufferLic(PathLicencia, IdUsuario, NombreArchivo)
    If BufferSend = "-1" Then
        MsgBox "El Archivo de Licencia Activada no se puede Abrir/Leer", vbCritical, "Codigo 5198"
    Else
        'Si no espero anda mal
        Esperar 1500
        'Envio el archivo de Licencia Activada
        
        'Estas lineas no tiene sentido PARA NADA, pero no funciona de otra manera
        WS.SendData "001" + "//" + Mid(BufferSend, 5)
        
        '(esta es la linea que deberia funcionar correctamente
        'WS.SendData BufferSend
    End If
    'Reinicio la conexion
    'WS.Close
    'WS.Listen

End Sub

Public Sub EnviarCreditosInsuficientes()
    'Envia un mensaje de creditos insuficientes
    Esperar 1500
    WS.SendData "003"
End Sub

Private Sub ReadData_DNI(DNI_Buffer As String)
    'Que pasa si envian datos fuera del protocolo?
    EnviarSuceso "Dato No Identificado - Cancelar esta conexion" + vbCrLf + "Datos: " + DNI_Buffer, DatoNoIdentificado
    'Respondo DNI
    WS.SendData "DNI"
    
    'Reinicio la conexion
    WS.Close
    WS.Listen
End Sub

Private Sub Esperar(Milisegundos As Long)
    Dim t1 As Long
    Dim t2 As Long
    t2 = Timer + (Milisegundos / 1000)
    
    While t2 > Timer
        DoEvents
    Wend
End Sub

