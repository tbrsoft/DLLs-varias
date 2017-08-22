VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private FSO As New Scripting.FileSystemObject
Private mLargoAcumula As String 'cantidad de caracteres que acumula
Private mFileLog As String 'archivo donde se registra todo
Private mFileLogGrabaTodo As String 'archivo donde se registra todo
Private mLogAcumulado As String 'acumulado de los lugares por los que paso
Private mMODOGRABATODO As Boolean
Private TextStreamGrabaTodo As TextStream
Private mADN As String 'cadena que puede imprimirse opcionalmente con todos los append
'se usa por ejemplo para saber la version del soft usada y saber si el error es por version vieja

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
Private mHwndMsg As Long 'opcionalmente manda a otro lado todo lo que pasa
'para ir viendolo

Private ContadorRegTodo As Long 'contador de lineas cuando registra todo
'al llegar 50 se borra

Private lastLog As String
Private mMaxBytesLog As Long 'tamaño maximo del log de errores

Public Function GetLastLog() As String
    GetLastLog = lastLog
End Function

Public Function Set_ADN(newADN As String) As String
    mADN = newADN
End Function

'sme es para anuunciar al exterior, writeevent es para acumular privado
Private Function SME(T As String)  'sendMessage Event
    SME = SendMessage(mHwndMsg, WM_SETTEXT, 0, ByVal T)
    'los mensajes van por tipo separados por |!
    '0|
    '1|inquiereFinish
    '2|ConnectionServiceStatus + texto del status
    '3|se mando el archivo ok
    '4|se mando el archivo MAL
    '5|llego un dispositivo + nombre + address
End Function

Public Sub Set_Hwn(n As Long)
    mHwndMsg = n
End Sub

Public Sub Anotar(NuevoTexto As String, Optional ByVal Var1, Optional ByVal Var2, Optional ByVal Var3, Optional ByVal Var4)
    
    On Local Error GoTo AlgunaVarMal
    
    lastLog = NuevoTexto
    
    If Not IsMissing(Var1) Then
        NuevoTexto = NuevoTexto + "." + CStr(Var1)
        If Not IsMissing(Var2) Then
            NuevoTexto = NuevoTexto + "." + CStr(Var2)
            If Not IsMissing(Var3) Then
                NuevoTexto = NuevoTexto + "." + CStr(Var3)
                If Not IsMissing(Var4) Then
                    NuevoTexto = NuevoTexto + "." + CStr(Var4)
                End If
            End If
        End If
    End If
    
    'al hwnd se manda sin el timer
    If mHwndMsg > 0 Then SME NuevoTexto
    'ahora agrego el timer
    NuevoTexto = CStr(Timer) + " " + NuevoTexto
    
    mLogAcumulado = mLogAcumulado + vbCrLf + NuevoTexto
    
    If Len(mLogAcumulado) > mLargoAcumula Then
        mLogAcumulado = Right(mLogAcumulado, mLargoAcumula)
    End If
    
    'si esta en modo que graba todo grabar!
    If mMODOGRABATODO Then
        ContadorRegTodo = ContadorRegTodo + 1
        If ContadorRegTodo > 250 Then
            'solo los ultimos 50
            'cerrar el archivo en curso
            TextStreamGrabaTodo.Close
            'borrarlo
            FSO.DeleteFile mFileLogGrabaTodo, True
            'empezar a grabr de nuevo
            StartGrabaTodo
            ContadorRegTodo = 0
        Else
            TextStreamGrabaTodo.WriteLine NuevoTexto
        End If
        
    End If
       
    Exit Sub
    
AlgunaVarMal:
    'anotarlo como error para que no de error de automatizacion o error en la DLL
    mLogAcumulado = mLogAcumulado + vbCrLf + _
        "****" + vbCrLf + "DLL_ERR_101" + vbCrLf + _
        ErrToTXT(Err)
        
    If Len(mLogAcumulado) > mLargoAcumula Then
        mLogAcumulado = Right(mLogAcumulado, mLargoAcumula)
    End If
    Resume Next
End Sub

Public Sub StartGrabaTodo()
    Set TextStreamGrabaTodo = FSO.OpenTextFile(mFileLogGrabaTodo, ForAppending, True)
End Sub

Public Sub StopGrabaTodo()
    'se cerro correctamente, cerrar el archivo y borrarlo
    TextStreamGrabaTodo.Close
    FSO.DeleteFile mFileLogGrabaTodo, True
End Sub

Public Function ErrToTXT(sERR As ErrObject) As String
    ErrToTXT = CStr(sERR.Number) + ":" + sERR.Description + vbCrLf + sERR.Source
End Function

Public Sub AppendLog(ByVal txtAdicional As String, _
    Optional ByVal txtAdicional2 As String)
    
    'VER EL TAMAÑO DEL ARCHIVO!!!
    If FSO.FileExists(mFileLog) Then
        If FileLen(mFileLog) > mMaxBytesLog Then
            'BORRARLO!!!
            FSO.DeleteFile mFileLog, True
        End If
    End If
    
    'ya tengo el log acumulado, solo necesito el error. Esto puede ser desde
    'un objeto err (con la funcion ErrToTXT o alguna descripciopn mia
    'txtadicional2 se usa cuando hay un err y ademas aclaracion mía
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(mFileLog, ForAppending, True)
        TE.WriteLine vbCrLf + vbCrLf + "*****AL********" + vbCrLf + CStr(Date) + "*" + CStr(Time) + vbCrLf + _
            mADN + vbCrLf + _
            mLogAcumulado + vbCrLf + _
            txtAdicional + vbCrLf + _
            txtAdicional2 + vbCrLf + "*******" + vbCrLf + vbCrLf
'            'si esta dentro de 3PM o mprock lo muestro!!
'            If LCase(mFileLog) = "d:\dev\3pm\reg3pm.log" Then
'                F1.Inicializar mLogAcumulado, txtAdicional + vbCrLf + txtAdicional2
'            End If
    TE.Close
    Set TE = Nothing
    
End Sub

Public Sub AppendSinHist(txt As String)
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(mFileLog, ForAppending, True)
        TE.Write vbCrLf + "****SH****" + CStr(Timer) + ": " + txt + vbCrLf
'            'si esta dentro de 3PM o mprock lo muestro!!
'            If LCase(mFileLog) = "d:\dev\3pm\reg3pm.log" Then
'                F1.Inicializar mLogAcumulado, txtAdicional + vbCrLf + txtAdicional2
'            End If
    TE.Close
    Set TE = Nothing
End Sub

Public Property Let FileLog(NewFile As String)
    mFileLog = NewFile
End Property

Public Property Get FileLog() As String
    FileLog = mFileLog
End Property

Public Property Let FileLogGrabaTodo(NewFile As String)
    mFileLogGrabaTodo = NewFile
End Property

Public Property Get FileLogGrabaTodo() As String
    FileLogGrabaTodo = mFileLogGrabaTodo
End Property

Public Property Let LargoAcumula(Largo As Long)
    mLargoAcumula = Largo
End Property

Public Property Get LargoAcumula() As Long
    LargoAcumula = mLargoAcumula
End Property

Private Sub Class_Initialize()
    mFileLog = App.Path + "\log.txt"
    mLargoAcumula = 430
    mMODOGRABATODO = False
    ContadorRegTodo = 0
    mMaxBytesLog = 200000 'predeterminado 200 kb
    mHwndMsg = 0
End Sub

Public Property Let ModoGrabaTodo(Graba As Boolean)
    mMODOGRABATODO = Graba
End Property

Public Property Get ModoGrabaTodo() As Boolean
    ModoGrabaTodo = mMODOGRABATODO
End Property

Public Property Get LogAcumulado() As String
    LogAcumulado = mLogAcumulado
End Property

Public Function EjecutarEXE(sFile As String, style As VbAppWinStyle)   'necesito probarlo externo
    EjecutarEXE = Shell(sFile, style)
End Function

Public Property Let MaxBytesLog(newVaL As Long)
    mMaxBytesLog = newVaL
End Property

Public Property Get MaxBytesLog() As Long
    MaxBytesLog = mMaxBytesLog
End Property
