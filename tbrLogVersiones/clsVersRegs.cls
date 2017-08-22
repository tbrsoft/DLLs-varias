VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVersRegs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase del sistema
'cuando un programa inicia revisa un archivo de texto donde se indica que version tenia
'cuando inicio la vez anterior. Si la version es diferente agrega un renglon con la fecha y el numero de version
'esto me permitira saber el historia de una PC !!!
'ademas en mi pc oficial habra un historial de versiones existentes y fechas de salida!!! finalmente !!!

'ademas se genera un codigo de pc al iniciar por primera vez
'esta clave unica servira para muchas cosas. Quise usar las claves de la pc pero cuando
'cambian de disco duro o formatean el numero de pc es siempre diferente !!

'como no se graba en la carpeta del sistema (3pm-mprock-etc)y si en la de win cuando se desinstale y se traten de esconder cosas quedara grabado
'al grabarse siempre un mismo archivo se vera cuando haya mas de un programa instalado de tbrsoft

Private sJuamai As String 'who am i, quien soy. Aqui el sistema que lo usa escribe una cadena, por ejemplo 3pm puede llamarse "doble-lana" y mprock "laloca". Un codigo interno por cada programa
Private sQii As String 'key, clave de este sistema para usos varios. Se genera la primera vez
Private sVr As String 'version actual del exe que llama a esta dll

Private fReg As String 'archivo donde se graba el detalle de todas las versiones usadas y la clave particular de este equipo
Private AP As String

Private TRR As New tbrErrores.clsTbrERR
Private FSO As New Scripting.FileSystemObject

Private Sub Class_Initialize()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    TRR.FileLog = AP + "sadq.sadd" 'se agrega al generador de informes todos los "c:\windows\system32\sad.*"
End Sub

'lee todos los datos y deja cargadas las variables.
Public Function OpenAll(sv As String, q As String) As Long
'Los parametros son:
    'la version del sistema que lo esta abriendo (como string = sv)
    'el software que lo esta llamando (parametro q de quien)
    
    On Local Error GoTo ERR33
    
    TRR.Anotar "aa", sv, q
    Dim TxR As TextStream
    
    fReg = AP + "sad." + q + ".sadd" 'va a ser la carpeta de systema por que esta es una dll
    'si 3pm se llama por ejemplo piris el registro estara en sad.piris.ffe
    'hago esto para que a mi en mi pc no se mezclen todos los programas que lo van a usar (mprock incluso tiene uno distinto para cada dueño (martino-e2games))
    
    If FSO.FileExists(fReg) = False Then 'si no existe el archivo es la primera vez que se usa el sistema !
        Set TxR = FSO.CreateTextFile(fReg, True)
            
            '1.- Crear una clave
            Dim TL As Long
            Randomize
            TL = CLng(Rnd * 99999999) + 1
            TxR.Write "qii" + Chr(5) + CStr(TL) + Chr(6)
            
            '2.-escribir para que soft es
            TxR.Write "juamai" + Chr(5) + q + Chr(6)
            
            '3.-escribir que version es
            TxR.Write "vr" + Chr(5) + sv + Chr(6)
            
            '4- escribir la fecha de inicio de todo
            TL = CLng(Date)
            TxR.Write "dt01" + Chr(5) + CStr(TL) + Chr(6)
            
        TxR.Close
        
    End If
    
    TRR.Anotar "ab"
    
    Dim RD As String
    Set TxR = FSO.OpenTextFile(fReg, ForReading)
        RD = TxR.ReadLine
    TxR.Close
    '/////////////////////////////////////
    txtToVars RD 'cargo todos los valores
    '/////////////////////////////////////
    
    'ver si hubo cambio de version
    If sVr <> sv Then 'version leida del archivo vs version que me acaban de avisar que usa
        
        TRR.Anotar "ad", sVr, sv, q, CLng(Date)
        
        Set TxR = FSO.OpenTextFile(fReg, ForAppending)
            TxR.Write "up" + Chr(5) + sv + Chr(5) + q + Chr(5) + CStr(CLng(Date)) + Chr(6)
            'up 3pm-to-7.9.877-39201 'me va a decir que soft se actualizo a que version o si volvio para atras
        TxR.Close
        
        OpenAll = 1 'indica cambio de version! puedo obligarme a explicar por que cambie para tener un registro de 3pm (solo si soy programador)
    Else
        OpenAll = 0 'no ha pasado nada
    End If
    
    Exit Function
ERR33:
    TRR.AppendLog "af", TRR.ErrToTXT(Err)
    Resume Next
End Function

Private Sub txtToVars(T As String) 'leer el archivo cargado y meterlo en las variables
    'el archivo esta separando los renglones con chr6 y cada uno dentro separa nombrevar de valor con chr(5)
    Dim Sp1() As String
    Sp1 = Split(T, Chr(6))
    
    Dim J As Long, SP2() As String
    For J = 0 To UBound(Sp1) - 1
        SP2 = Split(Sp1(J), Chr(5))
        Select Case LCase(SP2(0))
            Case "qii"
                sQii = SP2(1)
            Case "juamai"
                sJuamai = SP2(1)
            Case "vr"
                sVr = SP2(1)
            Case "up" 'actualziacion a otra version
                'si se indica actualizacion sVr debe guardar esta ultima
                sVr = SP2(1)
            Case Else
                'nada que hacer, ver errores ? informarlo ?
        End Select
    
    Next J
    
    TRR.Anotar "ac", sQii, sJuamai, sVr
End Sub

Public Property Get Juamai() As String
    Juamai = sJuamai
End Property

Public Property Let Juamai(J As String)
    sJuamai = J
End Property

Public Property Get Reg() As String
    Reg = fReg
End Property

Public Property Let Reg(J As String)
    fReg = J
End Property

Public Property Get Qii() As String
    Qii = sQii
End Property

