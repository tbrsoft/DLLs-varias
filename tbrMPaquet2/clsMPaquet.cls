VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMPaquet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase de paquete de musica mas conocido como disco
'el objetivo es que esta clase guarde infortmacion de cada disco

'fecha en que se cargo al sistema
'fecha de primera y ultima vez que se escucho
'cantidad de veces que se vio (entro a el)
'canciones mas escuchadas y vendidas
'etc etc etc

'todo esto va a un archivo de texto


Private dtActivacion As Long 'fecha de inicio del disco en long
Private dtUltimoUso As Long 'fecha de inicio se escucho/compro por ultima

Private CtInis As Long ' cantidad de veces que se inicio este disco. Es un contador de veces que el disco estaba cuando se inicio
Private CtInsides As Long 'cantidad de veces que se entro a ver el disco
Private CtLisen As Long 'cantidad de veces que se ejecuto canciones de este disco
'ctLisen/CtInsedes muy bajo puede significar que el contenido no es lo que se espera
Private CtSell As Long 'catidad de veces que se vendieron archivos de este disco

Public Enum ModoInEnum
    DejaloComoEsta = 0 'quiere decir que deje el que estaba de antes
    EstabaDeAntes = 1 'Lo debe haber copiado por que el sistema lo encontro al iniciar
    IngresoPorPenDriveAutomatico = 2 'Se ingreso por el pendrive automatizado (se puede usar para bloqueos tipo martino)
    AsistenteBasico = 3 'Uso el asistente para cargar musica
End Enum

Private ModoIn As ModoInEnum 'modo en el que este disco se ingreso al sistema
'esto me permitira mostrar solo los discos insertados de x modo que es lo que martino quiere


' la funciona grabar requiere el path cuado es nuevo,
' cuando se graba uno recien abierto que se modifico voy a llamar al ultimo
Private lastFileOpen As String

Private mPath As String 'path de esta carpeta

'SOLO SE CARGA EL PATH EN EL CHECK AND CREATE POR AHORA ME ASEGURA PROLIJIDAD
Public Function GetPath() As String
    GetPath = mPath
End Function

'el archivo debe ser
'ID-INI=3298749283
'dtactivacion=39811
'dtUltimoUso=39822
'ctInis=32
'CTInsides=41
'CtLisen = 9
'CtSell=0
'ModoIn=1

'revisa que el archivo exista en la carpeta y que sea para la pc actual

Public Function CheckAndCreate(ByVal pth As String, ByVal ForModoIn As ModoInEnum) As Long
    'devuelve cero para avisar que ya existe o 1 si se tuvo que crear
    If Right(pth, 1) <> "\" Then pth = pth + "\"
    Dim TxR As TextStream
    
    resetAll 'para que no se mezcle con datos analizado de otro disco. Se usara un objeto para todos los discos de la fonola
    mPath = pth 'para que se acuerde
    lastFileOpen = pth + "3h.dt"
    If FSo.FileExists(lastFileOpen) Then
        'ver que sea para esta pc
        Dim A As Long
        A = LoadFile(lastFileOpen, ForModoIn)
        CheckAndCreate = A
        '0=se cargo y encotro ok (puede haber cambiado ModoIn, hay que grabar)
        '1=se grabo como nuevo porque era de otra pc
    Else
        'crearlo como nuevo
        Grabar True, lastFileOpen, ForModoIn
        CheckAndCreate = 2
    End If
End Function

'leer el archivo, buscar el id de esta pc y cargarlo en las variables que correspondan
Private Function LoadFile(ByVal pth As String, ByVal ForModoIn As ModoInEnum)

    On Local Error GoTo errLF

    Dim TxR As TextStream, R As String
    Dim estoyEnIdOk As Boolean 'cada vez que se leen un id-ini se revisa esto
    'de esta forma solo se cargaran los datos correctos
    estoyEnIdOk = False
    
    Dim SP() As String
    Set TxR = FSo.OpenTextFile(pth, ForReading)
        Do While Not TxR.AtEndOfStream
            R = TxR.ReadLine 'la separacion de los renglones es vbcrlf
            SP = Split(R, Chr(5)) 'e internamente chr5
            '-------------------------------------------
            Select Case LCase(SP(0))
                Case "id-ini"
                    If SP(1) = idEstaPC Then
                        estoyEnIdOk = True
                    Else
                        estoyEnIdOk = False
                    End If
                Case "dtactivacion": If estoyEnIdOk Then dtActivacion = CLng(SP(1))
                Case "dtultimouso": If estoyEnIdOk Then dtUltimoUso = CLng(SP(1))
                Case "ctinis": If estoyEnIdOk Then CtInis = CLng(SP(1))
                Case "ctinsides": If estoyEnIdOk Then CtInsides = CLng(SP(1))
                Case "ctlisen": If estoyEnIdOk Then CtLisen = CLng(SP(1))
                Case "ctsell": If estoyEnIdOk Then CtSell = CLng(SP(1))
                Case "modoin"
                    If estoyEnIdOk Then
                        If ForModoIn = DejaloComoEsta Then ''quiere decir que deje el que esta!
                            ModoIn = CLng(SP(1))
                        Else 'quiere que le ponga otro 'por ejemplo validar que ingreo ok para martino
                            ModoIn = ForModoIn
                        End If
                    End If
                Case Else 'ver que pasa !!
                    TRR.AppendSinHist "aasa:" + R
                    
            End Select
        Loop
    TxR.Close

    If estoyEnIdOk Then
        LoadFile = 0 'se cargo y encontro ok
    Else
        'crear todo desde cero!
        LoadFile = 1
        Grabar True, pth, ForModoIn
    End If

    Exit Function
errLF:
    TeRR.AppendLog "aasb", TRR.ErrToTXT(Err)
    Resume Next
End Function

'crear el archivo cuando no existe (tambien lo uso para grabar existentes)
Private Sub Grabar(ByVal ISNEW As Boolean, ByVal pth As String, ByVal ForModoIn As ModoInEnum)
    
    On Local Error GoTo errGraba
    
    TRR.Anotar "sasb", ISNEW, pth, idEstaPC, ForModoIn
    
    'GRABAR EL QUE RECIEN SE ABRIO
    If LCase(pth) = "last" Then pth = lastFileOpen
    'si me tiraron ua carpeta era un 3H.DT!!!
    If Right(pth, 1) = "\" Then pth = pth + "3h.dt"
    
    Dim R As String
    Dim TxR As TextStream
    If ISNEW Then
        
        'si puso el cero para respetar lo que habia y no habia nada entonces
        'le pongo 1 que es que lo encontro "tirado" en el disco duro
        If ForModoIn = DejaloComoEsta Then ForModoIn = EstabaDeAntes
    
        R = "ID-INI" + Chr(5) + idEstaPC + vbCrLf + _
            "dtActivacion" + Chr(5) + CStr(CLng(Date)) + vbCrLf + _
            "dtUltimoUso" + Chr(5) + CStr(CLng(Date)) + vbCrLf + _
            "ctInis" + Chr(5) + "1" + vbCrLf + _
            "CTInsides" + Chr(5) + "0" + vbCrLf + _
            "CtLisen" + Chr(5) + "0" + vbCrLf + _
            "CtSell" + Chr(5) + "0" + vbCrLf + _
            "ModoIn" + Chr(5) + CStr(ForModoIn)
                    'tener en cuenta que ultimo uso se considera tambien su ingreso ya que si no tiene uso debe haber de toda formas una referencia
                    
        'cargar los valores
        dtActivacion = CStr(CLng(Date))
        dtUltimoUso = CStr(CLng(Date))
        CtInis = 1
        CtInsides = 0
        CtLisen = 0
        CtSell = 0
        ModoIn = ForModoIn
        
    Else
        R = "ID-INI" + Chr(5) + idEstaPC + vbCrLf + _
            "dtActivacion" + Chr(5) + CStr(dtActivacion) + vbCrLf + _
            "dtUltimoUso" + Chr(5) + CStr(dtUltimoUso) + vbCrLf + _
            "ctInis" + Chr(5) + CStr(CtInis) + vbCrLf + _
            "CTInsides" + Chr(5) + CStr(CtInsides) + vbCrLf + _
            "CtLisen" + Chr(5) + CStr(CtLisen) + vbCrLf + _
            "CtSell" + Chr(5) + CStr(CtSell) + vbCrLf + _
            "ModoIn" + Chr(5) + CStr(ModoIn)
    End If
    TRR.Anotar "sasc", idEstaPC, dtActivacion, dtUltimoUso, ModoIn
    
    Set TxR = FSo.CreateTextFile(pth, True) 'sobreescribe si ya existia (caso en que load encontro un archivo de otra pc)
        TxR.Write R
    TxR.Close
    
    Exit Sub
errGraba:
    TRR.AppendLog "sasa", TRR.ErrToTXT(Err)
    Resume Next
End Sub
'-------------------------------------------------------------

Private Sub Class_Initialize()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    TRR.FileLog = AP + "sadre.sadd" 'se agrega al generador de informes todos los "c:\windows\system32\sad.*"
End Sub

'--------------------------------------------
'fecha de alta, o sea la primera vez que se leyo el archivo
'solo se lee ya que la primera vez que se genera el archivos se carga alli y no se debe cambiarla
Public Property Get FechaAlta() As Long
    FechaAlta = dtActivacion
End Property

'--------------------------------------------
'fecha ultimo uso, la ultima vez que se gasto plata aca (escuchar o comprar)
'se actualiza cada vez que se usa el disco
Public Property Get FechaUltimoUso() As Long
    FechaUltimoUso = dtUltimoUso
End Property

Public Property Let FechaUltimoUso(newF As Long)
    dtUltimoUso = newF
End Property

'--------------------------------------------
Public Property Get CantidadEncendidas() As Long
    CantidadEncendidas = CtInis
End Property
'cantidadencendidas++
Public Sub ppEncencidas()
    CtInis = CtInis + 1
End Sub

'--------------------------------------------
Public Property Get CantidadEntradas() As Long
    CantidadEntradas = CtInsides
End Property
' ++
Public Sub ppEntradas()
    CtInsides = CtInsides + 1
End Sub

'--------------------------------------------
Public Property Get CantidadPlayed() As Long
    CantidadPlayed = CtLisen
End Property
' ++
Public Sub ppLisen()
    CtLisen = CtLisen + 1
End Sub

'--------------------------------------------
Public Property Get CantidadVendido() As Long
    CantidadVendido = CtSell
End Property
' ++
Public Sub ppSell()
    CtSell = CtSell + 1
End Sub

'--------------------------------------------
Public Property Get ModoAltaDisco() As ModoInEnum
    ModoAltaDisco = ModoIn
End Property

Private Sub resetAll()
    dtActivacion = 0
    dtUltimoUso = 0
    CtInis = 0
    CtInsides = 0
    CtLisen = 0
    CtSell = 0
    ModoIn = 0
End Sub
