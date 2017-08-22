VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------------------
' Module    : Crypt Módulo de clase
' DateTime  : 14/03/2004 10:23 -4UTC
' Author    : RadioMan (radioman@cantv.net)
'             es el Autor de la clase que encapsula
'             los métodos criptográficos hechos por Fredrik Qvarfort
'
'
'            Todas la clases, excepto ésta en este proyecto
'            fueron creadas por Fredrik Qvarfort.
'            El archivo original puede descargarse en:
'            http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=12023&lngWId=1
'
'            Según el autor:
'            Esta es una colección de algoritmos SEGUROS de encriptamiento
'            BLOWFISH, TWOFISH, SKIPJACK, TEA, GOST, RC4 y DES
'            Este Post contiene código altamente optimizado y si
'            alguno similr existe en PSC (PlanetSourceCode.com)
'            seguramente éste es al menos 1000x mas rápido.
'
'            Estos no son como todos los otros códigos irrompibles,
'            éstos son algoritmos estandarizados de cifrado
'            donde está la única manera posible de romperlos
'            es por la fuerza bruta
'
'            El DES es probablemente el menos seguro
'            hay que recordar que tomó 22 horas para una
'            supercomputadora + 100.000 computadoras para
'            romperlo (distributed.net)). Y es uno de los mas seguros jajaja
'
'
' Purpose :  Esta clase encapsula a todos los métodos criptográficos
'            contenidos en las diferentes clases las cuales son privadas
'            y esta expone todos sus métodos comunes y permite de una
'            forma sencilla seleccionar el método criptográfico
'            en un momento dado
'
'            '
'            '
' Declaration :
'             '
'             '
' Revision History:
'              Creada el Sabado 13/03/2004 y terminada el Domingo 14/03/2004
'              Agregado Base64 19/03/04
'
'---------------------------------------------------------------------------------------
Option Explicit


'---------------------------
'      EVENTOS PUBLICOS
'---------------------------
' El evento Progress indica el procentaje de avance
' cuando el objeto (Clase o método criptográfico)
' Se encuentra ejecutando alguna operacion
'Public Event Progress(Percent As Long)


'---------------------------
'  ENUMERACIONES PUBLICAS
'---------------------------
' Aquí se enumeran los métodos criptográficos posibles que
' se usarán en cualquiera de las operaciones
Public Enum eMetodoCriptografico
    eMC_Blowfish
    eMC_CryptAPI
    eMC_DES
    eMC_Gost
    eMC_XOR
    eMC_RC4
    eMC_Skipjack
    eMC_TEA
    eMC_Twofish
End Enum

Public Enum eBase64EncodeDecode
    eB64_Encode
    eB64_Decode
End Enum



'---------------------------
'    VARIABLES PRIVADAS
'---------------------------
' Clase Encode/Decode para Base64
Private m_cBase64 As clsB64


' Éste objeto es declarado genéricamente (Object) y puede asumir en cualquier
' momento la identidad de cualquiere de los métodos criptográficos
' existente
Private m_oObjetoCriptografico As Object

'Es el número total de métodos de encriptamiento definidos
' (se usa con EncryptObjects())
Private EncryptObjectsCount As Long

'---------------------------------------------------------------------------------------
' Procedure : Base64 Function
' DateTime  : 19/03/2004 14:06 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Function Base64String(str As String, EncodeDecode As eBase64EncodeDecode) As String
    Dim arrIn() As Byte, arrOut() As Byte
    
    Set m_cBase64 = New clsB64
    
    With m_cBase64
        .Init
        .StrToAry str, arrOut()
        If EncodeDecode = eB64_Encode Then
            .Encode arrOut(), arrIn()
        ElseIf EncodeDecode = eB64_Decode Then
            .Decode arrOut(), arrIn()
        End If
        
        .AryToStr arrIn(), str
    End With
    
    Set m_cBase64 = Nothing
    
    'Erase arrIn(), arrOut()
    
    Base64String = str
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : Base64File Function
' DateTime  : 19/03/2004 14:21 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Function Base64File(SourceFile As String, DestFile As String, EncodeDecode As eBase64EncodeDecode) As Boolean
    Dim arrIn() As Byte, arrOut() As Byte
    
    On Local Error Resume Next
    
    Set m_cBase64 = New clsB64
    
    With m_cBase64
        .Init
        .Load SourceFile, arrIn()
        If EncodeDecode = eB64_Encode Then
            .Encode arrIn(), arrOut()
        ElseIf EncodeDecode = eB64_Decode Then
            .Decode arrIn(), arrOut()
        End If
        .Save arrOut(), DestFile
    End With
    
    Set m_cBase64 = Nothing
    
    If Err.Number = 0 Then Base64File = True
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : DecryptBlock Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Private Static Sub DecryptBlock(m As eMetodoCriptografico, Xl As Long, Xr As Long)
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    Call m_oObjetoCriptografico.DecryptBlock(Xl, Xr)
    '
End Sub
'---------------------------------------------------------------------------------------
' Procedure : EncryptBlock Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Private Static Sub EncryptBlock(m As eMetodoCriptografico, Xl As Long, Xr As Long)
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    Call m_oObjetoCriptografico.EncryptBlock(Xl, Xr)
    '
End Sub
'---------------------------------------------------------------------------------------
' Procedure : EncryptByte Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Sub EncryptByte(m As eMetodoCriptografico, ByteArray() As Byte, Optional Key As String)
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        Call m_oObjetoCriptografico.EncryptByte(ByteArray())
    Else
        Call m_oObjetoCriptografico.EncryptByte(ByteArray(), Key)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : EncryptString Function
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Function EncryptString(m As eMetodoCriptografico, Text As String, Optional Key As String, Optional ConvertBase64 As Boolean = False) As String
    Dim str As String
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        str = m_oObjetoCriptografico.EncryptString(Text)
    Else
        str = m_oObjetoCriptografico.EncryptString(Text, Key)
    End If
    
    If ConvertBase64 Then
        EncryptString = Base64String(str, eB64_Encode)
    Else
        EncryptString = str
    End If
    
    
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : DecryptString Function
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Function DecryptString(m As eMetodoCriptografico, Text As String, Optional Key As String, Optional ConvertBase64 As Boolean = False) As String
    Dim str As String
    
    Inicializar m
    
    'antes de decodificar debe reconvertir desde Base64
    If ConvertBase64 Then
        str = Base64String(Text, eB64_Decode)
    Else
        str = Text
    End If
    
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        DecryptString = m_oObjetoCriptografico.DecryptString(str)
    Else
        DecryptString = m_oObjetoCriptografico.DecryptString(str, Key)
    End If
    
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : EncryptFile Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Sub EncryptFile(m As eMetodoCriptografico, SourceFile As String, _
    DestFile As String, Optional Key As String, Optional ConvertBase64 As Boolean = False)
    
    On Local Error GoTo EEFil
    
    TERR.Anotar "aac"
    Inicializar m
    
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        TERR.Anotar "aad"
        Call m_oObjetoCriptografico.EncryptFile(SourceFile, DestFile)
    Else
        TERR.Anotar "aae"
        Call m_oObjetoCriptografico.EncryptFile(SourceFile, DestFile, Key)
    End If
    
    If ConvertBase64 Then
        TERR.Anotar "aaf"
        Call Base64File(DestFile, DestFile, eB64_Encode)
    End If
    
    TERR.Anotar "aag"
    
    Exit Sub
    
EEFil:
    TERR.AppendLog "saaa"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DecryptFile Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Sub DecryptFile(m As eMetodoCriptografico, SourceFile As String, DestFile As String, Optional Key As String, Optional ConvertBase64 As Boolean = False)
    
    ' Reconvierta desde BAse64 si es indicado
    If ConvertBase64 Then
        Call Base64File(SourceFile, SourceFile, eB64_Decode)
    End If
    
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        Call m_oObjetoCriptografico.DecryptFile(SourceFile, DestFile)
    Else
        Call m_oObjetoCriptografico.DecryptFile(SourceFile, DestFile, Key)
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DecryptByte Sub
' DateTime  : 19/03/2004 14:07 -4UTC
' Purpose   :
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Public Sub DecryptByte(m As eMetodoCriptografico, ByteArray() As Byte, Optional Key As String)
    
    Inicializar m
    
    ' ejecute el método de la clase correspondiente
    If Key = "" Then
        Call m_oObjetoCriptografico.DecryptByte(ByteArray())
    Else
        Call m_oObjetoCriptografico.DecryptByte(ByteArray(), Key)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Inicializar Function
' DateTime  : 14/03/2004 10:03 -4UTC
' Purpose   : Sirve para inicializar correctamente el objeto
'             que servirá para efectuar la operación criptográfica
' Arguments :
' Returns   :
' Example   :
' Revision History:
'
'---------------------------------------------------------------------------------------
'
Private Sub Inicializar(ByVal m As eMetodoCriptografico)
    
    'inicializa la clase correspondiente al método criptográfico
    Select Case m
        
    Case eMC_Blowfish
        TERR.Anotar "aah"
        Set m_oObjetoCriptografico = New clsBlowfish
        
    Case eMC_CryptAPI
        TERR.Anotar "aai"
        Set m_oObjetoCriptografico = New clsCryptAPI
        
    Case eMC_DES
        TERR.Anotar "aaj"
        Set m_oObjetoCriptografico = New clsDES
        
    Case eMC_Gost
        TERR.Anotar "aak"
        Set m_oObjetoCriptografico = New clsGost
        
    Case eMC_XOR
        TERR.Anotar "aal"
        Set m_oObjetoCriptografico = New clsSimpleXOR
        
    Case eMC_RC4
        TERR.Anotar "aam"
        Set m_oObjetoCriptografico = New clsRC4
        
    Case eMC_Skipjack
        TERR.Anotar "aan"
        Set m_oObjetoCriptografico = New clsSkipjack
        
    Case eMC_TEA
        TERR.Anotar "aao"
        Set m_oObjetoCriptografico = New clsTEA
        
    Case eMC_Twofish
        TERR.Anotar "aap"
        Set m_oObjetoCriptografico = New clsTwofish
        
    Case Else
        
    End Select
    
    TERR.Anotar "aaq"
    
End Sub

Public Sub SetLog(sDir As String)
    TERR.FileLog = sDir
    TERR.Anotar "aab"
End Sub

Private Sub Class_Initialize()
    TERR.FileLog = App.Path + "\regCR.log"
    TERR.Set_ADN CStr(App.Major * 100000 + App.Minor * 1000 + App.Revision)
    
    TERR.LargoAcumula = 600
    TERR.Anotar "aaa"
End Sub
