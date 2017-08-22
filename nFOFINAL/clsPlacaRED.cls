VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPlacaRED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&


Private Type MIB_IFROW
    wszName(0 To 511) As Byte
    dwIndex As Long             '// index of the interface
    dwType As Long              '// type of interface
    dwMtu As Long               '// max transmission unit
    dwSpeed As Long             '// speed of the interface
    dwPhysAddrLen As Long       '// length of physical address
    bPhysAddr(0 To 7) As Byte   '// physical address of adapter
    dwAdminStatus As Long       '// administrative status
    dwOperStatus As Long        '// operational status
    dwLastChange As Long        '// last time operational status changed
    dwInOctets As Long          '// octets received
    dwInUcastPkts As Long       '// unicast packets received
    dwInNUcastPkts As Long      '// non-unicast packets received
    dwInDiscards As Long        '// received packets discarded
    dwInErrors As Long          '// erroneous packets received
    dwInUnknownProtos As Long   '// unknown protocol packets received
    dwOutOctets As Long         '// octets sent
    dwOutUcastPkts As Long      '// unicast packets sent
    dwOutNUcastPkts As Long     '// non-unicast packets sent
    dwOutDiscards As Long       '// outgoing packets discarded
    dwOutErrors As Long         '// erroneous packets sent
    dwOutQLen As Long           '// output queue length
    dwDescrLen As Long          '// length of bDescr member
    bDescr(0 To 255) As Byte    '// interface description
End Type

Private Declare Function GetIfTable Lib "iphlpapi" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Private m_InterfaceType As Long
Private m_InterfaceTypeString As String
Private m_AdapterAddress As String
Private m_AdapterAddress2 As String
Private m_CantidadPlacasRed As Long
Private m_InterfaceDescription As String
Private m_ArrayDatosDePlacas() As Byte
Private m_NamePlaca As String

Public Property Get NamePlaca() As String
    NamePlaca = m_NamePlaca
End Property

Public Property Get InterfaceDescription() As String
    InterfaceDescription = m_InterfaceDescription
End Property

Public Property Get AdapterAddress() As String
    AdapterAddress = m_AdapterAddress
End Property

Public Property Get AdapterAddress2() As String
    AdapterAddress2 = m_AdapterAddress2
End Property

Public Property Get InterfaceTypeString() As String
    InterfaceTypeString = m_InterfaceTypeString
End Property

Public Property Get InterfaceType() As Long
    InterfaceType = m_InterfaceType
End Property

Public Property Get CantidadPlacasRed() As Long
    CantidadPlacasRed = m_CantidadPlacasRed
End Property

Public Function ObtenerDatosPlaca(IndicePlaca As Long)
    Dim IfRowTable As MIB_IFROW
    'Copy the table row data to the IfRowTable structure
    CopyMemory IfRowTable, m_ArrayDatosDePlacas(4 + (IndicePlaca - 1) * Len(IfRowTable)), Len(IfRowTable)
    'se cargaron todas las propiedades mágicamente ¿¿??
    With IfRowTable
        Dim XX As Long
        m_NamePlaca = ""
        For XX = 0 To 511
            If .wszName(XX) > 0 Then
                m_NamePlaca = m_NamePlaca + Chr$(.wszName(XX))
            End If
        Next XX
        'el ultimo -1 no estaba, no se que es pero no permite concatenar m_InterfaceDescription
        'con nada mas desde el FRM, por esos se lo quito
        m_InterfaceDescription = Left(StrConv(.bDescr, vbUnicode), .dwDescrLen - 1)
        'dwPhysAddrLen es la cantidad de partes del MAC ADDRESS
        m_AdapterAddress = ""
        m_AdapterAddress2 = ""
        If .dwPhysAddrLen > 0 Then
            Dim J As Integer
            For J = 0 To .dwPhysAddrLen - 1
                'ver si es uno o dos CHR, relleno con ceros todo lo que falte para 2 lugares
                m_AdapterAddress = m_AdapterAddress + String(2 - Len(CStr(Hex(.bPhysAddr(J)))), "0") + CStr(Hex(.bPhysAddr(J)))
                m_AdapterAddress2 = m_AdapterAddress2 + String(2 - Len(CStr(Hex(.bPhysAddr(J)))), "0") + CStr(Hex(.bPhysAddr(J)))
                'poner el puto guion
                If J < .dwPhysAddrLen - 1 Then m_AdapterAddress = m_AdapterAddress + "-"
            Next J
        Else
            m_AdapterAddress = "NO"
        End If
        'cargar el tipo de interfase
        m_InterfaceType = .dwType
        
        Select Case m_InterfaceType
            Case 1
                m_InterfaceTypeString = "1 (OTHER)"
            Case 6
                m_InterfaceTypeString = "6 (ETHERNET)"
            Case 9
                m_InterfaceTypeString = "9 (TOKENRING)"
            Case 15
                m_InterfaceTypeString = "15 (FDDI)"
            Case 23
                m_InterfaceTypeString = "23 (PPP)"
            Case 24
                m_InterfaceTypeString = "24 (LOOPBACK)"
            Case 28
                m_InterfaceTypeString = "28 (SLIP)"
        End Select
    End With
End Function

Public Function Inicializar() As Long
    'devuelve -2 si no e compatible
    'devuelve -1 si no se puede ver los datos ¿¿??

    Dim TamanoGetIfTable As Long
    TamanoGetIfTable = 0
    '
    'Call the GetIfTable just to get the buffer size into the lngSize variable
    'llama la funcion solo para saber el tamaño de la matriz de Bytes que cargara _
        en la segunda llamada
    Dim Ret As Long
    Ret = GetIfTable(ByVal 0&, TamanoGetIfTable, 0)
    '
    If Ret = ERROR_NOT_SUPPORTED Then
        Inicializar = -2 'error debe ser Windows 95 o algo viejo
        'This API works only on Win 98/2000 and NT4 with SP4
        MsgBox "IP Helper is not supported by this system."
        Exit Function
        '
    End If

    'Redimensiono el buffer con la canbtidad de datos
    ReDim m_ArrayDatosDePlacas(0 To TamanoGetIfTable - 1) As Byte
    '
    'llama a la funciona de vuelta ya conociendo el tamaño necesario
    'carga el array de bytes con toda la información
    Ret = GetIfTable(m_ArrayDatosDePlacas(0), TamanoGetIfTable, 0)
    
    If Ret = ERROR_SUCCESS Then
        'Los primeros cuatro Bytes (valor LONG) contienen la cantidad de filas en la tabla
        'lngrows es la cantidad de placas encontradas
        Dim lngRows As Long
        CopyMemory lngRows, m_ArrayDatosDePlacas(0), 4
        'DESTINO,SOURCE, LArgo del source
        'COPY MEMORY pone en un lugar de memoria lo que le indico
        'en este caso pone en lngRows solo el primer indice de la matriz de bytes que
        'devolvio GetIfTable
        m_CantidadPlacasRed = lngRows
        Inicializar = 0
    Else
        Inicializar = -1 'error
    End If
End Function
