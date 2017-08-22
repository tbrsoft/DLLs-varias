VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPERM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'permisos que tendra este pendrive
'SEGUIRAQUI estaria bueno que se pueda tambien ejecutar el pendrive sin que
'el sistema este funcionando y se recaude informacion solo de los archivos encontrados sueltos
'serviria para casos como el pacha que se rompen equipos y no se accede al sistema


Public Type TPerm
    mIdNodo As Long
    mValPerm As Long
End Type

Private PermsCollec() As TPerm

Public Function AddPerm(nodop As Long, valp As Long) As TPerm
    Dim F As Long
    F = UBound(PermsCollec) + 1
    ReDim Preserve PermsCollec(F)
    PermsCollec(F).mIdNodo = nodop
    PermsCollec(F).mValPerm = valp
    AddPerm = PermsCollec(F)
End Function

Public Function GetCantPerms() As Long
    GetCantPerms = UBound(PermsCollec)
End Function

Public Function getPerm(index As Long) As TPerm
    getPerm = PermsCollec(index)
End Function

Public Function GetPermByIndexNodo(iNodo As Long) As Long
    Dim h As Long
    For h = 1 To GetCantPerms
        If PermsCollec(h).mIdNodo = iNodo Then
            GetPermByIndexNodo = PermsCollec(h).mValPerm
            Exit Function
        End If
    Next h
    
    GetPermByIndexNodo = -1 'no se encontro
End Function

Public Function toString() As String
    Dim h As Long, t As String
    t = ""
    For h = 1 To GetCantPerms
        t = t + CStr(PermsCollec(h).mIdNodo) + ":" + CStr(PermsCollec(h).mValPerm)
        If h < GetCantPerms Then t = t + "|"
    Next h
    
    toString = t
End Function

Public Sub clear()
    ReDim PermsCollec(0)
End Sub

Private Sub Class_Initialize()
    clear
End Sub


'///////////////////////////////////////////////////
''///////////////////////////////////////////////////
''anterior limitado solo a una config !!!!
''///////////////////////////////////////////////////
''///////////////////////////////////////////////////
'
'
'Private mUPDTMSC As Long 'puede actualziar musica?
''0 si no puede hacer nada
''1 si puede actualziar
''2 puede actualizar y definir destinos
'Public Enum enUpdtMSC
'    NoPuede = 0
'    SoloActualiza = 1
'    DefineOrigenes = 2
'End Enum
'
'Private mVerCNTs As Long 'puede ver los contadores ?
''0 no ve nada de los contadores
''1 ve los contadores (y el historico de variaciones)
''2 puede reiniciar el contador reiniciable
''3 puede hacer lo que quiera con los contadores
'Public Enum enContadores
'    NoVeContadores = 0
'    SoloLosVe = 1
'    SoloReiniContR = 2
'    FULLVariasContadores = 3
'End Enum
'
'Public Property Get UpdtMsc() As enUpdtMSC
'    UpdtMsc = mUPDTMSC
'End Property
'
'Public Property Let UpdtMsc(ByVal vNewValue As enUpdtMSC)
'    mUPDTMSC = vNewValue
'End Property
'
'Public Property Get VerCnts() As enContadores
'    VerCnts = mVerCNTs
'End Property
'
'Public Property Let VerCnts(ByVal vNewValue As enContadores)
'    mVerCNTs = vNewValue
'End Property
'
'Public Function toString() As String
'    toString = CStr(mUPDTMSC) + CStr(mVerCNTs)
'End Function
'
