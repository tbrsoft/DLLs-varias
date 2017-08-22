VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinalID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim MACs() As String
Dim DescrPlaca() As String
Dim TipoPlaca() As String
Dim NombrePlaca() As String
Dim nTipoPlaca() As Long


Private Sub Class_Initialize()
       
    On Local Error GoTo ErrFI
    
    Terr.Anotar "taa"
    Dim PLACA As New ClsPlacaRED
    'cargar los valores de cantidad de placas y los datos de estas sin traducir
    Dim Ret As Long
    Ret = PLACA.Inicializar
    Terr.Anotar "tab", Ret
    ReDim MACs(0)
    
    If Ret = -1 Then
        MACs(0) = "Error desconocido. No se pueden obtener los datos"
        Exit Sub
    End If
    If Ret = -2 Then
        MACs(0) = "Su version de Windows no es compatible, use Windows 98/98SE/Me/Xp/NT/2000"
        Exit Sub
    End If
    
    'agrega las placas encontradas
    
    Dim ConMac As Long, A As Long
    ConMac = 0
    MACs(0) = "" 'si es <>"" hubo un error
    Terr.Anotar "tac", PLACA.CantidadPlacasRed
    For A = 1 To PLACA.CantidadPlacasRed
        Terr.Anotar "tad", A, PLACA.AdapterAddress
        PLACA.ObtenerDatosPlaca A
        'ver solo placas que tengan MAC
        If Len(PLACA.AdapterAddress) > 2 Then 'puede ser "NO"(=2)
            ConMac = ConMac + 1
            ReDim Preserve MACs(ConMac)
            MACs(ConMac) = PLACA.AdapterAddress2
            Terr.Anotar "tae", MACs(ConMac)
            ReDim Preserve DescrPlaca(ConMac)
            DescrPlaca(ConMac) = PLACA.InterfaceDescription
            Terr.Anotar "taF", DescrPlaca(ConMac)
            ReDim Preserve TipoPlaca(ConMac)
            TipoPlaca(ConMac) = PLACA.InterfaceTypeString
            Terr.Anotar "tag", TipoPlaca(ConMac)
            ReDim Preserve nTipoPlaca(ConMac)
            nTipoPlaca(ConMac) = PLACA.InterfaceType
            Terr.Anotar "tah", nTipoPlaca(ConMac)
            ReDim Preserve NombrePlaca(ConMac)
            NombrePlaca(ConMac) = PLACA.NamePlaca
            Terr.Anotar "tai", NombrePlaca(ConMac)
        End If
    Next

    Terr.Anotar "taj"
    
    Exit Sub
    
ErrFI:
    Terr.AppendLog "ErrFi_INI", Terr.ErrToTXT(Err)
End Sub

Public Property Get CantidadMac()
    CantidadMac = UBound(MACs)
End Property

Public Function GetRealMac(i As Long) As String
    If i > UBound(MACs) Then
        GetRealMac = "MAC no existente"
    Else
        GetRealMac = MACs(i)
    End If
End Function

Public Function GetDescripcion(i As Long) As String
    If i > UBound(MACs) Then
        GetDescripcion = "MAC no existente"
    Else
        GetDescripcion = DescrPlaca(i)
    End If
End Function

Public Function GetTipo(i As Long) As String
    If i > UBound(MACs) Then
        GetTipo = "MAC no existente"
    Else
        GetTipo = TipoPlaca(i)
    End If
End Function

Public Function GetnTipo(i As Long) As Long
    If i > UBound(MACs) Then
        GetnTipo = -1
    Else
        GetnTipo = nTipoPlaca(i)
    End If
End Function


Public Function GetNombre(i As Long) As String
    If i > UBound(MACs) Then
        GetNombre = "MAC no existente"
    Else
        GetNombre = NombrePlaca(i)
    End If
End Function
