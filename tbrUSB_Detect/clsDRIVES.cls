VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDRIVES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iSubclass

'**************************************************************
'MIO
'**************************************************************
'todo sobre el dispositivo USB conectado
Private mDevSel As Long 'de los que hay el elegido
'**************************************************************

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
Alias "GetLogicalDriveStringsA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String _
) As Long

Private m_clsSubcls As cSubclass

Private mDRIVES() As String
Private udtInfo() As DEVICE_INFORMATION

'para comparar con anteriores guardo copia
Private LastDrives() As String
Private LastUdtInfo() As DEVICE_INFORMATION

'especificos segun necesidad
Public Event IngresaDrive(LetterUnit As String, SBT As Long)
Public Event SaleDrive(LetterUnit As String, SBT As Long)

'si quiero limitar el funcionamiento puedo reducirlo solo a USB!!
Private mSoloDispositivosUSB As Boolean

'cada vez que reciba una señal se la mando a alguien!
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
Private evHwnd As Long

Public Function UseEventMSG(useHWND As Long)
    evHwnd = useHWND
End Function

'sme es para anuunciar al exterior, writeevent es para acumular privado
Private Function SME(T As String)  'sendMessage Event
    SME = SendMessage(evHwnd, WM_SETTEXT, 0, ByVal T)
    'los mensajes van por tipo separados por |!
    '0|Letra        Ingresa unidad
    '1|Letra        Sale unidad
    '2|info         mensajes para acumular para el error seguiraqui no implementado
    '3|error        error exacto seguiraqui no implementado
End Function


Public Property Get SoloDispositivosUSB() As Boolean
    SoloDispositivosUSB = mSoloDispositivosUSB
End Property

Public Property Let SoloDispositivosUSB(NewDefUSB As Boolean)
    mSoloDispositivosUSB = NewDefUSB
End Property

Public Sub RefreshDriveList()

    Dim strDriveBuffer  As String
    Dim strDrives()     As String
    Dim i               As Long
    Dim i_USB           As Long
    
    i = 0
    i_USB = 0
    
    ReDim mDRIVES(0)
    ReDim udtInfo(0)
        
    strDriveBuffer = Space(240)
    strDriveBuffer = _
        Left$(strDriveBuffer, GetLogicalDriveStrings(Len(strDriveBuffer), strDriveBuffer))
    strDrives = Split(strDriveBuffer, Chr$(0))

    Dim TMP_mDRIVES() As String
    Dim TMP_udtInfo() As DEVICE_INFORMATION

    For i = 1 To UBound(strDrives)
        
        'cargo todos al temporal ...
        
        ReDim Preserve TMP_mDRIVES(i)
        TMP_mDRIVES(i) = strDrives(i)
        ReDim Preserve TMP_udtInfo(i)
        TMP_udtInfo(i) = GetDevInfo(strDrives(i))
        
        'y veo si pongo todos o solo los USB ...
        
        'SOLO PARA DISPOSITIVOS USB
        If mSoloDispositivosUSB Then
            If TMP_udtInfo(i).Valid Then
                If TMP_udtInfo(i).BusType = BusTypeUsb Then
                
                    'ver si no es algun  conector para poner un dispositivo!!!
                    'me doy cuenta por que dan cero MB
                    If TMP_udtInfo(i).TotalSizeMB > 0 Then
                        i_USB = i_USB + 1
                        ReDim Preserve mDRIVES(i_USB)
                        mDRIVES(i_USB) = TMP_mDRIVES(i)
                        ReDim Preserve udtInfo(i_USB)
                        udtInfo(i_USB) = TMP_udtInfo(i)
                    End If
                End If
            End If
        Else 'AQUI ENTRAN TODOS LOS DISPOSITIVOS
            ReDim Preserve mDRIVES(i)
            mDRIVES(i) = TMP_mDRIVES(i)
            ReDim Preserve udtInfo(i)
            udtInfo(i) = TMP_udtInfo(i)
        End If
    Next
    
    'ver los cambios para lanzar eventos
    RevisarCambios
    
    LastDrives = mDRIVES
    LastUdtInfo = udtInfo
    
End Sub

Private Sub RevisarCambios()
    'tengo por un lado lastdrives y lastudt y info y los nuevos
    
    'si es la primera vez es como que todos se agregan
    'If UBound(LastDrives) = 0 Then Exit Sub
    
    Dim H As Long
    For H = 1 To UBound(mDRIVES)
        'ver si esta estaba antes, si no estaba lanzar evento de que entro uno
        Dim J As Long, ENCONTRO As Boolean
        ENCONTRO = False
        For J = 1 To UBound(LastDrives)
            If LCase(mDRIVES(H)) = LCase(LastDrives(J)) Then
                ENCONTRO = True
                Exit For
            End If
        Next J
        If ENCONTRO = False Then
            RaiseEvent IngresaDrive(mDRIVES(H), CLng(udtInfo(H).BusType))
            If evHwnd <> 0 Then SME "0|" + Left(CStr(mDRIVES(H)), 1)
        End If
    Next H
    
    'ahora alreves para ver cuales no estan mas
    
    For H = 1 To UBound(LastDrives)
        'ver si esta estaba antes, si no estaba lanzar evento de que entro uno
        ENCONTRO = False
        For J = 1 To UBound(mDRIVES)
            If LCase(LastDrives(H)) = LCase(mDRIVES(J)) Then
                ENCONTRO = True
                Exit For
            End If
        Next J
        If ENCONTRO = False Then
            RaiseEvent SaleDrive(LastDrives(H), CLng(LastUdtInfo(H).BusType))
            If evHwnd <> 0 Then SME "1|" + Left(CStr(LastDrives(H)), 1)
        End If
    Next H
    
End Sub

Public Sub Iniciar(UseForm)
    Set m_clsSubcls = New cSubclass
    
    m_clsSubcls.Subclass UseForm.hwnd, Me
    m_clsSubcls.AddMsg UseForm.hwnd, WM_DEVICECHANGE
    
    RefreshDriveList
End Sub

Public Sub Terminar()
        'recomienda hacerlo en el unload del formulario y no en el terminate del mismo el auto _
        por eso no esta automatico en el terminate
    m_clsSubcls.Terminate
End Sub

Private Sub Class_Initialize()
    mDevSel = 0
    
    ReDim mDRIVES(0)
    ReDim udtInfo(0)
    
    ReDim LastDrives(0)
    ReDim LastUdtInfo(0)
    
    evHwnd = 0
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, _
    lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, _
    ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    
    If uMsg = WM_DEVICECHANGE Then
        RefreshDriveList
    End If
    
End Sub

Public Function EJECT(i As Long) As Long
    If EjectDevice(mDRIVES(i)) Then
        EJECT = 0
        'MsgBox "Successfully ejected the device from the system!", vbInformation
    Else
        EJECT = 1
        'MsgBox "Could not eject " & lvwDrives.SelectedItem.Tag & "!", vbExclamation
    End If
End Function

Public Function GetDriveList(i As Long) As String
    If i > UBound(mDRIVES) Then
        GetDriveList = "OUT"
    Else
        GetDriveList = mDRIVES(i)
    End If
End Function

Public Function GetDriveInfo(i As Long) As String
    If i > UBound(mDRIVES) Then
        GetDriveInfo = "OUT"
    Else
        Dim TMP As String
        TMP = ""
        If udtInfo(i).Valid Then
            Select Case udtInfo(i).BusType
                Case BusTypeUsb:        TMP = "USB"
                Case BusType1394:       TMP = "1394"
                Case BusTypeAta:        TMP = "ATA"
                Case BusTypeAtapi:      TMP = "ATAPI"
                Case BusTypeFibre:      TMP = "Fibre"
                Case BusTypeRAID:       TMP = "RAID"
                Case BusTypeScsi:       TMP = "SCSI"
                Case BusTypeSsa:        TMP = "SSA"
                Case BusTypeUnknown:    TMP = "Unknown"
            End Select
            TMP = TMP + " | " + udtInfo(i).VolumeName
            TMP = TMP + " | " + IIf(udtInfo(i).Removable, "Removable", "no Removable")
            TMP = TMP + " | " + Trim$(udtInfo(i).VendorID & " | " & _
                udtInfo(i).ProductID & " | " & _
                udtInfo(i).ProductRevision) & " | " & _
                udtInfo(i).SerialNumber
            
            TMP = TMP + " | TotalSize " + CStr(udtInfo(i).TotalSizeMB) + " MB" + " | Free Space " + _
                CStr(udtInfo(i).FreeSpaceMB) + " MB"
        
            GetDriveInfo = TMP
        End If
    End If
End Function

'**************************************************************
'MIO
'**************************************************************

Public Property Get DevSel() As Long
    DevSel = mDevSel
End Property

Public Property Let DevSel(NewDevSel As Long)
    mDevSel = NewDevSel 'sera el valor predeterminado cuando no se indique
End Property

Public Function GetCantidadUSB() As Long
    GetCantidadUSB = UBound(mDRIVES)
End Function

Public Function CanSave(cantMB As Long, iDev As Long) As Boolean
    'me dice si puedo o no grabar una determinada cantidad de MB
    If iDev = -1 Then
        CanSave = (cantMB < GetFreeMB(mDevSel))
    Else
        CanSave = (cantMB < GetFreeMB(iDev))
    End If
    
End Function

Public Function RefreshValues(iDev As Long) As Long
    If iDev = -1 Then
        RefreshSpacesMB mDRIVES(mDevSel), _
            udtInfo(mDevSel).TotalSizeMB, _
            udtInfo(mDevSel).FreeSpaceMB, _
            udtInfo(mDevSel).VolumeName
    Else
        RefreshSpacesMB mDRIVES(iDev), _
            udtInfo(iDev).TotalSizeMB, _
            udtInfo(iDev).FreeSpaceMB, _
            udtInfo(iDev).VolumeName
    End If
End Function

Public Function GetLetterUSB(i As Long) As String
    If i > GetCantidadUSB Then Exit Function
    If i = -1 Then
        GetLetterUSB = Left(mDRIVES(mDevSel), 1)
    Else
        GetLetterUSB = Left(mDRIVES(i), 1)
    End If
End Function

Public Function GetNameUSB(i As Long) As String
    If i > GetCantidadUSB Then Exit Function
    If i = -1 Then
        GetNameUSB = udtInfo(mDevSel).VolumeName
    Else
        GetNameUSB = udtInfo(i).VolumeName
    End If
End Function

Public Function GetTotalMB(i As Long) As Single
    If i > GetCantidadUSB Then Exit Function
    If i = -1 Then
        GetTotalMB = udtInfo(mDevSel).TotalSizeMB
    Else
        GetTotalMB = udtInfo(i).TotalSizeMB
    End If
End Function

Public Function GetFreeMB(i As Long) As Single
    If i > GetCantidadUSB Then Exit Function
    If i = -1 Then
        GetFreeMB = udtInfo(mDevSel).FreeSpaceMB
    Else
        GetFreeMB = udtInfo(i).FreeSpaceMB
    End If
End Function

Public Function GetIndexByLetter(DRVletter As String)
    Dim J As Long
    For J = 1 To UBound(mDRIVES)
        If Left(CStr(mDRIVES(J)), 1) = DRVletter Then
            GetIndexByLetter = J
            Exit Function
        End If
    Next J
    'no lo encontro!
    GetIndexByLetter = -1
End Function
