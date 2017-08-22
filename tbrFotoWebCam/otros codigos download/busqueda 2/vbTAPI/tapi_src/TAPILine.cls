VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CvbTAPILine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'****************************************************************
'*  VB file:   TAPILine.cls... VB32 wrapper for MS TAPI LINE API
'*
'*  created:        1999 by Ray Mercer
'*
'*  8/25/99 by Ray Mercer
'*
'*
'*  Copyright (C) 1999-2001 Ray Mercer.  All rights reserved.
'*  Latest version at http://www.shrinkwrapvb.com
'****************************************************************

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                    (dest As Any, src As Any, ByVal length As Long)

'DrawIconEx Constants
Private Const DI_MASK          As Long = &H1&
Private Const DI_IMAGE         As Long = &H2&
Private Const DI_NORMAL        As Long = &H3&
Private Const DI_COMPAT        As Long = &H4&
Private Const DI_DEFAULTSIZE   As Long = &H8&
Private Declare Function DrawIconEx Lib "user32.dll" _
                            (ByVal hdc As Long, ByVal left As Long, ByVal top As Long, ByVal hIcon As Long, _
                            ByVal width As Long, ByVal height As Long, ByVal step As Long, ByVal hBrush As Long, _
                            ByVal uFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long  'BOOL

'Initialization, error & version control variables
Private m_hLineApp As Long 'handle to TAPI
Private m_APIversions() As Long
Private m_ExtAPIversions() As LINEEXTENSIONID
Private m_ApplicationVersion As Long 'some TAPI calls need this instead of negotiated API version
Private m_LowAPI As Long 'default TAPI 1.3 (&H00010003)
Private m_HiAPI As Long 'default TAPI 3.0 (&H00030000)
Private m_LastTAPIError As Long

'line selection variables
Private m_numLines As Long
Private m_CurLineID As Long

'devcaps variables
Private m_linecaps As LINEDEVCAPS
Private m_LineName As String
Private m_ProviderInfo As String
Private m_SwitchInfo As String
Private m_PermanentLineID As Long
Private m_StringFormat As Long
Private m_numAddresses As Long
Private m_maxDataRate As Long
Private m_BearerModes As Long
Private m_AddressModes As Long
Private m_MediaModes As Long
Private m_GenerateToneMaxNumFreq As Long
Private m_GenerateToneModes As Long
Private m_numTerminals As Long

'Call-state and asynch function state variables
Private m_hLine As Long
Private m_hCall As Long
Private m_RequestingCall As Long
Private m_DroppingCall As Long
Private m_ReqPrivileges As Long
Private m_ReqMediaModes As Long

'Events
Event MakeCallResult(ByVal errorCode As Long)
Event DropCallResult(ByVal errorCode As Long)
Event Connected()
Event Idle()
Event Disconnected()

Private Sub Class_Initialize()
    Debug.Print "class init"
    'default TAPI 1.3 - 3.0
    m_LowAPI = &H10003
    m_HiAPI = &H30000
End Sub
Public Property Get LowAPI() As Long
    LowAPI = m_LowAPI
End Property
Public Property Let LowAPI(ByVal ver As Long)
    m_LowAPI = ver
End Property

Public Property Get HiAPI() As Long
    HiAPI = m_HiAPI
End Property
Public Property Let HiAPI(ByVal ver As Long)
    m_HiAPI = ver
End Property


Public Property Get LastError() As Long
    LastError = m_LastTAPIError
    'reset error value after access
    m_LastTAPIError = 0&
End Property

Public Function ErrorString(ByVal errCode As Long) As String
    ErrorString = GetLineErrString(errCode)
End Function

Public Function Create() As Boolean

    Dim rc As Long
    Dim line As Long
    Dim lip As LINEINITIALIZEEXPARAMS
    Dim lxid As LINEEXTENSIONID
    
    'init params
    lip.dwTotalSize = Len(lip)
    lip.dwOptions = LINEINITIALIZEEXOPTION_USEHIDDENWINDOW
    
    'initialize TAPI lines and get handle
    rc = lineInitializeEx(m_hLineApp, App.hInstance, AddressOf mTAPIvb.LineCallbackProc, App.Title, _
                                m_numLines, m_HiAPI, lip)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Exit Function
    Else
        If m_numLines < 1 Then Exit Function 'no lines!
        'negotiate and cache API versions for each line
        ReDim m_APIversions(0 To m_numLines - 1)
        ReDim m_ExtAPIversions(0 To m_numLines - 1)
        For line = 0 To m_numLines - 1
            rc = lineNegotiateAPIVersion(m_hLineApp, line, m_LowAPI, m_HiAPI, m_APIversions(line), lxid)
            If rc <> TAPI_SUCCESS Then
                m_APIversions(line) = 0 'no compatible API negotiated for this line
                rc = 0
            Else
                'also cache extension version (dev-specific features) if available
                m_ExtAPIversions(line).dwExtensionID0 = lxid.dwExtensionID0
                m_ExtAPIversions(line).dwExtensionID1 = lxid.dwExtensionID1
                m_ExtAPIversions(line).dwExtensionID2 = lxid.dwExtensionID2
                m_ExtAPIversions(line).dwExtensionID3 = lxid.dwExtensionID3
                'TODO! add call to lineNegotiateExtVersion()
                'and save the highest negotiated version as our app version
                If m_APIversions(line) > m_ApplicationVersion Then
                    m_ApplicationVersion = m_APIversions(line)
                End If
            End If
                
        Next
    End If
    rc = GetLineDevCaps()
    Create = True 'return success

End Function

Public Property Get numLines() As Long
    numLines = m_numLines
End Property

Public Property Get CurrentLineID() As Long
    CurrentLineID = m_CurLineID
End Property

Public Property Let CurrentLineID(ByVal id As Long)
    If id < 0 Then Err.Raise 380
    If id > m_numLines - 1 Then Err.Raise 380
    
    m_CurLineID = id
    Call GetLineDevCaps
    
End Property

Public Property Get NegotiatedAPIVersion() As Long
    NegotiatedAPIVersion = m_APIversions(m_CurLineID)
End Property

Public Property Get ExtAPI_ID0() As Long
   ExtAPI_ID0 = m_ExtAPIversions(m_CurLineID).dwExtensionID0
End Property

Public Property Get ExtAPI_ID1() As Long
   ExtAPI_ID1 = m_ExtAPIversions(m_CurLineID).dwExtensionID1
End Property

Public Property Get ExtAPI_ID2() As Long
   ExtAPI_ID2 = m_ExtAPIversions(m_CurLineID).dwExtensionID2
End Property

Public Property Get ExtAPI_ID3() As Long
   ExtAPI_ID3 = m_ExtAPIversions(m_CurLineID).dwExtensionID3
End Property

Private Function GetLineDevCaps() As Boolean
    Dim rc As Long
    
    'init required fields in devcaps UDT
    m_linecaps.dwTotalSize = Len(m_linecaps)
    
    'the LINEDEVCAPS UDT is just hard-allocated to be big enough to hold most
    'variable-length structures that TAPI may throw at you
    'if there is a better way from VB, I'd like to know (besides a Byte-array that is)...
    
    'request TAPI to fill UDT with info
    rc = lineGetDevCaps(m_hLineApp, m_CurLineID, m_APIversions(m_CurLineID), _
                        0&, m_linecaps)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Exit Function
    End If
    
    'store UDT info in class local variables
    m_LineName = GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwLineNameOffset, m_linecaps.dwLineNameSize)
    m_ProviderInfo = GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwProviderInfoOffset, m_linecaps.dwProviderInfoSize)
    m_SwitchInfo = GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwSwitchInfoOffset, m_linecaps.dwSwitchInfoSize)
    m_PermanentLineID = m_linecaps.dwPermanentLineID
    m_StringFormat = m_linecaps.dwStringFormat
    m_numAddresses = m_linecaps.dwNumAddresses
    m_maxDataRate = m_linecaps.dwMaxRate
    m_BearerModes = m_linecaps.dwBearerModes
    m_AddressModes = m_linecaps.dwAddressModes
    m_MediaModes = m_linecaps.dwMediaModes
    m_GenerateToneMaxNumFreq = m_linecaps.dwGenerateToneMaxNumFreq
    m_GenerateToneModes = m_linecaps.dwGenerateToneModes
    m_numTerminals = m_linecaps.dwNumTerminals
    
     'return success
    GetLineDevCaps = True
End Function
Public Property Get LineName() As String
    LineName = m_LineName
End Property
Public Property Get ProviderInfo() As String
    ProviderInfo = m_ProviderInfo
End Property
Public Property Get SwitchInfo() As String
    SwitchInfo = m_SwitchInfo
End Property
Public Property Get PermanentLineID() As Long
    PermanentLineID = m_PermanentLineID
End Property

Public Property Get StringFormat() As Long
    StringFormat = m_StringFormat
End Property
Public Property Get numAddresses() As Long
    numAddresses = m_numAddresses
End Property

Public Property Get maxDataRate() As Long
    maxDataRate = m_maxDataRate
End Property
Public Property Get BearerModes() As Long
    BearerModes = m_BearerModes
End Property
Public Property Get AddressModes() As Long
    AddressModes = m_AddressModes
End Property
Public Property Get mediamodes() As Long
    mediamodes = m_MediaModes
End Property
Public Property Get GenerateToneMaxNumFreq() As Long
    GenerateToneMaxNumFreq = m_GenerateToneMaxNumFreq
End Property
Public Property Get GenerateToneModes() As Long
    GenerateToneModes = m_GenerateToneModes
End Property
Public Property Get numTerminals() As Long
    numTerminals = m_numTerminals
End Property

Public Property Get LineSupportsVoiceCalls() As Boolean
    If m_BearerModes And LINEBEARERMODE_VOICE Then
        If m_MediaModes And LINEMEDIAMODE_INTERACTIVEVOICE Then
            LineSupportsVoiceCalls = True
        End If
    End If
End Property


Public Function OpenLine(Optional ByVal privileges As Long = LINECALLPRIVILEGE_NONE, _
                        Optional ByVal mediamodes As Long = LINEMEDIAMODE_INTERACTIVEVOICE) As Boolean
    'returns false on error or if a line is already open
    'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
    If m_hLine <> 0 Then
        Debug.Print "OpenLine called recursively!"
        Exit Function
    End If
    Dim rc As Long
    
    'open the line for outgoing call
    '(passes a reference to itself in the dwCallbackinstance parameter
    'tapi will pass that reference back to the callback procedure and
    'we will use the ITapiCallbackSink interface to call back to this
    'class instance)
    rc = lineOpen(m_hLineApp, _
                    m_CurLineID, _
                    m_hLine, _
                    m_APIversions(m_CurLineID), _
                    0&, _
                    ByVal ObjPtr(Me), _
                    privileges, _
                    mediamodes, _
                    ByVal 0&)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Debug.Print ErrorString(rc)
        Exit Function
    End If
    Debug.Assert m_hLine
    OpenLine = True

End Function
Public Function MakeCallAsynch(Optional ByVal phonenumber As String = "") As Boolean
    'returns false on error or if m_hLine has not been initialized(by calling OpenLine)
    'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
    Dim rc As Long
    'Dim callParams As LINECALLPARAMS
    
    If 0 = m_hLine Then
        Debug.Print "MakeCallAsynch entered recursively!"
        Exit Function
    End If
    Debug.Print "MakeCallAsynch entered"
    
    'callParams.dwTotalSize = Len(callParams)
    If phonenumber = "" Then 'just get dialtone
        rc = lineMakeCall(m_hLine, m_hCall, vbNullString, 0&, ByVal 0&)
    Else 'dial a number
        rc = lineMakeCall(m_hLine, m_hCall, phonenumber, 0&, ByVal 0&)
    End If
    If rc > 0 Then
        'now placing call - completion will be signalled by a LINE_REPLY event in the callback
        'store positive request identifier in class member variable for use in the callback handler
        m_RequestingCall = rc
    Else
        m_LastTAPIError = rc
        Debug.Print "MakeCallAsynch error"
        Exit Function
    End If
    Debug.Print "MakeCallAsynch successful"
    MakeCallAsynch = True
    
End Function

Public Function DropCallAsynch() As Boolean
    'returns false on error or if m_hLine has not been initialized(by calling OpenLine)
    'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
    Dim rc As Long
    
    If m_hCall = 0 Then
        Debug.Print "Can't drop call - no hCall!"
        Exit Function
    End If
    
    rc = lineDrop(m_hCall, vbNullString, 0&)
    If rc > 0 Then
        'dropping call
        m_DroppingCall = rc
    Else
        m_LastTAPIError = rc
        Exit Function
    End If
    DropCallAsynch = True
    
End Function

Public Function CloseLine() As Boolean
    'returns false on error or if a line is already closed
    'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
    Dim rc As Long
    
    If 0 = m_hLine Then
        Debug.Print "No hLine to close!"
        Exit Function
    End If
    
    rc = lineClose(m_hLine)
    'line handle is now invalid
    Debug.Print "Line Closed"
    m_hLine = 0
    'store errors
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Exit Function
    End If
    
    CloseLine = True
    
End Function

Public Function ConfigDialog(Optional ByVal OwnerHwnd As Long = 0&, Optional ByVal DeviceClass As String = "") As Boolean
    Dim rc As Long
    
    If DeviceClass = "" Then
        rc = lineConfigDialog(m_CurLineID, OwnerHwnd, vbNullString)
    Else
        rc = lineConfigDialog(m_CurLineID, OwnerHwnd, DeviceClass)
    End If
    If rc <> TAPI_SUCCESS Then
       m_LastTAPIError = rc
       Exit Function
    End If
    'indicate success
    ConfigDialog = True
    
End Function

Public Function DialingPropertiesDialog(Optional ByVal OwnerHwnd As Long = 0&, Optional ByVal phonenumber As String = "") As Boolean
    Dim rc As Long
    If phonenumber = "" Then
        rc = lineTranslateDialog(m_hLineApp, m_CurLineID, m_ApplicationVersion, OwnerHwnd, vbNullString)
    Else
        rc = lineTranslateDialog(m_hLineApp, m_CurLineID, m_ApplicationVersion, OwnerHwnd, phonenumber)
    End If
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Exit Function
    End If
    'indicate success
    DialingPropertiesDialog = True
End Function
Public Function PaintDevIcon(ByVal hdc As Long, _
                                Optional ByVal left As Long = 0, _
                                Optional ByVal top As Long = 0, _
                                Optional ByVal width As Long = 0, _
                                Optional ByVal height As Long = 0) As Boolean
    Dim hIcon As Long
    Dim rc As Long
    
    rc = lineGetIcon(m_CurLineID, 0&, hIcon)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        Exit Function
    End If
    rc = DrawIconEx(hdc, left, top, hIcon, width, height, 0&, 0&, DI_NORMAL)
    '(function succeeds even when there is no icon in the TSP)
    If 0 = hIcon Then  'function failed (check actual hIcon instead of return value)
        Exit Function
    End If
    Call DestroyIcon(hIcon)
    
    PaintDevIcon = True
    
End Function


Private Function GetTAPIStructString(ByVal ptrTapistruct As Long, ByVal offset As Long, ByVal length As Long) As String
'ugly C-hacker way to deal with ugly C-hacker TAPI structs (UDTs)
Dim buffer() As Byte

If length < 1 Then Exit Function 'handle erroneous input

If offset Then '
    ReDim buffer(0 To length - 1)
    CopyMemory buffer(0), ByVal ptrTapistruct + offset, length
    GetTAPIStructString = StrConv(buffer, vbUnicode)
End If

End Function

Private Sub Class_Terminate()
    Debug.Print "class term"
    If m_hCall <> 0 Then
        Call lineDeallocateCall(m_hCall)
        m_hCall = 0
    End If
    If m_hLine <> 0 Then
        Call CloseLine
    End If
    If m_hLineApp <> 0 Then
        Call lineShutdown(m_hLineApp)
        m_hLineApp = 0
    End If
End Sub

Friend Sub LineProcHandler(ByVal hDevice As Long, _
                                        ByVal dwMsg As Long, _
                                        ByVal dwParam1 As Long, _
                                        ByVal dwParam2 As Long, _
                                        ByVal dwParam3 As Long)
'Handle callbacks here in the class itself via ITapiCallbackSink
    Debug.Print "Entering LineProcHandler"
    Select Case dwMsg
        Case LINE_REPLY
            If dwParam1 = m_RequestingCall Then
                Debug.Print "LINE_REPLY-CALL REQUEST RETURNS"
                '(earlier we called the asynch lineMakeCall() function from the
                'MakeCallAsynch method and stored the requester in m_RequestingCall
                'now we can see that this is the asynchronous reply to that call)
                'don't need ID anymore - it matched
                m_RequestingCall = 0
                
                'if it was an error make sure the line is closed
                If dwParam2 <> 0 Then Call CloseLine
                'send result to parent via event
                RaiseEvent MakeCallResult(dwParam2)
                
            ElseIf dwParam1 = m_DroppingCall Then
                Debug.Print "LINE_REPLY-CALL DROP RETURNS"
                'asynch reply to lineDrop() call
                m_DroppingCall = 0
                RaiseEvent DropCallResult(dwParam2)
                
            End If
        Case LINE_CALLSTATE
            Select Case dwParam1
                Case LINECALLSTATE_DISCONNECTED
                    'remote party has disconnected from the call
                    Debug.Print "LINECALLSTATE_DISCONNECTED"
                    Call DropCallAsynch
                    RaiseEvent Disconnected
                                    
                Case LINECALLSTATE_IDLE
                    'no call exists - the line is idle
                    Debug.Print "LINECALLSTATE_IDLE"
                    If m_hCall <> 0 Then
                        Debug.Print "Deallocating Call"
                        Call lineDeallocateCall(m_hCall)
                        Debug.Print "Closing Line"
                        Call CloseLine
                    End If
                    RaiseEvent Idle
                Case LINECALLSTATE_CONNECTED
                    Debug.Print "LINECALLSTATE_CONNECTED"
                    RaiseEvent Connected
                                       
                Case LINECALLSTATE_BUSY
                    Debug.Print "LINECALLSTATE_BUSY"
                    'nuff said
                
                Case LINECALLSTATE_DIALTONE
                    'switch is ready to receive a dialed number
                    Debug.Print "LINECALLSTATE_DIALTONE"
                
                Case LINECALLSTATE_RINGBACK
                    'the other station has been reached and is being alerted (ringing)
                    Debug.Print "LINECALLSTATE_RINGBACK"
                
                Case LINECALLSTATE_DIALING
                    Debug.Print "LINECALLSTATE_DIALING"
                
                Case LINECALLSTATE_PROCEEDING
                    Debug.Print "LINECALLSTATE_PROCEEDING"
                
                Case LINECALLSTATE_SPECIALINFO
                    'network error occured
                    Debug.Print "LINECALLSTATE_SPECIALINFO"
                    
                Case Else
                    Debug.Print "CallSTATE: " & Hex(dwParam1)
            End Select
        Case Else
    End Select

End Sub

