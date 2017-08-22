VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CMixer class for VB
'Copyright (c) 1998-1999 by Shrinkwrap Visual Basic, all rights reserved
'Version 1.0 by Ray Mercer, 1998
'Version 2.0 by Dave Snyder, 04/02/1999
'wraps the Win32 audio mixer object
'freely available - exclusively from http://i.am/shrinkwrapvb
'
Option Explicit

'Mixers
Private m_numMixers As Long             'number of mixer devices on this system at the time Create() is called
Private m_mxID As Long                  'index to currently selected mixer
Private m_hmixer() As Long              'array of mixer handles
Private m_mxCaps() As MIXERCAPS         'array of mixer DEVCAPS UDTs
'Destination Lines
Private m_DstLine() As MIXERLINE         '2-dimesional array of MIXERLINE UDTS (I store ALL Destination Lines on ALL mixers in here)
Private m_currentDstLineID() As Long     'index to currently selected Dest Line for each mixer
'Source Line
Private m_SrcLine As MIXERLINE          'one MIXERLINE struct is reused over and over again to retieve info on Source Lines
Private m_currentSrcLineID As Long             'index to current source line for current Dst line
'line control struct
Private m_LineControls As MIXERLINECONTROLS
'Destination Control Info
Private m_DstControl() As MIXERCONTROL
Private m_currentDstControlID
'Destination Channel Info
Private m_DstChannel() As Variant
Private m_currentDstChannelID As Long
Private m_currentDstControlDetails As MIXERCONTROLDETAILS
'Source Control Info
Private m_SrcControl() As MIXERCONTROL
Private m_currentSrcControlID
'Source Channel info
Private m_SrcChannel() As Variant
Private m_currentSrcChannelID As Long
Private m_currentSrcControlDetails As MIXERCONTROLDETAILS

Public Function Create(ByVal parentHwnd As Long) As Long
    'initializes CWavMixer Class, should be called first
    'returns number of mixer devices on success
    'returns 0 if there are no mixer devices on this system
    Dim i As Long
    Dim ret As vbMMSYSERRORS
    
    'if this is not the first time Create() has been called then destroy and re-initialize class
    If 0 <> m_numMixers Then
        Call Destroy
    End If
    m_numMixers = mixerGetNumDevs()
    If 0 = m_numMixers Then
        'no audio mixer devices available
        Exit Function
    End If
    ReDim m_currentDstLineID(0 To m_numMixers - 1) 'allocate array of DstLineID pointers
    
    'initialize mixer variables
    ReDim m_mxCaps(0 To m_numMixers - 1)  'allocate array MIXERCAPS
    For i = 0 To m_numMixers - 1
        ret = mixerGetDevCaps(i, m_mxCaps(i), Len(m_mxCaps(i)))
        If ret <> MMSYSERR_NOERROR Then
            Err.Raise ret, "CMixer:Create", mmsysGetErrorString(ret)
            Exit Function
        End If
    Next i
    Debug.Assert m_mxCaps(m_mxID).cDestinations > 0 'should never be zero because all audio devices must support at least one dest
    'open mixers to guarantee info will not change
    Call OpenMixers
    
    'fill member variables with info about Destination lines for all mixers
    Call InitializeDestLines
    
    Create = m_numMixers 'return number of mixers
End Function

Private Function OpenMixers() As Long
      'Opens all mixers
      'returns handle of default mixer (ID 0) on success, zero on failure
   Dim ret As vbMMSYSERRORS
   Dim i As Long
   
   If m_numMixers < 1 Then
        'no mixers or Create() not called yet
        Debug.Print "m_nummixers = 0 - Can't Open Mixers!!!"
        Exit Function
   End If
   
   ReDim Preserve m_hmixer(0 To m_numMixers - 1)  'allocate array of hmixers
   
   'close previously opened mixers
   For i = 0 To m_numMixers - 1
            If m_hmixer(i) <> 0 Then
            'close mixer first
            Call mixerClose(m_hmixer(i))
            m_hmixer(i) = 0
        End If
   Next i
   'now open all mixers and save handles
   For i = 0 To m_numMixers - 1
        ret = mixerOpen(m_hmixer(i), i, 0&, 0&, MIXER_OBJECTF_MIXER)
        If MMSYSERR_NOERROR <> ret Then
            Err.Raise ret, _
                        "CMixer:OpenMixers", _
                        "Could not open " & GetMixerName(i) & vbCrLf & "Error: " & mmsysGetErrorString(ret)
        End If
    Next i
   OpenMixers = m_hmixer(m_mxID)
End Function

Private Function InitializeDestLines() As Long
    'initialize DstLine variables for all mixers
    'returns number of controls associated with currently selected line
    Dim i As Long, j As Long
    Dim ret As vbMMSYSERRORS
    
    'Get largest number of destination lines needed (save in j)
    j = 1 'minimum is 1
    For i = 0 To m_numMixers - 1
        If m_mxCaps(i).cDestinations > j Then
            j = m_mxCaps(i).cDestinations
        End If
    Next i
    ReDim m_DstLine(0 To m_numMixers - 1, 0 To j - 1) 'allocate 2 dimensions (mixer, DstLines array)
    'now loop through and get Line Info for each DstLine in each mixer
    For i = 0 To m_numMixers - 1 'for each mixer
        For j = 0 To m_mxCaps(i).cDestinations - 1 'for each dest line on that mixer
            'initialize the two necessary members of DstLine
            m_DstLine(i, j).cbStruct = Len(m_DstLine(i, j))
            m_DstLine(i, j).dwDestination = j
            ret = mixerGetLineInfo(m_hmixer(i), m_DstLine(i, j), MIXER_GETLINEINFOF_DESTINATION)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                            "CMixer:InitializeDestLines", _
                            "Could not retrieve line info for " & GetMixerName(i) & vbCrLf & "Error: " & mmsysGetErrorString(ret)
            End If
            
        Next j
    Next i
    'm_currentDstControlID = 0
    InitializeDestLines = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cConnections 'number of controls (can be zero)
End Function

Private Function InitializeSrcLine() As Long
    'initializes m_SrcLine variable with line indicated by LineID
    'returns number of controls associated with currently selected line
    
   Dim ret As vbMMSYSERRORS
     
   
     
     
    m_SrcLine.cbStruct = Len(m_SrcLine)
    m_SrcLine.dwDestination = m_currentDstLineID(m_mxID)
    m_SrcLine.dwSource = m_currentSrcLineID
    ret = mixerGetLineInfo(m_hmixer(m_mxID), m_SrcLine, MIXER_GETLINEINFOF_SOURCE)
    If ret <> MMSYSERR_NOERROR Then
        Err.Raise ret, _
                    "CMixer:InitializeSrcLine", _
                    "Could not retrieve line info for " & GetMixerName(m_mxID) & vbCrLf & "Error: " & mmsysGetErrorString(ret)
        Exit Function
    End If
    'm_currentSrcControlID = 0
    InitializeSrcLine = m_SrcLine.cConnections 'number of controls (can be zero)
End Function

Public Property Get MixerID() As Long
    MixerID = m_mxID
End Property

Public Property Let MixerID(ByVal mxID As Long)
    If mxID < 0 Or mxID > m_numMixers - 1 Then
        Err.Raise 380, "CMixer:MixerID[LET]", "There are only " & m_numMixers & "mixers on this system." 'invalid property
        Exit Property
    End If
    If m_mxID <> mxID Then
        m_mxID = mxID
    End If
        End Property

Public Property Get DstLineID() As Long
    DstLineID = m_currentDstLineID(m_mxID)
End Property

Public Property Let DstLineID(ByVal lineID As Long)
    If lineID < 0 Or lineID > m_mxCaps(m_mxID).cDestinations - 1 Then
        Err.Raise 380, "CMixer:DstLineID[LET]", "There are only " & m_mxCaps(m_mxID).cDestinations & " destination lines on this mixer." 'invalid property
        Exit Property
    End If
    m_currentDstLineID(m_mxID) = lineID
    m_currentDstControlID = 0
    Call InitializeDestControls
    
End Property

Public Property Get SrcControlID() As Long
    SrcControlID = m_currentSrcControlID
End Property
Public Property Get DstControlID() As Long
    DstControlID = m_currentDstControlID
End Property
Public Property Get DstChannelMaxValue() As Long
    DstChannelMaxValue = m_DstControl(m_currentDstControlID).lMaximum
End Property
Public Property Get DstChannelMinValue() As Long
    DstChannelMinValue = m_DstControl(m_currentDstControlID).lMinimum
End Property
Public Property Get SrcChannelMaxValue() As Long
    SrcChannelMaxValue = m_SrcControl(m_currentSrcControlID).lMaximum
End Property
Public Property Get SrcChannelMinValue() As Long
    SrcChannelMinValue = m_SrcControl(m_currentSrcControlID).lMinimum
End Property

Public Property Let DstControlID(ByVal controlid As Long)
    If controlid < 0 Or controlid > m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls - 1 Then
        Err.Raise 380, "CMixer:DstcontrolID[LET]", "There are only " & m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls & " controls  on this line." 'invalid property
        Exit Property
    End If
    
    m_currentDstControlID = controlid
    m_currentDstChannelID = 0
    Call GetDestChannelInfo
    End Property

Public Property Let SrcControlID(ByVal controlid As Long)
    If controlid < 0 Or controlid > m_SrcLine.cControls - 1 Then
        Err.Raise 380, "CMixer:SrccontrolID[LET]", "There are only " & m_SrcLine.cControls & " controls  on this line." 'invalid property
        Exit Property
    End If
    
    m_currentSrcControlID = controlid
    m_currentSrcChannelID = 0
    Call GetSrcChannelInfo
    End Property

Public Property Get SrcLineID() As Long
    SrcLineID = m_currentSrcLineID
End Property

Public Property Let SrcLineID(ByVal lineID As Long)
    If lineID < 0 Or lineID > m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cConnections - 1 Then
        Err.Raise 380, "CMixer:SrcLineID[LET]", "There are only " & CStr(m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cConnections) & _
                                                " source lines associated with the current destination line ( " _
                                                & GetDstLineName(LongName:=False) & " )" 'invalid property
        Exit Property
    End If
    m_currentSrcLineID = lineID
    m_currentSrcControlID = 0
        Call InitializeSrcLine
        Call InitializeSrcControls
        
   End Property

Public Property Get MixerName() As String
    MixerName = GetMixerName(m_mxID)
End Property

Private Function GetMixerName(ByVal ID As Long) As String
    Debug.Assert Not (m_mxID < 0 Or m_mxID > m_numMixers - 1)
    GetMixerName = Left$(m_mxCaps(m_mxID).szPname, InStr(m_mxCaps(m_mxID).szPname, vbNullChar) - 1)
End Function

Public Property Get numMixerDevs() As Long
    'NumMixerDevs = mixerGetNumDevs()
    numMixerDevs = m_numMixers
End Property

Public Property Get hMixer() As Long
    hMixer = m_hmixer(m_mxID)
End Property

Public Property Get numDestinations() As Long
    numDestinations = m_mxCaps(m_mxID).cDestinations
End Property

Public Property Get MixerVerMajor() As Long
    Dim wVer As Integer
    
    wVer = LoWord(m_mxCaps(m_mxID).vDriverVersion)
    
    MixerVerMajor = HiByte(wVer)
End Property

Public Property Get MixerVerMinor() As Long
    Dim wVer As Integer
    
    wVer = LoWord(m_mxCaps(m_mxID).vDriverVersion)
    
    MixerVerMinor = LoByte(wVer)
End Property

Public Property Get DstLineName(Optional ByVal LongName As Boolean = True) As String
    DstLineName = GetDstLineName(LongName)
End Property

Private Function GetDstLineName(ByVal LongName As Boolean) As String
    If LongName Then
        GetDstLineName = Left$(m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).szName, InStr(m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).szName, vbNullChar) - 1)
    Else
        GetDstLineName = Left$(m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).szShortName, InStr(m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).szShortName, vbNullChar) - 1)
    End If
End Function

Public Property Get SrcLineName(Optional ByVal LongName As Boolean = True) As String
    SrcLineName = GetSrcLineName(LongName)
End Property

Private Function GetSrcLineName(ByVal LongName As Boolean) As String
    If LongName Then
        GetSrcLineName = Left$(m_SrcLine.szName, InStr(m_SrcLine.szName, vbNullChar) - 1)
    Else
        GetSrcLineName = Left$(m_SrcLine.szShortName, InStr(m_SrcLine.szShortName, vbNullChar) - 1)
    End If
End Function

Public Property Get numDstLineChannels() As Long
    numDstLineChannels = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels
End Property

Public Property Get numDstLineControls() As Long
    numDstLineControls = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls
End Property
Public Property Get numDstControlChannels() As Long
    numDstControlChannels = m_currentDstControlDetails.cChannels
End Property

Public Property Get numDstLineConnections() As Long
    numDstLineConnections = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cConnections
End Property

Public Property Get DstLineActive() As Boolean
    DstLineActive = MIXERLINE_LINEF_ACTIVE And m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).fdwLine
End Property

Public Property Get DstLineDisconnected() As Boolean
    DstLineDisconnected = MIXERLINE_LINEF_DISCONNECTED And m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).fdwLine
End Property

Public Property Get DstLineType() As MIXER_LINE_TYPE
    DstLineType = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).dwComponentType
End Property

Public Property Get DstLineTarget() As TARGET_TYPE
    DstLineTarget = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).Target.dwType
End Property

Public Property Get numSrcLineChannels() As Long
    numSrcLineChannels = m_SrcLine.cChannels
End Property

Public Property Get numSrcLineControls() As Long
    numSrcLineControls = m_SrcLine.cControls
End Property
Public Property Get numSrcControlChannels() As Long
    numSrcControlChannels = m_currentSrcControlDetails.cChannels
End Property

Public Property Get numSrcLineConnections() As Long
    numSrcLineConnections = m_SrcLine.cConnections
End Property

Public Property Get SrcLineActive() As Boolean
    SrcLineActive = MIXERLINE_LINEF_ACTIVE And m_SrcLine.fdwLine
End Property

Public Property Get SrcLineDisconnected() As Boolean
    SrcLineDisconnected = MIXERLINE_LINEF_DISCONNECTED And m_SrcLine.fdwLine
End Property

Public Property Get SrcLineType() As MIXER_LINE_TYPE
    SrcLineType = m_SrcLine.dwComponentType
End Property

Public Property Get SrcLineTarget() As TARGET_TYPE
    SrcLineTarget = m_SrcLine.Target.dwType
End Property

Public Property Get DstControlType() As Long
    DstControlType = m_DstControl(m_currentDstControlID).dwControlType
End Property

Public Property Get SrcControlType() As Long
    SrcControlType = m_SrcControl(m_currentSrcControlID).dwControlType
End Property

Public Property Get DstControlName(Optional ByVal LongName As Boolean = True) As String
    DstControlName = GetDstControlName(LongName)
End Property

Public Property Get SrcControlName(Optional ByVal LongName As Boolean = True) As String
    SrcControlName = GetSrcControlName(LongName)
End Property


Public Property Get DstChannelID() As Long
    DstChannelID = m_currentDstChannelID
End Property

Public Property Let DstChannelID(ByVal controlid As Long)
    If controlid < 0 Or controlid > m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels - 1 Then
        Err.Raise 380, "CMixer:DstChannelID[LET]", "There are only " & m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels & " channels  on this line." 'invalid property
        Exit Property
    End If
    
    m_currentDstChannelID = controlid
   
    End Property
Public Property Get DstChannelValue() As Variant
    Call GetDestChannelInfo
    DstChannelValue = m_DstChannel(m_currentDstChannelID)
End Property
Public Property Let DstChannelValue(ByVal value As Variant)
    m_DstChannel(m_currentDstChannelID) = value
    
    Call SetDestChannelInfo
End Property

Public Property Get SrcChannelID() As Long
    SrcChannelID = m_currentSrcChannelID
End Property

Public Property Let SrcChannelID(ByVal controlid As Long)
    If controlid < 0 Or controlid > m_SrcLine.cChannels - 1 Then
        Err.Raise 380, "CMixer:SrcChannelID[LET]", "There are only " & m_SrcLine.cChannels & " channels  on this line." 'invalid property
        Exit Property
    End If
    
    m_currentSrcChannelID = controlid
    
    End Property
Public Property Get SrcChannelValue() As Variant
    Call GetSrcChannelInfo
    SrcChannelValue = m_SrcChannel(m_currentSrcChannelID)
End Property
Public Property Let SrcChannelValue(ByVal value As Variant)
    m_SrcChannel(m_currentSrcChannelID) = value
    
    Call SetSrcChannelInfo
End Property

Private Function GetDstControlName(ByVal LongName As Boolean) As String
    If LongName Then
        GetDstControlName = m_DstControl(m_currentDstControlID).szName
    Else
        GetDstControlName = m_DstControl(m_currentDstControlID).szShortName
    End If
End Function

Private Function GetSrcControlName(ByVal LongName As Boolean) As String
    If LongName Then
        GetSrcControlName = m_SrcControl(m_currentSrcControlID).szName
    Else
        GetSrcControlName = m_SrcControl(m_currentSrcControlID).szShortName
    End If
End Function

Private Sub Destroy()
    Dim i As Long
    'On Local Error Resume Next
    'close all mixers
    For i = 0 To m_numMixers - 1
        Call mixerClose(m_hmixer(i))
        m_hmixer(i) = 0
    Next i
    'zero all variables
    m_numMixers = 0
    m_mxID = 0
    'm_SrcLine 'struct
    m_currentSrcLineID = 0
    m_currentDstControlID = 0
    m_currentSrcControlID = 0
    m_currentDstChannelID = 0
    m_currentSrcChannelID = 0
    ReDim m_mxCaps(0)
    ReDim m_hmixer(0)
    ReDim m_DstLine(0)
    ReDim m_currentDstLineID(0)
    ReDim m_DstControl(0)
    ReDim m_SrcControl(0)
    ReDim m_DstChannel(0)
    ReDim m_SrcChannel(0)
End Sub



Private Sub Class_Terminate()
    Call Destroy
End Sub


Private Function InitializeSrcControls()

    Dim k As Long, ret As Long, mcsize As Long, hMem
    
    
    mcsize = LenB(m_SrcControl(0))
    ReDim m_SrcControl(0 To m_SrcLine.cControls - 1)
    ReDim m_SrcChannel(0 To m_SrcLine.cChannels - 1)
    
    m_SrcControl(0).cbStruct = mcsize
    m_LineControls.cbStruct = LenB(m_LineControls)
    m_LineControls.dwLineID = m_SrcLine.dwLineID
    m_LineControls.cControls = m_SrcLine.cControls
    m_LineControls.cbmxctrl = mcsize
    hMem = GlobalAlloc(&H40, m_LineControls.cbmxctrl * m_SrcLine.cControls)
    m_LineControls.pamxctrl = GlobalLock(hMem)
    ret = mixerGetLineControls(m_hmixer(m_mxID), m_LineControls, MIXER_GETLINECONTROLSF_ALL)
    If ret <> MMSYSERR_NOERROR Then
        Err.Raise ret, _
                    "CMixer:InitializeSrcLine", _
                    "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error: " & mmsysGetErrorString(ret)
        GlobalFree (hMem)
        Exit Function
    End If
 
        For k = 0 To m_SrcLine.cControls - 1
            CopyStructFromPtr m_SrcControl(k), m_LineControls.pamxctrl + (mcsize * k), Len(m_SrcControl(k))
        Next k
        GlobalFree (hMem)
   
InitializeSrcControls = m_SrcLine.cControls 'number of channels
            
            'Debug.Print m_SrcControl(m_currentSrcControlID).szShortName
            'Debug.Print m_SrcControl(m_currentSrcControlID).szName
            
End Function

Private Function InitializeDestControls()

    Dim k As Long, ret As Long, mcsize As Long, hMem
  ' get Destination controls
  
    mcsize = LenB(m_DstControl(0))
    ReDim m_DstControl(0 To m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls - 1)
    ReDim m_DstChannel(0 To m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels - 1)
    
        m_DstControl(0).cbStruct = mcsize
        m_LineControls.cbStruct = LenB(m_LineControls)
        m_LineControls.dwLineID = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).dwLineID
        m_LineControls.cControls = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls
        m_LineControls.cbmxctrl = mcsize
        hMem = GlobalAlloc(&H40, m_LineControls.cbmxctrl * m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls)
        m_LineControls.pamxctrl = GlobalLock(hMem)
        ret = mixerGetLineControls(m_hmixer(m_mxID), m_LineControls, MIXER_GETLINECONTROLSF_ALL)
        If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
        GlobalFree (hMem)
        Exit Function
        End If
 
        For k = 0 To m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls - 1
            CopyStructFromPtr m_DstControl(k), m_LineControls.pamxctrl + (mcsize * k), Len(m_DstControl(k))
        Next k
        GlobalFree (hMem)
           'Debug.Print m_DstControl(m_currentDstControlID).szShortName
           'Debug.Print m_DstControl(m_currentDstControlID).szName
InitializeDestControls = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cControls 'number of Channels
             
End Function


Private Function GetDestChannelInfo()
 Dim m As Long, hMem As Long, ret As Long
 Dim BoolValue As MIXERCONTROLDETAILS_BOOLEAN, UnsignedValue As MIXERCONTROLDETAILS_UNSIGNED
 Dim SignedValue As MIXERCONTROLDETAILS_SIGNED, ListTextValues As MIXERCONTROLDETAILS_LISTTEXT
               
            'check to see if control is stereo or mono
            
        If m_DstControl(m_currentDstControlID).fdwControl And MIXERCONTROL_CONTROLF_UNIFORM Then
            m_currentDstControlDetails.cChannels = 1
        Else
            m_currentDstControlDetails.cChannels = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels
        End If
            
        Select Case GetDetailType(m_DstControl(m_currentDstControlID).dwControlType)
            Case 0
                
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.cbDetails = Len(BoolValue)
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
                
            ret = mixerGetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

         For m = 0 To m_currentDstControlDetails.cChannels - 1
            CopyStructFromPtr BoolValue, m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), Len(BoolValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_DstChannel(m) = CBool(BoolValue.fValue)
         Next m
            GlobalFree hMem
            
            Case 1
                
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.cbDetails = Len(UnsignedValue)
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
                
            ret = mixerGetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

         For m = 0 To m_currentDstControlDetails.cChannels - 1
            CopyStructFromPtr UnsignedValue, m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), Len(UnsignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_DstChannel(m) = UnsignedValue.dwValue
         Next m
            GlobalFree hMem
           
            Case 2
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
                m_currentDstControlDetails.cbDetails = Len(m_currentDstControlDetails.paDetails)
            ret = mixerGetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If
         
         For m = 0 To m_currentDstControlDetails.cChannels - 1
            CopyStructFromPtr SignedValue, m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), Len(SignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_DstChannel(m) = SignedValue.Lvalue
         Next m
            GlobalFree hMem
            
            Case Else
            m_DstChannel(0) = vbNull 'ListText details forthcoming
            
        End Select
       GetDestChannelInfo = m_currentDstControlDetails.cChannels
End Function
Private Function SetDestChannelInfo()
 Dim m As Long, hMem As Long, ret As Long
 Dim BoolValue As MIXERCONTROLDETAILS_BOOLEAN, UnsignedValue As MIXERCONTROLDETAILS_UNSIGNED
 Dim SignedValue As MIXERCONTROLDETAILS_SIGNED, ListTextValues As MIXERCONTROLDETAILS_LISTTEXT
               
            'check to see if control is stereo or mono
            
        If m_DstControl(m_currentDstControlID).fdwControl And MIXERCONTROL_CONTROLF_UNIFORM Then
            m_currentDstControlDetails.cChannels = 1
        Else
            m_currentDstControlDetails.cChannels = m_DstLine(m_mxID, m_currentDstLineID(m_mxID)).cChannels
        End If
            
        Select Case GetDetailType(m_DstControl(m_currentDstControlID).dwControlType)
            Case 0
                
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.cbDetails = Len(BoolValue)
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
         
         For m = 0 To m_currentDstControlDetails.cChannels - 1
            BoolValue.fValue = CLng(m_DstChannel(m))
            CopyPtrFromStruct m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), BoolValue, Len(BoolValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            
         Next m
                
            ret = mixerSetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

            GlobalFree hMem
            
            Case 1
                
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.cbDetails = Len(UnsignedValue)
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
         
         For m = 0 To m_currentDstControlDetails.cChannels - 1
            UnsignedValue.dwValue = m_DstChannel(m)
            CopyPtrFromStruct m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), UnsignedValue, Len(UnsignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            
         Next m
                                
            ret = mixerSetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

        GlobalFree hMem
           
            Case 2
                m_currentDstControlDetails.cbStruct = Len(m_currentDstControlDetails)
                m_currentDstControlDetails.dwControlID = m_DstControl(m_currentDstControlID).dwControlID
                m_currentDstControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentDstControlDetails.cbDetails)
                m_currentDstControlDetails.paDetails = GlobalLock(hMem)
                m_currentDstControlDetails.cbDetails = Len(m_currentDstControlDetails.paDetails)
         
         For m = 0 To m_currentDstControlDetails.cChannels - 1
            SignedValue.Lvalue = m_DstChannel(m)
            CopyPtrFromStruct m_currentDstControlDetails.paDetails + (m_currentDstControlDetails.cbDetails * m), SignedValue, Len(SignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
         Next m
         
            ret = mixerGetControlDetails(m_mxID, m_currentDstControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeDstControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If
            GlobalFree hMem
            
            Case Else
            'ListText details forthcoming
            
        End Select
       SetDestChannelInfo = m_currentDstControlDetails.cChannels
End Function
Private Function SetSrcChannelInfo()
 Dim m As Long, hMem As Long, ret As Long
 Dim BoolValue As MIXERCONTROLDETAILS_BOOLEAN, UnsignedValue As MIXERCONTROLDETAILS_UNSIGNED
 Dim SignedValue As MIXERCONTROLDETAILS_SIGNED, ListTextValues As MIXERCONTROLDETAILS_LISTTEXT
               
            'check to see if control is stereo or mono
            
        If m_SrcControl(m_currentSrcControlID).fdwControl And MIXERCONTROL_CONTROLF_UNIFORM Then
            m_currentSrcControlDetails.cChannels = 1
        Else
            m_currentSrcControlDetails.cChannels = m_SrcLine.cChannels
        End If
            
        Select Case GetDetailType(m_SrcControl(m_currentSrcControlID).dwControlType)
            Case 0
                
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.cbDetails = Len(BoolValue)
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
         
         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            BoolValue.fValue = CLng(m_SrcChannel(m))
            CopyPtrFromStruct m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), BoolValue, Len(BoolValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            
         Next m
                
            ret = mixerSetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

            GlobalFree hMem
            
            Case 1
                
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.cbDetails = Len(UnsignedValue)
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
         
         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            UnsignedValue.dwValue = m_SrcChannel(m)
            CopyPtrFromStruct m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), UnsignedValue, Len(UnsignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            
         Next m
                                
            ret = mixerSetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

        GlobalFree hMem
           
            Case 2
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
                m_currentSrcControlDetails.cbDetails = Len(m_currentSrcControlDetails.paDetails)
         
         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            SignedValue.Lvalue = m_SrcChannel(m)
            CopyPtrFromStruct m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), SignedValue, Len(SignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
         Next m
         
            ret = mixerGetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If
            GlobalFree hMem
            
            Case Else
            'ListText details forthcoming
            
        End Select
       SetSrcChannelInfo = m_currentSrcControlDetails.cChannels
End Function

Private Function GetSrcChannelInfo()
        Dim m As Long, hMem As Long, ret As Long
        Dim BoolValue As MIXERCONTROLDETAILS_BOOLEAN, UnsignedValue As MIXERCONTROLDETAILS_UNSIGNED
        Dim SignedValue As MIXERCONTROLDETAILS_SIGNED, ListTextValues As MIXERCONTROLDETAILS_LISTTEXT
        
        
            'check to see if control is stereo or mono(uniform)
            
        If m_SrcControl(m_currentSrcControlID).fdwControl And MIXERCONTROL_CONTROLF_UNIFORM Then
            m_currentSrcControlDetails.cChannels = 1
        Else
            m_currentSrcControlDetails.cChannels = m_SrcLine.cChannels
        End If
            
        Select Case GetDetailType(m_SrcControl(m_currentSrcControlID).dwControlType)
            Case 0
                
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.cbDetails = Len(BoolValue)
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
                
            ret = mixerGetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            CopyStructFromPtr BoolValue, m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), Len(BoolValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_SrcChannel(m) = CBool(BoolValue.fValue)
         Next m
            
            GlobalFree hMem
            
            Case 1
                
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.cbDetails = Len(UnsignedValue)
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
                
            ret = mixerGetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If

         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            CopyStructFromPtr UnsignedValue, m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), Len(UnsignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_SrcChannel(m) = UnsignedValue.dwValue
         Next m
            GlobalFree hMem
           
            Case 2
                m_currentSrcControlDetails.cbStruct = Len(m_currentSrcControlDetails)
                m_currentSrcControlDetails.dwControlID = m_SrcControl(m_currentSrcControlID).dwControlID
                m_currentSrcControlDetails.Item = 0
                hMem = GlobalAlloc(&H40, m_currentSrcControlDetails.cbDetails)
                m_currentSrcControlDetails.paDetails = GlobalLock(hMem)
                m_currentSrcControlDetails.cbDetails = Len(m_currentSrcControlDetails.paDetails)
            ret = mixerGetControlDetails(m_mxID, m_currentSrcControlDetails, MIXER_GETCONTROLDETAILSF_VALUE)
            If ret <> MMSYSERR_NOERROR Then
                Err.Raise ret, _
                   "CMixer:InitializeSrcControls", _
                   "Could not retrieve control info for " & GetMixerName(m_mxID) & vbCrLf & "Error:" & mmsysGetErrorString(ret)
            GlobalFree hMem
            Exit Function
            End If
         
         For m = 0 To m_currentSrcControlDetails.cChannels - 1
            CopyStructFromPtr SignedValue, m_currentSrcControlDetails.paDetails + (m_currentSrcControlDetails.cbDetails * m), Len(SignedValue) 'add mcd.cbdetails to mcd,padetails to get other half of stereo volume
            m_SrcChannel(m) = SignedValue.Lvalue
         Next m
            GlobalFree hMem
            
            Case Else
            m_SrcChannel(0) = vbNull 'ListText details forthcoming
            
        End Select
    GetSrcChannelInfo = m_currentSrcControlDetails.cChannels
End Function


Private Function GetDetailType(ByVal ctrlType As MIXERCONTROL_TYPE) As Integer
    Select Case ctrlType
        
        Case mcBASS_FADER
            'Debug.Print "BASS FADER CONTROL"
            GetDetailType = 1 'unsigned
        Case mcBOOLEAN_METER
            'Debug.Print "BOOLEAN METER CONTROL"
            GetDetailType = 0  'boolean
        Case mcBOOLEAN_SWITCH
            'Debug.Print "BOOLEAN SWITCH CONTROL"
            GetDetailType = 0
        Case mcBUTTON_SWITCH
            'Debug.Print "BUTTON SWITCH CONTROL"
            GetDetailType = 0
        Case mcDECIBELS_NUMBER
            'Debug.Print "DECIBELS NUMBER CONTROL"
            GetDetailType = 2 'signed
        Case mcEQUALIZER_FADER
            'Debug.Print "EQUALIZER FADER CONTROL"
            GetDetailType = 1
        Case mcFADER_FADER
            'Debug.Print "FADER CONTROL"
            GetDetailType = 1
        Case mcGENERIC_CUSTOM
            'Debug.Print "CUSTOM CONTROL"
            GetDetailType = 3
        Case mcLOUDNESS_SWITCH
            'Debug.Print "LOUDNESS SWITCH CONTROL"
            GetDetailType = 0
        Case mcMICROTIME_TIME
            'Debug.Print "MICROTIME TIME CONTROL"
            GetDetailType = 1
        Case mcMILLITIME_TIME
            'Debug.Print "MILLITIME TIME CONTROL"
            GetDetailType = 1
        Case mcMIXER_LIST
            'Debug.Print "MIXER LIST CONTROL"
            GetDetailType = 3
        Case mcMONO_SWITCH
            'Debug.Print "MONO SWITCH CONTRL"
            GetDetailType = 0
        Case mcMULTIPLESELECT_LIST
            'Debug.Print "MULTIPLESELECT LIST CONTROL"
            GetDetailType = 3
        Case mcMUTE_SWITCH
            'Debug.Print "MUTE SWITCH CONTROL"
            GetDetailType = 0
        Case mcMUX_LIST
            'Debug.Print "MUX LIST CONTROL"
            GetDetailType = 3
        Case mcONOFF_SWITCH
            'Debug.Print "ONOFF SWITCH CONTROL"
            GetDetailType = 0
        Case mcPAN_SLIDER
            'Debug.Print "PAN SLIDER CONTROL"
            GetDetailType = 2
        Case mcPERCENT_NUMBER
            'Debug.Print "PERCENT NUMBER CONTROL"
            GetDetailType = 1
        Case mcQSOUNDPAN_SLIDER
            'Debug.Print "QSOUND PAN SLIDER CONTROL"
            GetDetailType = 2
        Case mcSIGNED_NUMBER
            'Debug.Print "SIGNED NUMBER CONTROL"
            GetDetailType = 2
        Case mcSINGLESELECT_LIST
            'Debug.Print "SINGLESELECT LIST CONTROL"
            GetDetailType = 3
        Case mcSLIDER_SLIDER
            'Debug.Print "SLIDER CONTROL"
            GetDetailType = 2
        Case mcSTEREOENH_SWITCH
            'Debug.Print "STEREO ENHANCE SWITCH CONTROL"
            GetDetailType = 0
        Case mcTREBLE_FADER
            'Debug.Print "TREBLE FADER CONTROL"
            GetDetailType = 1
        Case mcUNSIGNED_NUMBER
            'Debug.Print "UNSIGNED NUMBER CONTROL"
            GetDetailType = 1
        Case mcVOLUME_FADER
            'Debug.Print "VOLUME FADER CONTROL"
            GetDetailType = 1
        Case mcPEAK_METER
            'Debug.Print "SIGNED PEAK METER"
            GetDetailType = 2
        Case Else
            'Debug.Print "UNKNOWN"
            GetDetailType = 3
    End Select
End Function


