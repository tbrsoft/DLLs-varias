Attribute VB_Name = "mTAPIDecs"
Option Explicit
'****************************************************************
'*  VB file:   TAPIDecs.bas...
'*             Partial VB32 translation of tapi.h types and declararions
'*
'*  created:        1999 by Ray Mercer
'*
'*  8/25/99: First public version (added comments)
'*
'*  3/09/2001: bug fixed in lineInitializeEx
'*
'*  These tyeps and decs are in a standard module to facilitate
'*  cutting and pasting into your own code.  Please note that there
'*  are various ways to approach API & type declarations in VB.
'*  This file represents one possible method.
'*
'*  Copyright (c) 1999-2001 Ray Mercer.  All rights reserved.
'*  Latest version at http://www.shrinkwrapvb.com
'****************************************************************

Type LINEINITIALIZEEXPARAMS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwOptions As Long
    hEvent As Long 'union hEvent and Completion port
    dwCompletionKey As Long
End Type

Declare Function lineInitialize Lib "tapi32.dll" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszAppName As String, _
    ByRef lpdwNumDevs As Long) As Long

Declare Function lineInitializeEx Lib "tapi32.dll" Alias "lineInitializeExA" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszFriendlyAppName As String, _
    ByRef lpdwNumDevs As Long, _
    ByRef lpdwAPIVersion As Long, _
    ByRef lpLineInitializeExParams As LINEINITIALIZEEXPARAMS) As Long
    
Declare Function lineGetDevCaps Lib "tapi32.dll" Alias "lineGetDevCapsA" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByRef lpLineDevCaps As LINEDEVCAPS) As Long
 
Declare Function lineConfigDialog Lib "tapi32.dll" Alias "lineConfigDialogA" _
    (ByVal dwDeviceID As Long, _
    ByVal hwndOwner As Long, _
    ByVal lpszDeviceClass As String) As Long
    
Declare Function lineTranslateDialog Lib "tapi32.dll" Alias "lineTranslateDialogA" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal hwndOwner As Long, _
    ByVal lpszAddressIn As String) As Long

Declare Function lineShutdown Lib "tapi32.dll" _
    (ByVal hLineApp As Long) As Long
    
Declare Function lineMakeCall Lib "tapi32.dll" Alias "lineMakeCallA" _
    (ByVal hLine As Long, _
    ByRef lphCall As Long, _
    ByVal lpszDestAddress As String, _
    ByVal dwCountryCode As Long, _
    ByRef lpCallParams As Any) As Long 'LINECALLPARAMS declared As Any so Null value can be passed
    
Declare Function lineDeallocateCall Lib "tapi32.dll" _
    (ByVal hCall As Long) As Long
    
Declare Function lineDrop Lib "tapi32.dll" _
    (ByVal hCall As Long, _
    ByVal lpsUserUserInfo As String, _
    ByVal dwSize As Long) As Long
    
Declare Function lineGetIcon Lib "tapi32.dll" Alias "lineGetIconA" _
    (ByVal dwDeviceID As Long, _
    ByVal lpszDeviceClass As Long, _
    ByRef lphIcon As Long) As Long
    
Public Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Declare Function lineNegotiateAPIVersion Lib "tapi32.dll" _
(ByVal hLineApp As Long, _
ByVal dwDeviceID As Long, _
ByVal dwAPILowVersion As Long, _
ByVal dwAPIHighVersion As Long, _
ByRef lpdwAPIVersion As Long, _
ByRef lpExtensionID As LINEEXTENSIONID) As Long

Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Type LINEDEVCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long
    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long
    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long
    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long
    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long
    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long
    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    dwLineStates As Long
    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserUserInfoSize As Long
    dwUUICallInfoSize As Long
    MinDialParams As LINEDIALPARAMS
    MaxDialParams As LINEDIALPARAMS
    DefaultDialParams As LINEDIALPARAMS
    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long
    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    dwLineFeatures As Long                                 '// TAPI v1.4
'#if (TAPI_CURRENT_VERSION >= 0x00020000)
    dwSettableDevStatus As Long                            '// TAPI v2.0
    dwDeviceClassesSize As Long                            ' // TAPI v2.0
    dwDeviceClassesOffset As Long                          ' // TAPI v2.0
'#End If
    'my way of handling TAPI variable sized structures (yech!)
    vbByteBuffer(0 To 2048) As Byte
    'note*  if you get LINEERR_STRUCTURETOOSMALL and you know that you are
    'doing everything else right (Like initializing the dwActualSize parameter
    'of structs you are passing) then you *might* need to increase this buffer
    'size and recompile.  However, I have not had any problems with this size yet.
End Type
 

Declare Function lineClose Lib "tapi32.dll" _
    (ByVal hLine As Long) As Long

Declare Function lineOpen Lib "tapi32.dll" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByRef lphLine As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwPrivileges As Long, _
    ByVal dwMediaModes As Long, _
    ByRef lpCallParams As Any) As Long 'LINECALLPARAMS declared As Any so NULL can be passed


Public Type LINECALLPARAMS                 '// DEFAULTS
    dwTotalSize As Long                    '// ---------
    dwBearerMode As Long                   '// voice
    dwMinRate As Long                      '// (3.1kHz)
    dwMaxRate As Long                      '// (3.1kHz)
    dwMediaMode As Long                    '// interactiveVoice
    dwCallParamFlags As Long               '// 0
    dwAddressMode As Long                  '// addressID
    dwAddressID As Long                    '// (any available)
    DialParams As LINEDIALPARAMS           '// (0, 0, 0, 0)
    dwOrigAddressSize As Long              '// 0
    dwOrigAddressOffset As Long
    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long
    dwCalledPartySize As Long              '// 0
    dwCalledPartyOffset As Long
    dwCommentSize As Long                  '// 0
    dwCommentOffset As Long
    dwUserUserInfoSize As Long             '// 0
    dwUserUserInfoOffset As Long
    dwHighLevelCompSize As Long            '// 0
    dwHighLevelCompOffset As Long
    dwLowLevelCompSize As Long             '// 0
    dwLowLevelCompOffset As Long
    dwDevSpecificSize As Long              '// 0
    dwDevSpecificOffset As Long
'#if (TAPI_CURRENT_VERSION >= 0x00020000)
    dwPredictiveAutoTransferStates As Long                 '// TAPI v2.0
    dwTargetAddressSize As Long                            '// TAPI v2.0
    dwTargetAddressOffset As Long                          '// TAPI v2.0
    dwSendingFlowspecSize As Long                          '// TAPI v2.0
    dwSendingFlowspecOffset As Long                        '// TAPI v2.0
    dwReceivingFlowspecSize As Long                        '// TAPI v2.0
    dwReceivingFlowspecOffset As Long                      '// TAPI v2.0
    dwDeviceClassSize As Long                              '// TAPI v2.0
    dwDeviceClassOffset As Long                            '// TAPI v2.0
    dwDeviceConfigSize As Long                             '// TAPI v2.0
    dwDeviceConfigOffset As Long                           '// TAPI v2.0
    dwCallDataSize As Long                                 '// TAPI v2.0
    dwCallDataOffset As Long                               '// TAPI v2.0
    dwNoAnswerTimeout As Long                              '// TAPI v2.0
    dwCallingPartyIDSize As Long                           '// TAPI v2.0
    dwCallingPartyIDOffset As Long                         '// TAPI v2.0
'#End If
End Type

