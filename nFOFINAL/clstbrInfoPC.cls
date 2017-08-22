VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstbrInfoPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim TD As New tbrDATA.clsFinalID
Dim TK As New clsKEYS
Private mResFinal As String

Private PthApp As String
Private FSO As New Scripting.FileSystemObject

Public Property Get ResFinal() As String
    ResFinal = mResFinal
End Property

Public Function GetIM() As String 'get info mac
    'devuelve todas las MAC adress juntas
        
    On Local Error GoTo ErrIM
    
    Terr.Anotar "saa"
    Dim TMP As String, FINAL As String
    TMP = "": FINAL = ""
    
    'PR.ShowProgreso "INI", 1
    Terr.Anotar "sab", TD.CantidadMac
    If TD.CantidadMac = 0 Then
        'el hecho de no tener mac no tiene que matar todo lo otro
        'GetIM = "0"
        'Exit Function
    Else
        Dim A As Long
        Dim tmpMac As String
        For A = 1 To TD.CantidadMac
            Terr.Anotar "sac", A
            tmpMac = TD.GetRealMac(A)
            Terr.Anotar "sad", tmpMac
            If tmpMac = "" Then GoTo SIgMAC
            'PR.ShowProgreso "MAC " + CStr(A), 1 + (A * 5)
            'solo las que tienen MAC
            Terr.Anotar "sae", TD.GetnTipo(A)
            If TD.GetnTipo(A) = 6 Then
                TMP = TMP + "MAC " + CStr(A) + ": " + tmpMac + vbCrLf
                FINAL = FINAL + tmpMac
                Terr.Anotar "saf", FINAL
                If A < TD.CantidadMac Then FINAL = FINAL + "."
            End If
SIgMAC:
        Next A
    End If
    Terr.Anotar "sag"
    FINAL = FINAL + "|"
    
    TK.LeerTodo
    
    Terr.Anotar "sah"
    'PR.ShowProgreso "Escribiendo", 90
    TMP = TMP + "BiosCaption: " + TK.mBIOSCaption + vbCrLf
    TMP = TMP + "BiosDate: " + TK.mBiosDate + vbCrLf
    TMP = TMP + "BIOSDescription: " + TK.mBIOSDescription + vbCrLf
    TMP = TMP + "BIOSIdentificationCode: " + TK.mBIOSIdentificationCode + vbCrLf
    TMP = TMP + "BIOSManufacturer: " + TK.mBIOSManufacturer + vbCrLf
    TMP = TMP + "BIOSName: " + TK.mBIOSName + vbCrLf
    TMP = TMP + "BIOSReleaseDate: " + TK.mBIOSReleaseDate + vbCrLf
    TMP = TMP + "BIOSSerialNumber: " + TK.mBIOSSerialNumber + vbCrLf
    Terr.Anotar "sai"
    FINAL = FINAL + TK.mBIOSSerialNumber + "|"
    TMP = TMP + "BIOSSMBIOSBIOSVersion: " + TK.mBIOSSMBIOSBIOSVersion + vbCrLf
    TMP = TMP + "BIOSSMBIOSMajorVersion: " + TK.mBIOSSMBIOSMajorVersion + vbCrLf
    TMP = TMP + "BIOSSMBIOSMinorVersion: " + TK.mBIOSSMBIOSMinorVersion + vbCrLf
    TMP = TMP + "BIOSVersion: " + TK.mBIOSVersion + vbCrLf
    'PR.ShowProgreso "Escribiendo", 95
    TMP = TMP + "Country: " + TK.mCountry + vbCrLf
    TMP = TMP + "PROCCaption: " + TK.mPROCCaption + vbCrLf
    TMP = TMP + "PROCCCspeed: " + TK.mPROCCCspeed + vbCrLf
    TMP = TMP + "PROCDeviceId: " + TK.mPROCDeviceId + vbCrLf
    TMP = TMP + "PROCFamily: " + TK.mPROCFamily + vbCrLf
    TMP = TMP + "PROCId: " + TK.mPROCId + vbCrLf
    Terr.Anotar "saj"
    FINAL = FINAL + TK.mPROCId + "|"
    TMP = TMP + "PROCInfoReserved: " + TK.mPROCInfoReserved + vbCrLf
    FINAL = FINAL + TK.mPROCInfoReserved + "|"
    TMP = TMP + "PROCRevis: " + TK.mPROCRevis + vbCrLf
    TMP = TMP + "PROCType: " + TK.mPROCType + vbCrLf
    TMP = TMP + "WinVersion: " + TK.mWinVersion + vbCrLf
    
    Terr.Anotar "sak"
    Dim H As New HDSN, hC As Long
    
    For hC = 0 To 3
        Terr.Anotar "sal", hC
        H.CurrentDrive = hC
        If H.GetSerialNumber <> "" Then
            TMP = TMP + "Disk SN(" + CStr(hC) + "):" + H.GetSerialNumber + vbCrLf
            TMP = TMP + "Disk FR(" + CStr(hC) + "):" + H.GetFirmwareRevision + vbCrLf
            TMP = TMP + "Disk MN(" + CStr(hC) + "):" + H.GetModelNumber + vbCrLf
            FINAL = FINAL + H.GetSerialNumber + "|"
            Terr.Anotar "sam", H.GetSerialNumber
        End If
    Next hC
    
    Terr.Anotar "san"
    TMP = TMP + "FINAL: " + FINAL
    mResFinal = FINAL
    Terr.Anotar "sao", FINAL
    'PR.ShowProgreso "Escribiendo", 99
    GetIM = TMP
    'PR.HideProgreso
    
    Exit Function
ErrIM:
    Terr.AppendLog "errIM:", Terr.ErrToTXT(Err)
    
End Function

