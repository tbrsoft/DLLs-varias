Attribute VB_Name = "Globales"
Public tbrPintaNoPix As tbrPintar
'Public Type RECT
'  qLeft As Long
'  qTop As Long
'  qRight As Long
'  qBottom As Long
'End Type

Public Function dcr(st As String) As String
    Dim d As New tbrCrypto.Crypt
    Dim d2 As String
    d2 = d.DecryptString(eMC_Blowfish, st, "sargiotto", True)
    dcr = d2
End Function


