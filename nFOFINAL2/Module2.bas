Attribute VB_Name = "Module1"
'Todas las cosas que estan aqui son unicas para cualquier dll de estas
'como en los licencieros hay dos de estos en ese caso el terr es el mismo para los dos

Public Terr As New tbrErrores.clsTbrERR
Public AP As String

'desencriptar cadenas para o se vean al decompilar ele ejecutable
Public Function dcr(sT As String) As String
    Dim D As New tbrCrypto.Crypt
    Dim d2 As String
    d2 = D.DecryptString(eMC_Blowfish, sT, "Cerrar sistema", True)  'cuh1000v
    dcr = d2
End Function



