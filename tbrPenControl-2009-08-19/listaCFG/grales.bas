Attribute VB_Name = "grales"
Private FSo As New Scripting.FileSystemObject
Public terr As New tbrErrores.clsTbrERR
'variables usadas'contador reiniciable e historico respectivamente
'antes era de ctl main de todo pero necesito traducir en otros lugares!
Public VS2 As New clsVARS

Public Sub esperar(n As Single)
    n = Timer + n
    Do While Timer < n
        DoEvents
    Loop
End Sub

Public Sub myCopyFolder(org As String, dest As String, overWrite As Boolean)
    'copy folder ecesita QUE NO HAYA barras al final de los paths y simepre lo uso CON basrras al final
    On Local Error GoTo NOCOPY
    
    terr.Anotar "qaa", org, dest, overWrite
    
    Dim org2 As String
    Dim dest2 As String
    
    If Right(org, 1) = "\" Then
        org2 = mID(org, 1, Len(org) - 1)
    Else
        org2 = org
    End If
    
    If Right(dest, 1) = "\" Then
        dest2 = mID(dest, 1, Len(dest) - 1)
    Else
        dest2 = dest
    End If
    
    terr.Anotar "qac", org, dest, overWrite
    FSo.CopyFolder org2, dest2, overWrite
    
    Exit Sub
NOCOPY:
    terr.AppendLog "qab", terr.ErrToTXT(Err)
    Resume Next
End Sub
