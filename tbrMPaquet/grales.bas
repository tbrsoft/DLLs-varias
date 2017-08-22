Attribute VB_Name = "grales"
Private FSo As New Scripting.FileSystemObject

Public Sub esperar(n As Single)
    n = Timer + n
    Do While Timer < n
        DoEvents
    Loop
End Sub

Public Sub myCopyFolder(org As String, dest As String, overWrite As Boolean)
    'copy folder ecesita QUE NO HAYA barras al final de los paths y simepre lo uso CON basrras al final
    
    Dim org2 As String
    Dim dest2 As String
    
    If Right(org, 1) = "\" Then
        org2 = Mid(org, 1, Len(org) - 1)
    Else
        org2 = org
    End If
    
    If Right(dest, 1) = "\" Then
        dest2 = Mid(dest, 1, Len(dest) - 1)
    Else
        dest2 = dest
    End If
    
    FSo.CopyFolder org2, dest2, overWrite
    
End Sub

Public Function txtInLista(lista As String, Orden As Long, Separador As String) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    'si pongo 99999 en orden saco el ultimo
    Dim lAct As String, lOrden As Integer
    Dim palabra(40) As String
    Dim C As Integer
    C = 1: lOrden = 0
    Do While C <= Len(lista)
        lAct = Mid(lista, C, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
            If lOrden > Orden Then Exit Do
        End If
        C = C + 1
    Loop
    'si oreden solicitado>ultimo oreden de la lista...
    If Orden > lOrden Then
        If Orden = 99999 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden): Exit Function
        End If
        If Orden = 99998 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden - 1): Exit Function
        End If
        If Orden <> 99999 And Orden <> 99998 Then
            txtInLista = "OUT LISTA": Exit Function
        End If
    End If
    txtInLista = palabra(Orden)
End Function

