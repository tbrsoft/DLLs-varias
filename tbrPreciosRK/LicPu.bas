Attribute VB_Name = "LicPu"
Public Function tbrFIX(n As Single, DecimalesTruncar As Long) As Single
    'truncar a una X cantidad de decimales
    Dim SN As String
    'tratarlo como caracter es mas facil
    SN = CStr(n)
    'si es entero entonces salgo, no hay nada que hacer
    Dim TieneDec As Boolean
    If InStr(SN, ",") > 0 Then TieneDec = True
    If InStr(SN, ".") > 0 Then TieneDec = True
    If TieneDec = False Then
        tbrFIX = n
        Exit Function
    End If
    
    Dim AA As Long, Largo As Long, BB As Long
    BB = 0 'cuenta la cantidad de decimales
    Largo = Len(SN)
    Dim EmpezoDec As Boolean
    EmpezoDec = False
    For AA = 1 To Largo
        If EmpezoDec Then BB = BB + 1
        'si se llega al total cortar ahi
        If BB = DecimalesTruncar Then
            tbrFIX = CSng(Mid(SN, 1, AA))
            Exit Function
        End If
        If Mid(SN, AA, 1) = "." Or Mid(SN, AA, 1) = "," Then EmpezoDec = True
    Next AA
    'si sale de aqui sin haber salido antes es porque no llega a la cantida deseada
    tbrFIX = n
End Function

