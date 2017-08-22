Attribute VB_Name = "Module1"
Public Llamador As New clsLlamarEvento

Public tErr As New tbrErrores.clsTbrERR

'Public PcjeActual As Integer

Public Function EnumEncoding(ByVal nStatus As Integer) As Boolean
    'RaiseEvent Estado(nStatus)
    Llamador.CambiarPorcentaje nStatus
    
    'PcjeActual = nStatus
    'do the above
    
    ''''DoEvents sacado marzo 2010 para que no joda la interface USB !
    
    EnumEncoding = True
End Function

