Attribute VB_Name = "Module1"

Public Function EnumEncoding(ByVal nStatus As Integer) As Boolean
    
    'set probar and label %'s to update with the control
    
    Form1.lblPercent.Caption = Format(nStatus, "00") & "%"
    Form1.Label1.Caption = Format(nStatus, "00") & "%"
    Form1.prog.Value = nStatus
    'do the above
    DoEvents
    EnumEncoding = True
    
End Function
