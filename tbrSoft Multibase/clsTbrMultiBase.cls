VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrMultiBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'sistema para traducir entre multiples bases

Private elBASE() As String

Public Function GetD(strIN As String, BaseIn As Long, BaseOut As Long) As String
    'si base in es cero quiere decir que mando un texto plano de entrada
    
End Function

'se puede definir una nueva base (cuenta la cantidad de elementos sola) que sera por ejemplo
'para los elementos 0,1,3 será la base -2 (es base 2 pero negativa por qwue la cree yo)
Public Sub defineBase(strElements As String)

    If Len(strElements) > UBound(elBASE) Then
        ReDim Preserve elBASE(Len(strElements))
    End If
    
    elBASE(Len(strElements)) = strElements

End Sub

Private Sub Class_Initialize()
    ReDim elBASE(0)
End Sub
