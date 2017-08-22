VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function GetClaves(IdTest As String, Var As Long)
    'paso un ID y genero las claves
    'Var es la variación
    'la cantidad de carcteres de respuesta debe ser variable para complicarlo mas
    Dim CV As String
    CV = "GHEFMRTW34NP68XYQJKAC29MTR43HGYXP68KKA"
    Dim TmpRes As String, L As String 'L es cada letra
    Dim TmpVal As Long
    'son 35 caracteres HEX + "H" + "-"
    Dim A As Long
    TmpRes = ""
    
    For A = 1 To 35
        L = Mid(IdTest, A, 1)
        If L = "-" Or L = "H" Then GoTo sigLetra
        '1F70-4F5F-HE-6H-0009HDAC11-H041H060
        Select Case L
            Case "1", "F"
                TmpVal = 410 + (Var * A)
            Case "2", "E"
                TmpVal = (1410 + Var) * A
            Case "3", "D"
                TmpVal = Abs(Var - A) * 5
            Case "4", "C"
                TmpVal = Abs(9 * A + 488 - Var)
            Case "5", "B"
                TmpVal = Abs(-A - Var * 19)
            Case "6", "A"
                TmpVal = Abs(Var * A - 19)
            Case "7", "0"
                TmpVal = Abs(Var * 2 + 511 - A)
            Case "8"
                TmpVal = Abs((-5 * A) * Var - 19)
            Case "9"
                TmpVal = Abs(A + Var - 19)
        End Select
        'asegurarse que este dentro de len del CV
        If TmpVal = 0 Then TmpVal = 1
        If TmpVal > Len(CV) Then TmpVal = (TmpVal Mod Len(CV)) + 1 'cero no sirve!
        'como es el resto y el divisor es el maximo no se pasa nunca del Len(CV)
        TmpRes = TmpRes + Mid(CV, TmpVal, 1)
        'poner guiones
        If Len(TmpRes) / 6 = Len(TmpRes) \ 6 Then TmpRes = TmpRes + "-"
    
sigLetra:

    Next A
    
    GetClaves = TmpRes
    
End Function

