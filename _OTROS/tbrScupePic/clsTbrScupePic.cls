VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrScupePic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim PP As Object

'puede escupir pintura en algun objeto que lo permita (frm, pic,etc)

Public Sub AeroSol(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, _
    lColor As OLE_COLOR, Ancho As Long)
    'hace una linea con aerosol de color especificado

End Sub

Public Sub Levantar3D(X1 As Single, Y1 As Single, _
                      X2 As Single, Y2 As Single, _
                      X3 As Single, Y3 As Single, _
                      Direccion As Long)
    
    'corre los pixeles en una linea para simular inclinacion
    'x1 e y1 son los puntos de inicio
    'x2 e y2 son los puntos de fin de lo que se va a pintar
    'x3 e y3 son los puntos de fin de la zona que se va a representar hasta
    
    'direccion indica para que lado se corre
        '0= de izquierda a derecha (y3 debe ser igual a y2)
        '1= de derecha a izquierda (y3 debe ser igual a y2)
        '2= de arriba a bajo (x3 debe ser igual a x2)
        '3= de abajo a arriba (x3 debe ser igual a x2)
    
    
    Dim I As Long, J As Long
    Dim CO As Long
    Dim Dif As Long, Coef As Long
    Select Case Direccion
        Case 0
            For I = Y1 To Y2 Step 15 'por que es en twips
                CO = 0
                Coef = (X3 - X1) / (X2 - X1)
                For J = X1 To X1 + X2 Step 15  'la primera mitad
                    CO = CO + 1
                    PP.PSet (J, I), PP.Point(J + (15 * Coef * CO), I)
                Next J
                
'                For J = X1 + ((X2 - X1) / 2) To X2 Step 15
'                    PP.PSet (J, I), 0
'                Next J
                
            Next I
            
            
    End Select
End Sub

Public Sub SetPP(x As Object)
    Set PP = x
End Sub
