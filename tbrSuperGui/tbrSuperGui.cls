VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrSuperGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase que guarda las coordeadas de todos los objetos visibles
'ademas las adapta a la pantalla

'esta separado de clsGui para que el otro cargue desde archivos y esta clase se mantenga limpia de eso
'y de las funciones que el software pide

Public HechoParaPixlesAncho As Long
Public HechoParaPixlesAlto As Long

Public HechoEnPixlesAncho As Long
Public HechoEnPixlesAlto As Long

Public Type ObjFULL 'suma de un objeto basico mio y el del manu correspondiente
    oSimple As tbrSuperGuiObject  'OBjetoS
    oManu As Object 'varia segun el tipo del manu que sea
    Tipo As String 'tipo del objeto del manu
End Type

Private OBS() As ObjFULL ' As New tbrSuperGuiObject 'OBjetoS

'deberia haber un archivo de configuracion co esto o un archivo parte del skin!

Public Function AddSgo(sName As String) As ObjFULL
    Dim k As Long
    k = UBound(OBS) + 1
    ReDim Preserve OBS(k)
    
    'SEGUIRAQUI deberia verificar que no existan dos con el mismo nombre
    Dim SGO As New tbrSuperGuiObject
    Set OBS(k).oSimple = SGO
    OBS(k).oSimple.SGOName = sName
    'el del manu se asigna despues
    
    AddSgo = OBS(k)
    
End Function

Public Function GetSgoByID(ix As Long) As ObjFULL
    If ix < 0 Or ix > UBound(OBS) Then
        GetObjByID = OBS(0)
    Else
        GetObjByID = OBS(ix)
    End If
End Function

Public Function gets()

End Function

Public Function getSGO(sName As String) As ObjFULL
    Dim k As Long
    For k = 1 To UBound(OBS)
        If Trim(LCase(OBS(k).oSimple.SGOName)) = Trim(LCase(sName)) Then
            getSGO = OBS(k)
            Exit Function
        End If
    Next k
    
    getSGO = OBS(0) 'para que no de error doy uno vacio
End Function

Public Sub TraslateALL()
    Dim k As Long
    For k = 1 To UBound(OBS)
        TraslateToThisScreen OBS(k).oSimple
    Next k
End Sub

Private Sub TraslateToThisScreen(OB As tbrSuperGuiObject)
    On Local Error Resume Next
    
    Dim coefW As Double
    Dim coefH As Double
    
    coefW = (HechoEnPixlesAncho / 15) / HechoParaPixlesAncho
    coefH = (HechoEnPixlesAlto / 15) / HechoParaPixlesAlto
    
    'aplicar a "2"
    
    
    If OB.Estirable Then 'joiaaaaaaaaaaaaa
        OB.X2 = CLng(CDbl(OB.x) * coefW)
        OB.Y2 = CLng(CDbl(OB.Y) * coefH)
        OB.W2 = CLng(CDbl(OB.W) * coefW)
        OB.H2 = CLng(CDbl(OB.H) * coefH)
    Else
        'si no es estirable debe correr su x e y para acomodarse y seguir mas o menos centrado en su zona
        OB.X2 = CLng(CDbl(OB.x) * coefW)
        OB.Y2 = CLng(CDbl(OB.Y) * coefH)
        
        OB.W2 = OB.W
        OB.H2 = OB.H
        
        'ver que alineacion tiene!
        Dim ExtraX As Long
        Dim ExtraY As Long
        
        'alineacion centrada predeterminada
        'movieminto extra (para que quede centrado en la zona que ocuparia el objeto si fuera estirable. S O Y   U N   G E N I O
        If OB.AlignementH = 1 Then ExtraX = ((CLng(CDbl(OB.W) * coefW)) - OB.W) / 2
        If OB.AlignementV = 1 Then ExtraY = ((CLng(CDbl(OB.H) * coefH)) - OB.H) / 2
        'izquierda en H y arriba en V
        'If ob.AlignementH = 0 Then 'izquierda (nada, ya esta ahi!!)
        'If ob.AlignementV = 0 Then 'arriba (nada, ya esta ahi!!)
        'derecha en H y abajo en Y
        If OB.AlignementH = 2 Then ExtraX = ((CLng(CDbl(OB.W) * coefW)) - OB.W)
        If OB.AlignementV = 2 Then ExtraY = ((CLng(CDbl(OB.H) * coefH)) - OB.H)
        
        OB.X2 = OB.X2 + ExtraX
        OB.Y2 = OB.Y2 + ExtraY
        
    End If
    
    '... y ya se puede usar x2, y2, etc
End Sub

'se mnecesita asignar ,un objeto del manu a cada objeto de los que administra esta clase
'necesito un proveedor de objetos del manu para cada tipo

Public Function GetNewManu(sTipo As String) As Object
    Select Case Trim(LCase(sTipo))
        Case "manu"
            Dim x As New clsMANU
            
        Case Else
            'error, no se de que tipo
    
    End Select
    
    Set GetNewManu = x
    
End Function

Private Sub Class_Initialize()
    ReDim OBS(0)
End Sub
