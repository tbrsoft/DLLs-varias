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

Public HechoEnPixlesAncho As Long
Public HechoEnPixlesAlto As Long
Public HechoParaPixlesAncho As Long
Public HechoParaPixlesAlto As Long

Private OBS_Padres() As New ObjFullPadre  'todos los contenedores de objetos gráficos
Private OBS() As New objFULL  ' As New tbrSuperGuiObject 'OBjetoS

'deberia haber un archivo de configuracion co esto o un archivo parte del skin!

Public Function AddPadre(sName As String) As ObjFullPadre
    Dim k As Long
    k = UBound(OBS_Padres) + 1
    ReDim Preserve OBS_Padres(k)
    'Set OBS_Padres(k) = New ObjFullPadre
    OBS_Padres(k).sName = sName
    OBS_Padres(k).sHDC = -1 'bandera no inicializado
    
    Set AddPadre = OBS_Padres(k)
End Function


Public Function GetPadre(sName As String) As ObjFullPadre
    Dim k As Long
    For k = 1 To UBound(OBS_Padres)
        If LCase(OBS_Padres(k).sName) = LCase(sName) Then
            Set GetPadre = OBS_Padres(k)
            Exit Function
        End If
    Next k
    
    terr.AppendLog "NoFTH"
    Set GetPadre = OBS_Padres(0) 'vacio, error, seguiraqui, marcar el error
    
End Function

Public Function AddSgo(sNamePadre As String, sName As String) As objFULL
    Dim k As Long
    k = UBound(OBS) + 1
    ReDim Preserve OBS(k)
    'Set OBS(k) = New objFULL
    'SEGUIRAQUI deberia verificar que no existan dos con el mismo nombre
    Dim SGO As New tbrSuperGuiObject
    Set OBS(k).oSimple = SGO
    OBS(k).oSimple.SGOName = sName
    Set OBS(k).Padre = GetPadre(sNamePadre)
    
    'el objeto del manu se asigna despues
    
    Set AddSgo = OBS(k)
    
End Function

Public Function GetSgoMaxID() As Long
    GetSgoMaxID = UBound(OBS)
End Function

Public Function GetSgoByID(Ix As Long) As objFULL
    If Ix < 0 Or Ix > UBound(OBS) Then
        Set GetSgoByID = OBS(0)
    Else
        Set GetSgoByID = OBS(Ix)
    End If
End Function

Public Function getSGO(sName As String) As objFULL
    Dim k As Long
    For k = 1 To UBound(OBS)
        If Trim(LCase(OBS(k).oSimple.SGOName)) = Trim(LCase(sName)) Then
            Set getSGO = OBS(k)
            Exit Function
        End If
    Next k
    
    Set getSGO = OBS(0) 'para que no de error doy uno vacio
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
        OB.Y2 = CLng(CDbl(OB.y) * coefH)
        OB.W2 = CLng(CDbl(OB.W) * coefW)
        OB.H2 = CLng(CDbl(OB.H) * coefH)
    Else
        'si no es estirable debe correr su x e y para acomodarse y seguir mas o menos centrado en su zona
        OB.X2 = CLng(CDbl(OB.x) * coefW)
        OB.Y2 = CLng(CDbl(OB.y) * coefH)
        
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

'se necesita asignar ,un objeto del manu a cada objeto de los que administra esta clase
'necesito un proveedor de objetos del manu para cada tipo

Public Function GetNewManu(sTipo As String) As Object

    Dim x As Object 'valor de retorno
    
    Select Case Trim(LCase(sTipo))
        Case "manu":            Set x = New clsMANU
        Case "label":           Set x = New clsLabel
        Case "numerik":         Set x = New clsNumeriks
        Case "pngboton":        Set x = New clsPNGBoton
        Case "lista":           Set x = New clsTemasManager 'coleccion de temasList
        Case "listah":          Set x = New tbrTextoSelect
        Case "promociones2":    Set x = New tbrPromociones2 'textos que van cambiando
        Case "discomanager":    Set x = New clsDiscoManager
        
        
        Case Else
            
            'error, no se de que tipo SEGUIRAQUI
    
    End Select
    
    Set GetNewManu = x
    
End Function

'QUEDEAQUI estoy mareado para saber donde va cada cosa
Public Function setPropertyManu(nProp As String, nValue As String)
    
    
End Function

Private Sub Class_Initialize()
    ReDim OBS(0)
    ReDim OBS_Padres(0)
End Sub

'mostrar todo lo que tiene
Public Function toString() As String
    Dim M As Long, TMP As String
    TMP = ""
    For M = 1 To UBound(OBS)
        TMP = TMP + "OBS[" + CStr(M) + "]" + vbCrLf + _
                    OBS(M).oSimple.toString + vbCrLf + _
                    " * Padre: " + OBS(M).Padre.sName + " hdc=" + CStr(OBS(M).Padre.sHDC) + vbCrLf + _
                    " * Tipo " + OBS(M).Tipo
        If OBS(M).oManu Is Nothing Then
            TMP = TMP + "oManu = NOTHING"
        Else
            If TypeOf OBS(M).oManu Is tbrTextoSelect Then TMP = TMP + "oManu = tbrTextSelect"
            If TypeOf OBS(M).oManu Is tbrTextoMovil Then TMP = TMP + "oManu = tbrTextoMovil"
        End If
        
        TMP = TMP + vbCrLf + vbCrLf
    
    Next M
    
    toString = TMP
End Function
