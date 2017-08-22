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

Private OBS_Padres() As New ObjFullPadre  'todos los contenedores de objetos gráficos

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

Public Function GetPadresMaxID() As Long
    GetPadresMaxID = UBound(OBS_Padres)
End Function

Public Function GetPadresByID(Ix As Long) As ObjFullPadre
    If Ix < 0 Or Ix > UBound(OBS_Padres) Then
        Set GetPadresByID = OBS_Padres(0)
    Else
        Set GetPadresByID = OBS_Padres(Ix)
    End If
End Function

Private Sub Class_Initialize()
    Clear
End Sub

Public Sub Clear()
    ReDim OBS_Padres(0)
End Sub
