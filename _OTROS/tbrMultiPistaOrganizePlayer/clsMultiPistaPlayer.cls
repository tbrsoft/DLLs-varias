VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiPistaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'maneja la reproduccion de tbrPlayer siempre teniendo en cuenta los enganches, _
    los tipos de archivos la duracion, si es gratis el volumen, etc, etc
    
Private mMiliSegundosEnganche As Long 'para cada separaion de canciones

Private WithEvents MMP As tbrPlayer.MainPlayer
Attribute MMP.VB_VarHelpID = -1
Private LST As tbrListaRep.clsListaRep 'lista que se carga desde 3pm y se _
    descarga desde aqui

Private Sub MMP_BeginPlay(iAlias As Long)
    
End Sub

Private Sub MMP_Played(SecondsPlayed As Long, iAlias As Long)

End Sub

Public Sub Iniciar(ByRef tList As tbrPlayer.MainPlayer)
    'ver si hay algo en la lista
    
    Set LST = tList 'engancho la lista del 3PM a esta, desde el 3pm _
        le agrego y desde aqui le quito.
        
    If LST.GetLastIndex = 0 Then Exit Sub
    'aqui hay canciones
    
    
    
End Sub

Private Sub Class_Initialize()
    Set MMP = New tbrPlayer.MainPlayer
    
End Sub

Public Function CancelarEjecucion()
    
End Function
