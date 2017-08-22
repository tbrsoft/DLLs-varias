VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cstTBRCFGs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'tipos de opciones
Public Enum TipoOpcion
    tipoOpcionString = 0
    tipoOpcionNumero = 1
    tipoOpcionFile = 2
End Enum

Private mFileCFG As String 'archivo de configuracion
Private vNamesConfigs() As String
Private vTiposConfigs() As Long
Private vDefValConfig() As String 'valor predeterminado
Private vValuesConfig() As String

Public Sub AddConfig(NameConfig As String, tipoDato As TipoOpcion, _
    Optional DefaultValue = "")
    
    Dim NewI As Long
    NewI = UBound(vNamesConfigs) + 1
    ReDim Preserve vNamesConfigs(NewI)
    ReDim Preserve vTiposConfigs(NewI)
    ReDim Preserve vDefValConfig(NewI)
    
    vNamesConfigs(NewI) = NameConfig
    vTiposConfigs(NewI) = tipoDato
    vDefValConfig(NewI) = 1
        
End Sub

Public Sub GrabarDefConfig()
    
End Sub

    
    
