VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada una de las traducciones

Public Finaltext As String
Public RefIdioma As String     'nombre del idioma que representa
Public fechaTrans As Long     'ni en pedo uso fechas, solo el dia como long
Public Terminada As Boolean
Public NoTerminada As Boolean     'predeterminado es true, la marca asi si no entiende o si le faltan buscar palabras
Public NoEntiendo As Boolean     'me marca que revise algo, en obs pondra que es
Public FechaBaseNewer As Boolean     'solo lectura, cuando lo abro me fijo si la fecha de la base es superior a la de la traduccion
Public Observaciones As String     'texto donde el traductor dice que cosa no entiende

