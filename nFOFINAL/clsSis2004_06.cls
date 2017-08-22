VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSis2004_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'aqui se puede saber segun una clave cuando se habilito y a quien

Private mCliente As String
Private mClave As String
Private mTipoL As String
Private mFecha As String

Public Sub SetClave(sArch As String, sClave As String, ByRef sCLI As String, _
    ByRef sTIPOL As String, ByRef sFECHA As String)
    
    mClave = sClave
    'ahi nomas obtener los datos de como se licencio
    
    Dim FS2 As New Scripting.FileSystemObject
    Dim TE As TextStream
    Dim HABS() As String 'cada una de las habilitaciones
    ReDim HABS(0) 'uso a partir del 1
    Dim i As Long
    'leo todos los registros
    'sArch = "D:\dev\DLL\nFOFINAL\listaSistCl2004-2006.txt"
    Set TE = FS2.OpenTextFile(sArch, ForReading, False)
        Do While Not TE.AtEndOfStream
            i = UBound(HABS) + 1
            ReDim Preserve HABS(i)
            HABS(i) = TE.ReadLine
        Loop
    TE.Close
    
    'ejemplo de renglon
    '683HC-DVMAA-3NAVK-eACVG-FzMMK-42EG3-2xE83-HTT42:Tomas Nuñez (Mex): _
        12/02/2007:SupL:23:13063-235405318:Mexico..E8078500070668.33300077: _
        226,5 antonio arteaga
    
    Dim SP() As String, Encontrado As Boolean
    
    Encontrado = False
    mCliente = ""
    mFecha = ""
    mTipoL = ""
    
    For i = 1 To UBound(HABS)
        SP = Split(HABS(i), ":")
        If SP(0) = mClave Then
            'encontre el que me piden!
            Encontrado = True
            
            mCliente = SP(1)
            mFecha = SP(2)
            mTipoL = SP(3)
            
            Exit For
        End If
    Next i
    
    'devuelvo toda aca y listo
    sCLI = mCliente
    sFECHA = mFecha
    sTIPOL = mTipoL
    
End Sub
