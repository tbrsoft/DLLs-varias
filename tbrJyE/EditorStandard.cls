VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditorBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Nuevo()
Public Event Abrir(elArchivo As String)
Public Event Grabar(elArchivo As String)
Public Event Salir()

Public Modificado As Boolean
Public MiPath As String 'Archivo con el que trabajo actualmente

Dim CD As New CommonDialog
Dim qExt As String
Dim qRec As String

'laextencion son solo 3 letras ej: bmp
Public Sub Iniciar(laExtencion As String, NombreProg As String)
    qExt = laExtencion
    qRec = NombreProg
End Sub

Public Sub Abrir()
    Dim Carp As String
    Carp = GetSetting("Editor", "Base", qRec, "C:\")
    CD.InitDir = Carp
    
    CD.FileName = ""
    CD.Filter = "Archivos ." + qExt + "|*." + qExt
    CD.DialogTitle = "Abrir"
    CD.ShowOpen
    
    If CD.FileName <> "" Then
        If Dir(CD.FileName) <> "" Then
            MiPath = CD.FileName
            RaiseEvent Abrir(MiPath)
    
            Carp = Mid(CD.FileName, 1, InStrRev(CD.FileName, "\"))
            SaveSetting "Editor", "Base", qRec, Carp
        End If
    End If
End Sub

Public Sub GuardarComo()
    
    Dim Carp As String
    Carp = GetSetting("Editor", "Base", qRec, "C:\")
    CD.InitDir = Carp
    
    CD.FileName = ""
    CD.Filter = "Archivos ." + qExt + "|*." + qExt
    CD.DialogTitle = "Guardar"
    CD.ShowSave
    
    If CD.FileName <> "" Then
        If LCase(Right(CD.FileName, 4)) <> "." + qExt Then
            CD.FileName = CD.FileName + "." + qExt
        End If
        If Dir(CD.FileName) <> "" Then
            If MsgBox("Este archivo ya existe, quiere sobreescribirlo?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
                MiPath = CD.FileName
                Modificado = False
                Kill CD.FileName
                RaiseEvent Grabar(MiPath)
                Exit Sub
            Else
                Exit Sub
            End If
        End If
        MiPath = CD.FileName
        Modificado = False
        RaiseEvent Grabar(MiPath)
        Carp = Mid(CD.FileName, 1, InStrRev(CD.FileName, "\"))
        SaveSetting "Editor", "Base", qRec, Carp
    End If
End Sub

Public Sub Guardar()
    If MiPath = "" Then
        GuardarComo
    Else
        'GrabarArchivo MiPath
        Modificado = False
        RaiseEvent Grabar(MiPath)
    End If
End Sub

Public Sub Nuevo()
    If Modificado = True Then
        If MsgBox("Quiere guardar este archivo?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
            Guardar
            End
        End If
    End If
    MiPath = ""
    Modificado = False
    RaiseEvent Nuevo
End Sub

Public Sub Salir()
    If Modificado = True Then
        If MsgBox("Quiere guardar este archivo?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
            Guardar
            End
        End If
    End If
    RaiseEvent Salir
End Sub
