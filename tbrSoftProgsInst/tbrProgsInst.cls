VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrProgsInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private IntString As String

Private Type Prog
    Nombre As String
    Ubicacion As String
    HelpLink As String 'pagina web por ejemplo
    InstallDate As String
    Publisher As String
    URLInfoAbout As String
    URLUpdateInfo As String
End Type

Private mProgs() As Prog

Public Function GetPath2(LikeNombre As String, LikePublisher As String) As String

    'puede servir buscar carpetas en base a datos del nombre y la empresa que lo crea
    'al primero que encuentra sale
    Dim J As Long
    For J = 1 To UBound(mProgs)
        If InStr(1, mProgs(J).Nombre, LikeNombre, vbTextCompare) Then
            If InStr(1, mProgs(J).Publisher, LikePublisher, vbTextCompare) Then
                GetPath2 = mProgs(J).Ubicacion
                Exit Function
            End If
        End If
    Next J
End Function


Public Function GetURLUpdateInfo(i As Long)
    If i > UBound(mProgs) Then
        GetURLUpdateInfo = ""
    Else
        GetURLUpdateInfo = mProgs(i).URLUpdateInfo
    End If
End Function


Public Function GetURLInfoAbout(i As Long)
    If i > UBound(mProgs) Then
        GetURLInfoAbout = ""
    Else
        GetURLInfoAbout = mProgs(i).URLInfoAbout
    End If
End Function

Public Function GetPublisher(i As Long)
    If i > UBound(mProgs) Then
        GetPublisher = ""
    Else
        GetPublisher = mProgs(i).Publisher
    End If
End Function

Public Function GetInstallDate(i As Long)
    If i > UBound(mProgs) Then
        GetInstallDate = ""
    Else
        GetInstallDate = mProgs(i).InstallDate
    End If
End Function

Public Function GetHelpLink(i As Long)
    If i > UBound(mProgs) Then
        GetHelpLink = ""
    Else
        GetHelpLink = mProgs(i).HelpLink
    End If
End Function

Public Function GetPath(i As Long)
    If i > UBound(mProgs) Then
        GetPath = ""
    Else
        GetPath = mProgs(i).Ubicacion
    End If
End Function

Public Function GetName(i As Long)
    If i > UBound(mProgs) Then
        GetName = ""
    Else
        GetName = mProgs(i).Nombre
    End If
End Function

Public Function Cantidad() As Long
    Cantidad = UBound(mProgs)
End Function

Public Sub LoadList()
    
    Dim Icnt As Integer
    IntString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
    REG32.GetKeyNames HKEY_LOCAL_MACHINE, IntString
    
    Dim N As String, C As Long
    For Icnt = 1 To sKeys.Count - 1
        N = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "DisplayName")
        If Len(N) > 0 Then
            C = C + 1
            ReDim Preserve mProgs(C)
            mProgs(C).Nombre = N
            mProgs(C).Ubicacion = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "InstallLocation")
            
            mProgs(C).HelpLink = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "HelpLink")
            mProgs(C).InstallDate = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "InstallDate")
            mProgs(C).Publisher = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "Publisher")
            mProgs(C).URLInfoAbout = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "URLInfoAbout")
            mProgs(C).URLUpdateInfo = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "URLUpdateInfo")
        End If
    Next
    
    'Set sKeys = Nothing
    
End Sub

Private Sub Class_Initialize()
    ReDim mProgs(0)
    Set sKeys = New Collection
    LoadList
End Sub
