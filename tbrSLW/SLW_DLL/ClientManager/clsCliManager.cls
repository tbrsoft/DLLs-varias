VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CN As New ADODB.Connection

Private Sub Class_Initialize()
    ConectarBases
End Sub

Private Sub ConectarBases()
    Dim DireccionBase As String
    DireccionBase = AP + "db\db.mdb"
    
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        DireccionBase + ";Persist Security Info=False"
    CN.Open
End Sub

Public Function GetCredCli(idCli As Long, IdSoft As Long) As Long
    'obtener la cantidad de creditos de un cliente en determinado software
    
'    Dim Rs2 As New ADODB.Recordset
'    Rs2.CursorLocation = adUseClient
'
'    Rs2.Open "SELECT tblLicDispVariaciones.idCLI, tblLicDispVariaciones.idSoft, " + _
'        "Sum(tblLicDispVariaciones.variacion) AS SumaDevariacion From " + _
'        "tblLicDispVariaciones GROUP BY tblLicDispVariaciones.idCLI, " + _
'        "tblLicDispVariaciones.idSoft HAVING (((tblLicDispVariaciones.idCLI)=" + _
'        CStr(idCli) + ") " + _
'        "AND ((tblLicDispVariaciones.idSoft)=" + _
'        CStr(idSoft) + "))", CN, adOpenStatic, adLockReadOnly
'
'
'    GetCredCli = Rs2.Fields("SumaDevariacion")
    
    Dim Rs2 As New ADODB.Recordset
    Rs2.CursorLocation = adUseClient

    Rs2.Open "select * from tblLicDispVariaciones " + _
        "where idcli=" + CStr(idCli) + " and " + _
        "idsoft=" + CStr(IdSoft), CN, adOpenStatic, adLockReadOnly

    If Rs2.RecordCount = 0 Then
        GetCredCli = 0 'nunca tuvo nada!
        Exit Function
    End If

    If Rs2.RecordCount >= 1 Then
        'pasar uno por uno validando var2
        Rs2.MoveFirst
        Dim ACU As Long: ACU = 0
        Dim IC As Long 'id cli
        Dim ISo As Long 'id soft
        Dim V1 As Long 'var pedida
        Dim V2 As Long 'validacion
        
        Do While Not Rs2.EOF
            IC = Rs2.Fields("idcli")
            ISo = Rs2.Fields("idsoft")
            V1 = Rs2.Fields("variacion")
            V2 = Rs2.Fields("var2")
            '(idC - var) * (IdS - var)
            If V2 = (IC - V1) * (ISo - V1) Then ACU = ACU + V1
            Rs2.MoveNext
        Loop
        GetCredCli = ACU 'resolver quilombo!
        Exit Function
    End If

    If Rs2.RecordCount < 0 Then
        GetCredCli = -3 'resolver quilombo!
        Exit Function
    End If

End Function

Public Sub RestarCredito(IDC As Long, IdSoft As Long)
        Dim OBS2 As String
        Dim NIDT As Long
        
        NIDT = GetNextIdTRAM
        OBS2 = "L"
        
        'lo vacuno nomás
        If OBS2 = "SL" Then VarCredCli IDC, IdSoft, -2, NIDT
        If OBS2 = "L" Then VarCredCli IDC, IdSoft, -1, NIDT
End Sub

Private Sub VarCredCli(IDC As Long, idS As Long, var As Long, IdTra As Long)
    'registrar variaciones en
    
    Dim CRANT As Long
    CRANT = GetCredCli(IDC, idS)
    
    Dim RsT As New ADODB.Recordset
    RsT.CursorLocation = adUseClient
    
    RsT.Open "select * from tblLicDispVariaciones", CN, adOpenStatic, adLockOptimistic
    RsT.AddNew
    RsT.Fields("idcli") = IDC
    RsT.Fields("idsoft") = idS
    RsT.Fields("idtra") = IdTra
    RsT.Fields("variacion") = var
    'la validacion depende de los 3 numeros anteriores
    RsT.Fields("var2") = (IDC - var) * (idS - var)
    RsT.Fields("crant") = CRANT
    
    RsT.Update
    
End Sub

Public Function GetNextIdTRAM() As Long
    'obtener un numero de tramite disponible
    'creando el registro
    
    Dim RsT As New ADODB.Recordset
    RsT.CursorLocation = adUseClient
    RsT.Open "select top 1 * from tbltras order by tbltras.id desc", CN, adOpenStatic, adLockOptimistic
    
    Dim LastID As Long
    If RsT.RecordCount <= 0 Then
        LastID = 0
    Else
        RsT.MoveFirst
        LastID = RsT.Fields("id")
    End If
    
    LastID = LastID + 1
    
    RsT.AddNew
    RsT.Fields("momento") = Now
    RsT.Fields("id") = LastID
    RsT.Update
    
    GetNextIdTRAM = LastID
    
    RsT.Close
End Function

