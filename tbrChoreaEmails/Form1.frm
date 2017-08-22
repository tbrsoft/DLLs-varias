VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   780
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   615
      Left            =   2310
      TabIndex        =   1
      Top             =   5850
      Width           =   3795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Levantar emails de un texto cualquiera"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New Scripting.FileSystemObject

Private Function DetectarEmails(Texto As String) As String()
    
    Dim Ret() As String
    ReDim Ret(0)
    
    Dim B As Long, LastB As Long
    LastB = 1
    Do
        B = InStr(LastB, Texto, "@", vbTextCompare)
        If B = 0 Then Exit Do
        
        LastB = B + 1
        
        'se encontro un email
        'ver donde empieza y donde termina
        Dim Ini As Long, FIN As Long 'puntos donde empeiza y termina el email
        Ini = GetPosChars(Texto, ",; |123()[]" + Chr(34), False, B)
        FIN = GetPosChars(Texto, ",; |123()[]" + Chr(34), True, B)
        
        Dim EsteMail As String
        EsteMail = Mid(Texto, Ini + 1, FIN - Ini - 1)
        
        Dim K As Long
        K = UBound(Ret) + 1
        ReDim Preserve Ret(K)
        'quitar carcteres molestos
        EsteMail = Replace(EsteMail, "'", "")
        
        Ret(K) = EsteMail
    Loop
    
    DetectarEmails = Ret
    
End Function

Private Function GetPosChars(TT As String, Seps As String, Adelante As Boolean, StartOn As Long) As Long

    Dim C1 As Long, C2 As Long, cMax As Long
    Dim SP() As String  'posibles separadores
    
    ReDim Preserve SP(Len(Seps))
    Dim Letra As String
    For C1 = 1 To UBound(SP)
        Letra = Mid(Seps, C1, 1)
        If Letra = "1" Then Letra = vbCr
        If Letra = "2" Then Letra = vbLf
        If Letra = "3" Then Letra = vbCrLf
        SP(C1) = Letra
    Next C1
    
    If Adelante Then
        cMax = StartOn + 64
    Else
        cMax = -1
    End If
    
    For C1 = 1 To UBound(SP)
        If Adelante Then
            C2 = InStr(StartOn, TT, SP(C1), vbTextCompare)
            If C2 < cMax And C2 > 0 Then cMax = C2 'marco la mejor posicion en el texto donde empezr a cortar el email
        Else
            C2 = InStrRev(TT, SP(C1), StartOn, vbDatabaseCompare)
            If C2 > cMax And C2 > 0 Then cMax = C2 'marco la mejor posicion en el texto donde empezr a cortar el email
        End If
    
    Next C1
    
    GetPosChars = cMax 'si es -1 no vale
End Function

Private Sub Command1_Click()
    
    Dim CM As New CommonDialog
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    Dim TE As TextStream, R As String
    Set TE = FSO.OpenTextFile(F, ForReading, False)
        If TE.AtEndOfStream Then
            MsgBox "El archivo esta vacio"
            Exit Sub
        End If
        
        R = TE.ReadAll
    
    TE.Close
    
    'ahora detectar todos los emails en el archivo sin importar su formato
    'esto es buscar los arrobas y detectar el espacio o ";" o "," o vbcr o vblf antes y despues
    Dim MLS() As String
    MLS = DetectarEmails(R)
    
    MsgBox "Se encontaron " + CStr(UBound(MLS)) + " emails"
    
    'ya tengo la lista definir cuantos rebotados les voy a sumar
    Dim K As Long, s As String
    
    For K = 1 To UBound(MLS)
        List1.AddItem MLS(K)
        Me.Caption = List1.ListCount
        Me.Refresh
    Next K
End Sub

Private Sub Command2_Click()
    
    Dim CM As New CommonDialog
    
    CM.ShowSave
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    Dim TE As TextStream
    Dim K As Long
    Set TE = FSO.OpenTextFile(F, ForWriting, True)
        
        TE.WriteLine List1.List(0)
        
        For K = 1 To List1.ListCount - 1
            If LCase(List1.List(K)) <> LCase(List1.List(K - 1)) Then
                TE.WriteLine List1.List(K)
            End If
        Next K
    
        
    TE.Close
    
End Sub
