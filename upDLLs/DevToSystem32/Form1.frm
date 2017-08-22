VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "actualizar"
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   7170
      Width           =   2715
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      IntegralHeight  =   0   'False
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   750
      Width           =   13305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Elegir de que carpeta (recursivo)"
      Height          =   465
      Left            =   8190
      TabIndex        =   3
      Top             =   90
      Width           =   2235
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   150
      Width           =   7845
   End
   Begin VB.TextBox LIST1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   7830
      Width           =   13245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listar encontrados"
      Height          =   555
      Left            =   10620
      TabIndex        =   0
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2850
      TabIndex        =   6
      Top             =   7110
      Width           =   10275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FOL As String
Dim RES() As String
Dim FSO As New Scripting.FileSystemObject
Dim ArchsOrig() As String 'encontrados para procesar despues
Dim ArchsDest() As String 'encontrados para procesar despues

Private Sub Command1_Click()
    
    Command1.Enabled = False
    'grabar esta como usada
    saveFileUsadas Combo1.Text
    
    Dim BR As New tbrPaths.clsPATHS
    
    BR.LeerTodo Combo1.Text, False, False   'no pongo parametro para que devuelva *.* y despues filtro a mano
    
    RES = BR.GetLista
    
    LIST1.Text = ""
    
    On Local Error GoTo errCOPY
    
    Dim J As Long, TMP As String
    Dim ii As Long
    ReDim ArchsOrig(0)
    ReDim ArchsDest(0)
    ii = 0
    List2.Clear
    For J = 1 To UBound(RES)
        TMP = RES(J)
        Dim RESP As Long
        If Right(TMP, 1) <> "\" Then
            If LCase(Right(TMP, 3)) = "dll" Or LCase(Right(TMP, 3)) = "ocx" Then
                
                Dim Orig As String, Dest As String, reemplazar As Boolean
                Orig = TMP
                Dest = "C:\Windows\System32\" + FSO.GetBaseName(TMP) + "." + FSO.GetExtensionName(TMP)
                reemplazar = False
                Dim Extras As String
                Extras = ""
                'si suma puntos dejar marcado!
                Dim PTs As Long
                PTs = 0
                If FSO.FileExists(Dest) = False Then
                    Extras = Extras + " [no existe destino]"
                    reemplazar = True
                    PTs = PTs + 3
                Else 'solo si existe se puede comparar !
                    'choreado del copiseg
                    
                    If FileDateTime(Orig) <> FileDateTime(Dest) Then
                        Extras = Extras + " [fecha: " + CStr(FileDateTime(Dest)) + " => " + CStr(FileDateTime(Orig)) + "]"
                        reemplazar = True
                        If FileDateTime(Orig) > FileDateTime(Dest) Then PTs = PTs + 1
                    End If
                    
                    If FileLen(Orig) <> FileLen(Dest) Then
                        Extras = Extras + " [tamaño: " + CStr(FileLen(Dest)) + " => " + CStr(FileLen(Orig)) + "]"
                        reemplazar = True
                        If FileLen(Orig) > FileLen(Dest) Then PTs = PTs + 1
                    End If
                    
                    If FSO.GetFileVersion(Orig) <> FSO.GetFileVersion(Dest) Then
                        Extras = Extras + " [vers: " + CStr(FSO.GetFileVersion(Dest)) + " => " + CStr(FSO.GetFileVersion(Orig)) + "]"
                        reemplazar = True
                        If FSO.GetFileVersion(Orig) > FSO.GetFileVersion(Dest) Then PTs = PTs + 1
                    End If
                End If
                
                
                If reemplazar Then
                    ReDim Preserve ArchsOrig(ii)
                    ReDim Preserve ArchsDest(ii)
                    ArchsOrig(ii) = Orig
                    ArchsDest(ii) = Dest
                    ii = ii + 1
                    Dim TXT As String
                    TXT = FSO.GetBaseName(TMP) + "." + FSO.GetExtensionName(TMP)
                    TXT = TXT + " [" + CStr(PTs) + "] " + Extras
                    List2.AddItem TXT
                    If PTs > 1 Then List2.Selected(List2.ListCount - 1) = True
                    List2.Refresh
                End If
            End If
        End If
    Next J
    
    MsgBox "Terminado"
    
    Command1.Enabled = True
    
    Exit Sub
    
errCOPY:
    LIST1.Text = LIST1.Text + "    **********" + CStr(Err.Number) + ": " + Err.Description + vbCrLf
    LIST1.Text = LIST1.Text + "    FALLO" + vbCrLf
    LIST1.Text = LIST1.Text + "    **********" + vbCrLf
    Resume Next
End Sub

Private Sub ReadUsadas()
    Combo1.Text = ""
    Dim TE As TextStream
    
    'ver si esta el texto! para no duplicar
    Dim r As String
    Set TE = FSO.OpenTextFile(App.path + "\usadas.txt", ForReading, True)
        
        Do While Not TE.AtEndOfStream
            r = TE.ReadLine
            'ver que exista, en la pc de los diferentes programadores se usaran diferentes cosas
            If FSO.FolderExists(r) Then Combo1.AddItem r
        Loop
        
    TE.Close
End Sub


Private Sub saveFileUsadas(newT As String)
    Dim TE As TextStream
    
    'ver si esta el texto! para no duplicar
    Dim r As String
    Set TE = FSO.OpenTextFile(App.path + "\usadas.txt", ForReading, True)
        If TE.AtEndOfStream Then
            r = ""
        Else
            r = TE.ReadAll
        End If
    TE.Close
    
    If InStr(r, newT) = 0 Then 'no existe, agregarlo
        Set TE = FSO.OpenTextFile(App.path + "\usadas.txt", ForAppending, True)
            TE.WriteLine newT
        TE.Close
    End If
    
End Sub

Private Sub Command2_Click()
    Dim CM As New CommonDialog
    If Combo1 = "" Then
        CM.InitDir = "c:\"
    Else
        Dim Pt As String
        Pt = Combo1
        If Right(Pt, 1) <> "\" Then Pt = Pt + "\"
        CM.InitDir = Pt
    End If
    CM.ShowFolder
    
    If CM.InitDir = "" Or LCase(CM.InitDir) = "c:\" Then
        Exit Sub
    End If
    
    Combo1.Text = CM.InitDir
    
End Sub

Private Sub Command3_Click()
    Dim K As Long
    For K = 0 To List2.ListCount - 1
        If List2.Selected(K) Then
            LIST1.Text = LIST1.Text + "DLL: " + ArchsOrig(K) + vbCrLf
            LIST1.Text = LIST1.Text + "   v" + CStr(FSO.GetFileVersion(ArchsDest(K))) + "==> v" + CStr(FSO.GetFileVersion(ArchsOrig(K))) + vbCrLf
            FSO.CopyFile ArchsOrig(K), ArchsDest(K), True
            RESP = Shell("regsvr32 /s [" + ArchsDest(K) + "]") 'silent
            LIST1.Text = LIST1.Text + "   REG OK?= " + CStr(RESP) + vbCrLf
            LIST1.Refresh
        End If
    Next K
End Sub

Private Sub Form_Load()
    ReadUsadas
End Sub

Private Sub List2_Click()
    If List2.ListIndex < 0 Then Exit Sub
    Dim I As Long
    Dim Extras As String
    I = List2.ListIndex
    Extras = ""
    If FSO.FileExists(ArchsDest(I)) = False Then
        Extras = Extras + "NO DESTINO"
        reemplazar = True
    Else 'solo si existe se puede comparar !
        'choreado del copiseg
        
        If FileDateTime(ArchsOrig(I)) > FileDateTime(ArchsDest(I)) Then
            Extras = Extras + " FECHA"
            reemplazar = True
        End If
        
        If FileLen(ArchsOrig(I)) > FileLen(ArchsDest(I)) Then
            Extras = Extras + " TAMAÑO"
            reemplazar = True
        End If
        
        If FSO.GetFileVersion(ArchsOrig(I)) > FSO.GetFileVersion(ArchsDest(I)) Then
            Extras = Extras + " VERSION"
            reemplazar = True
        End If
    End If
    
    
    Label1.Caption = Extras
End Sub
