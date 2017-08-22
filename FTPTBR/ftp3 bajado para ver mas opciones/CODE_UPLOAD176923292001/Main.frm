VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Click to upload to server"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox Dirname 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "windows\desktop\"
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "Main.frx":0000
      Left            =   120
      List            =   "Main.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret As Boolean
Dim l As Long
Dim Parts() As String

Private Sub cmdAdd_Click()
Dim c As New cCommonDialog
Dim finprod As String
Dim athh As String
Dim retval As Boolean
Dim fir As String
    With c
        .DialogTitle = "Choose File..."
        .CancelError = False
        .hwnd = Me.hwnd
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
        .InitDir = CurDir
        .Filename = "Autoup"
        .Filter = "All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
        athh = .Filename
        fir = .FileTitle
        
        finprod = Mid(.Filename, 1, Len(.Filename) - Len(.FileTitle))
    End With
    
    If athh = "" Then Exit Sub
    
    List1.AddItem athh
    List1.ItemData(List1.NewIndex) = Len(fir)


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
  List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub cmdSave_Click()
Dim j As Integer
  With inigo
  
    If Check1.Value = 1 Then
      .Section = "AutoUP"
    ElseIf Check1.Value = 0 Then
      .Section = "AutoDN"
    End If
    
      .Key = "count"
      .Value = List1.ListCount
      
      For j = 0 To List1.ListCount - 1
        .Key = "file" & (j + 1)
        .Value = List1.List(j)
        
      If Check1.Value = 1 Then
         .Key = "path" & (j + 1)
         .Value = Mid(List1.List(j), 1, Len(List1.List(j)) - List1.ItemData(j))
      End If
      
      Next
      
      .Key = "ChDirName"
      .Value = Dirname
      
  End With

End Sub
Private Sub Form_Load()
'
' If you were to have it autoload the ini file also
' you have to remember to put the length of filename
' into the list.itemdata
'
 Set inigo = New cIniFile
  Retrned = inigo.LastReturnCode
  
  ret = Validate_File(App.Path & "\autoup.ini")
  'ret = Validate_File("c:\windows\desktop\autoup.ini")
  
  If ret = False Then
  CreateIt (App.Path & "\autoup.ini")
  End If
  
  inigo.Path = App.Path & "\autoup.ini"
  'inigo.Path = "c:\windows\desktop\autoup.ini"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set inigo = Nothing
End Sub
