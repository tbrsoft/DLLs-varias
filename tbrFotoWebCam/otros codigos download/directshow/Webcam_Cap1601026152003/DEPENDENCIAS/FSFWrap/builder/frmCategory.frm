VERSION 5.00
Begin VB.Form frmCategory 
   Caption         =   "Filters By Category"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox listFilters 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox comboCategory 
      Height          =   315
      ItemData        =   "frmCategory.frx":0000
      Left            =   120
      List            =   "frmCategory.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public graph As IMediaControl
Public filter As IFilterInfo
Private catlist As IVBCollection

Public Sub RefreshList()

    Dim catname As String
    catname = comboCategory.Text
    
    ' map category names to GUIDs. further guids can be found in DirectShow sdk uuids.h
    Dim strcat As String
    If catname = "Video Capture" Then
        strcat = "{860BB310-5D01-11d0-BD3B-00A0C911CE86}"
    ElseIf catname = "Video Compressor" Then
        strcat = "{33D9A760-90C8-11d0-BD43-00A0C911CE86}"
    ElseIf catname = "Audio Capture" Then
        strcat = "{33D9A762-90C8-11d0-BD43-00A0C911CE86}"
    ElseIf catname = "Audio Compressor" Then
        strcat = "{33D9A761-90C8-11d0-BD43-00A0C911CE86}"
    ElseIf catname = "Audio Render" Then
        strcat = "{E0F158E1-CB04-11d0-BD4E-00A0C911CE86}"
    ElseIf catname = "Midi Render" Then
        strcat = "{4EFE2452-168A-11d1-BC76-00C04FB9453B}"
    ElseIf catname = "DirectShow Filters" Then
        strcat = "{083863F1-70DE-11d0-BD40-00A0C911CE86}"
    End If
        
    Dim fce As FilterCatEnumerator
    Set fce = New FilterCatEnumerator
    Set catlist = fce.Category(strcat)
        
    listFilters.Clear
    
    Dim f As IFilterClass
    For Each f In catlist
        listFilters.AddItem f.Name
    Next f
End Sub


Private Sub cmdCancel_Click()
    Set filter = Nothing
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    If catlist Is Nothing Then Exit Sub
    If listFilters.ListIndex < 0 Then Exit Sub
    
    Dim f As IFilterClass
    Set f = catlist.Item(listFilters.ListIndex)
    Set filter = f.Create(graph)
    Unload Me
    
End Sub



Private Sub cmdShow_Click()
    RefreshList
End Sub

Private Sub Form_Load()
    ' fill list box
    comboCategory.AddItem "Video Capture"
    comboCategory.AddItem "Video Compressor"
    comboCategory.AddItem "Audio Capture"
    comboCategory.AddItem "Audio Compressor"
    comboCategory.AddItem "Audio Render"
    comboCategory.AddItem "Midi Render"
    comboCategory.AddItem "DirectShow Filters"
    comboCategory.ListIndex = 0
    RefreshList
End Sub

Private Sub listFilters_DblClick()
    cmdInsert_Click
End Sub
