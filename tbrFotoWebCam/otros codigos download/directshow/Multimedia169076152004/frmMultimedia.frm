VERSION 5.00
Begin VB.Form frmMultimedia 
   Caption         =   "Sample: Multimedia Player"
   ClientHeight    =   6435
   ClientLeft      =   1140
   ClientTop       =   1620
   ClientWidth     =   9915
   Icon            =   "frmMultimedia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9915
   Begin VB.PictureBox VisBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   8
      Top             =   0
      Width           =   3495
      Visible         =   0   'False
   End
   Begin VB.Timer StatusUpdate 
      Interval        =   100
      Left            =   360
      Top             =   5520
   End
   Begin VB.Frame Frame 
      Caption         =   "Playstatus:"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   9735
      Begin VB.Timer Visualization 
         Interval        =   1
         Left            =   840
         Top             =   120
      End
      Begin VB.PictureBox BG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   629
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   360
         Width           =   9495
         Begin VB.PictureBox FG 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   0
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1
            TabIndex        =   7
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   0
            Width           =   15
         End
      End
   End
   Begin VB.Frame FramePL 
      Caption         =   "Playlist"
      Height          =   5295
      Left            =   7200
      TabIndex        =   1
      Top             =   60
      Width           =   2655
      Begin VB.CommandButton butAddFolder 
         Caption         =   "Add f&older..."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton butAddFiles 
         Caption         =   "Add f&iles..."
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   2415
      End
      Begin VB.ListBox Playlist 
         Height          =   3825
         IntegralHeight  =   0   'False
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox Video 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   60
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   60
      Width           =   7095
   End
End
Attribute VB_Name = "frmMultimedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sample: Multimedia Player by Vesa Piittinen aka Merri
' http://merri.net
'
' demonstrates the usage of clsActiveMovie, clsBrowseForFolderDialog, clsFileDialog and clsID3v1


Option Explicit


Private Type MyMediaInfo
    Info As New clsID3v1
    File As String
End Type


'API declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Dim ID1() As MyMediaInfo
Dim MyMedia As New clsActiveMovie
Dim Visual As New clsVisualization
Private Sub AddFile(ByVal Filename As String)
    Dim A As Integer, Temp As String
    If ID1(UBound(ID1)).File = vbNullString Then A = 0 Else A = UBound(ID1) + 1
    ReDim Preserve ID1(A)
    ID1(A).File = Filename
    If ID1(A).Info.OpenFile(Filename) Then If ID1(A).Info.Artist <> vbNullString And ID1(A).Info.Title <> vbNullString Then Temp = ID1(A).Info.Artist & " - " & ID1(A).Info.Title
    If Temp = vbNullString Then Temp = Right$(Filename, InStr(StrReverse$(Filename), "\") - 1)
    Playlist.AddItem Temp
    Playlist.ItemData(Playlist.NewIndex) = A
End Sub
Private Function FixPathFile(ByVal Path As String, File As String) As String
    'make sure last character of the path is a \
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    'return path & file
    FixPathFile = Path & File
End Function
Private Sub UpdateProgress()
    FG.Move -1, 0, (BG.ScaleWidth + 1) / BG.Tag * FG.Tag, BG.ScaleHeight
End Sub
Private Sub BG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    MyMedia.SeekTo BG.Tag / BG.ScaleWidth * x
End Sub
Private Sub butAddFiles_Click()
    Dim FileDialog As New clsFileDialog, Filename As String
    Dim A As Integer, Files() As String
    With FileDialog
        .Filter = "All files (*.*)|*.*"
        .Flags = clsFDAllowMultiselect Or clsFDLongNames Or clsFDHideReadOnly Or clsFDFileMustExist Or clsFDExplorer
        .ObjectOwner = Me
        .WindowTitle = "Open file(s)"
        If .FileOpen(Filename) Then
            If InStr(Filename, "|") Then
                Playlist.Visible = False
                Files = Split(Filename, "|")
                For A = LBound(Files) To UBound(Files)
                    AddFile Files(A)
                Next A
                Playlist.Visible = True
            Else
                AddFile Filename
            End If
        End If
    End With
End Sub
Private Sub butAddFolder_Click()
    Dim BrowseDialog As New clsBrowseForFolderDialog, Path As String
    Dim DirList As New Collection, Temp As String, NameSplit() As String
    On Error Resume Next
    With BrowseDialog
        .InitialDirectory = App.Path
        .ObjectOwner = Me
        .WindowMessage = "Open files under path:"
        If Not .FolderDialog(Path) Then Exit Sub
    End With
    Playlist.Visible = False
    DirList.Add FixPathFile(Path, "")
    Do While DirList.Count
        Temp = Dir$(DirList(1), vbDirectory)
        Do Until Temp = ""
            If Temp = "." Or Temp = ".." Then
                'do nothing
            ElseIf (GetAttr(FixPathFile(DirList(1), Temp)) And vbDirectory) = vbDirectory Then
                DirList.Add FixPathFile(DirList(1), Temp) & "\"
            ElseIf InStr(Temp, ".") Then
                NameSplit = Split(Temp, ".")
                Select Case LCase(NameSplit(UBound(NameSplit)))
                    Case "mp3", "mpg", "mpeg", "avi", "ogg"
                        AddFile FixPathFile(DirList(1), Temp)
                End Select
            End If
            Temp = Dir$
        Loop
        DirList.Remove 1
    Loop
    Playlist.Visible = True
End Sub
Private Sub Form_Load()
    ReDim ID1(0)
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Video.Move Video.Left, Video.Top, ScaleWidth - Video.Left * 3 - FramePL.Width, ScaleHeight - Video.Top * 3 - Frame.Height
    VisBG.Move 0, 0, Video.Width, Video.Height
    FramePL.Move ScaleWidth - Video.Left - FramePL.Width, Video.Top, FramePL.Width, Video.Height
    Frame.Move Video.Left, ScaleHeight - Video.Top - Frame.Height, ScaleWidth - Video.Left * 2
    Playlist.Height = FramePL.Height - Playlist.Top - butAddFiles.Height - butAddFolder.Height - Video.Top * 3
    butAddFiles.Top = Playlist.Top + Playlist.Height + Video.Top
    butAddFolder.Top = butAddFiles.Top + butAddFiles.Height
    BG.Move BG.Left, BG.Top, Frame.Width - BG.Left * 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MyMedia.CloseFile
    ReDim ID1(0)
End Sub
Private Sub Playlist_DblClick()
    If Playlist.ListIndex < 0 Then Exit Sub
    If MyMedia.OpenFile(Video.hWnd, ID1(Playlist.ItemData(Playlist.ListIndex)).File) Then Video_Resize: MyMedia.PlayFile: BG.Tag = MyMedia.Length: Playlist.Tag = Playlist.Text
End Sub
Private Sub Playlist_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim A As Integer
    If KeyCode <> vbKeyDelete Then Exit Sub
    If Playlist.SelCount = 0 Then Exit Sub
    For A = Playlist.ListCount - 1 To 0 Step -1
        If Playlist.Selected(A) Then Playlist.RemoveItem A
    Next A
End Sub
Private Sub StatusUpdate_Timer()
    Dim Cur As String, Pos As String
    
    FG.Tag = MyMedia.Position
    UpdateProgress
    
    FG.BackColor = FG.BackColor
    BG.BackColor = BG.BackColor
    
    Cur = MyMedia.FormatHMS(FG.Tag)
    FG.CurrentX = 5
    FG.CurrentY = (FG.ScaleHeight - FG.TextHeight(Cur)) / 2
    BG.CurrentX = 4
    BG.CurrentY = FG.CurrentY
    FG.Print Cur
    BG.Print Cur
    
    Pos = MyMedia.FormatHMS(BG.Tag)
    BG.CurrentX = BG.ScaleWidth - 4 - BG.TextWidth(Pos)
    BG.CurrentY = (BG.ScaleHeight - BG.TextHeight(Pos)) / 2
    FG.CurrentX = BG.CurrentX + 1
    FG.CurrentY = BG.CurrentY
    FG.Print Pos
    BG.Print Pos
End Sub
Private Sub Video_Resize()
    Dim NewWidth As Long, NewHeight As Long
    NewWidth = MyMedia.VideoWidth: NewHeight = MyMedia.VideoHeight
    If NewWidth < 1 Then NewWidth = 1
    If NewHeight < 1 Then NewHeight = 1
    If NewWidth > Video.ScaleWidth Then NewHeight = NewHeight / NewWidth * Video.ScaleWidth: NewWidth = Video.ScaleWidth
    If NewHeight > Video.ScaleHeight Then NewWidth = NewWidth / NewHeight * Video.ScaleHeight: NewHeight = Video.ScaleHeight
    MyMedia.Move (Video.ScaleWidth - NewWidth) / 2, (Video.ScaleHeight - NewHeight) / 2, NewWidth, NewHeight
End Sub
Private Sub VisBG_Resize()
    VisBG.Cls
End Sub
Private Sub Visualization_Timer()
    Static Max(63) As Integer, Avg(63) As Integer, Curs(63) As New Collection, Maxs(63) As New Collection
    Dim A As Integer, B As Byte, Cur(63) As Long, tMax(63) As Long, Bigger As Long, Smaller As Long
    Static C As Byte
    Dim Red As Integer, Green As Integer, Blue As Integer, Soften As Single, Mover As Integer, NewTop As Integer
    If MyMedia.VideoWidth > 1 Then Exit Sub
    Visualization.Enabled = False
    Visual.Update
    For A = 0 To 63
        Cur(A) = CInt(Abs(CLng(Visual.GetData(A * 8)) + CLng(Visual.GetData(A * 8 + 1)) + CLng(Visual.GetData(A * 8 + 2)) + CLng(Visual.GetData(A * 8 + 3)) + CLng(Visual.GetData(A * 8 + 4)) + CLng(Visual.GetData(A * 8 + 5)) + CLng(Visual.GetData(A * 8 + 6)) + CLng(Visual.GetData(A * 8 + 7))) / 8)
        If Cur(A) > Max(A) Then Max(A) = Cur(A)
        If Cur(A) Then Avg(A) = CInt((CLng(Avg(A)) + CLng(Cur(A))) / 2)
        Cur(A) = Cur(A) + Avg(A) / 2
        If Cur(A) < 0 Then Cur(A) = 0
        tMax(A) = Max(A) + Avg(A) / 2
        If tMax(A) = 0 Then tMax(A) = 1
        Curs(A).Add Cur(A)
        Maxs(A).Add tMax(A)
        If Curs(A).Count > 10 Then Curs(A).Remove 1
        If Maxs(A).Count > 10 Then Maxs(A).Remove 1
        Smaller = 0
        Bigger = 0
        For B = 1 To Curs(A).Count
            Soften = Abs(1 + (B - Curs(A).Count / 2) / 50)
            Smaller = Smaller + Val(Curs(A).Item(B)) * Soften
            Bigger = Bigger + Val(Maxs(A).Item(B)) * Soften
        Next B
        Mover = (Cur(A) - Avg(A)) / 1024
        Blue = (384 / Bigger * Smaller / (A / 10 + 0.1) + Mover)
        Green = 208 / Bigger * Smaller + Mover
        Red = 8 / Bigger * Smaller * (32 - Abs(16 - A))
        If Red < 0 Then Red = 0
        If Green < 0 Then Green = 0
        If Blue < 0 Then Blue = 0
        If Red > 255 Then Red = 255
        If Green > 255 Then Green = 255
        If Blue > 255 Then Blue = 255
        Mover = Abs((Sin(C - A) - Cos(C - A)) + (55 - A)) * 4
        If Mover > VisBG.ScaleWidth Or Mover = 0 Then Mover = VisBG.ScaleWidth
        BitBlt VisBG.hdc, 0, NewTop, VisBG.ScaleWidth / Mover * (Mover - 1), VisBG.ScaleHeight / 64 + 1, VisBG.hdc, VisBG.ScaleWidth / Mover, NewTop, vbSrcCopy
        SetPixel VisBG.hdc, VisBG.ScaleWidth - 1, NewTop, RGB(CByte(Red), CByte(Green), CByte(Blue))
        StretchBlt VisBG.hdc, VisBG.ScaleWidth / Mover * (Mover - 1), NewTop, VisBG.ScaleWidth / Mover, VisBG.ScaleHeight / 64 + 1, VisBG.hdc, VisBG.ScaleWidth - 1, NewTop, 1, 1, vbSrcCopy
        NewTop = NewTop + VisBG.ScaleHeight / 64 + 1
    Next A
    If C < 63 Then C = C + 1 Else C = 0
    VisBG.Refresh
    BitBlt Video.hdc, 0, 0, Video.ScaleWidth, Video.ScaleHeight, VisBG.hdc, 0, 0, vbSrcCopy
    Video.CurrentX = 5
    Video.CurrentY = Video.ScaleHeight - Video.TextHeight(Playlist.Tag) - 2
    Video.Print Playlist.Tag
    Video.Refresh
    Visualization.Enabled = True
End Sub
