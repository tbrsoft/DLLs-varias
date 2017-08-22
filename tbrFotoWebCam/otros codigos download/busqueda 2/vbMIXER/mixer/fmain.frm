VERSION 5.00
Begin VB.Form fmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Mixer Line Info Example"
   ClientHeight    =   6690
   ClientLeft      =   1980
   ClientTop       =   1305
   ClientWidth     =   8115
   Icon            =   "fmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8115
   Begin VB.ListBox lstSrcChannelInfo 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "fmain.frx":1272
      Left            =   4968
      List            =   "fmain.frx":1274
      TabIndex        =   21
      Top             =   6072
      Width           =   2952
   End
   Begin VB.ListBox lstDstChannelInfo 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "fmain.frx":1276
      Left            =   4968
      List            =   "fmain.frx":1278
      TabIndex        =   20
      Top             =   2748
      Width           =   2952
   End
   Begin VB.ComboBox cbSrcChannelList 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5916
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5724
      Width           =   1836
   End
   Begin VB.ComboBox cbDstChannelList 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5916
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Width           =   1836
   End
   Begin VB.Frame framDestInfo 
      Caption         =   "Destination Line Info"
      Height          =   2580
      Left            =   45
      TabIndex        =   14
      Top             =   360
      Width           =   4632
      Begin VB.ComboBox cbDstLineList 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "fmain.frx":127A
         Left            =   720
         List            =   "fmain.frx":127C
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   274
         Width           =   1668
      End
      Begin VB.ListBox lstDstLineInfo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         ItemData        =   "fmain.frx":127E
         Left            =   132
         List            =   "fmain.frx":1280
         TabIndex        =   15
         Top             =   564
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "Lines:"
         Height          =   240
         Left            =   132
         TabIndex        =   17
         Top             =   280
         Width           =   600
      End
   End
   Begin VB.Frame framSrcInfo 
      Caption         =   "Source Line Info"
      Height          =   2616
      Left            =   45
      TabIndex        =   10
      Top             =   3696
      Width           =   4632
      Begin VB.ListBox lstSrcLineInfo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         ItemData        =   "fmain.frx":1282
         Left            =   132
         List            =   "fmain.frx":1284
         TabIndex        =   12
         Top             =   588
         Width           =   4380
      End
      Begin VB.ComboBox cbSrcLineList 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "fmain.frx":1286
         Left            =   720
         List            =   "fmain.frx":1288
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   309
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Lines:"
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Source Control Info"
      Height          =   1932
      Left            =   4860
      TabIndex        =   6
      Top             =   3696
      Width           =   3120
      Begin VB.ListBox lstSrcControlInfo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         ItemData        =   "fmain.frx":128A
         Left            =   168
         List            =   "fmain.frx":128C
         TabIndex        =   8
         Top             =   552
         Width           =   2856
      End
      Begin VB.ComboBox cbSrcControlList 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "fmain.frx":128E
         Left            =   1032
         List            =   "fmain.frx":1290
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   276
         Width           =   1716
      End
      Begin VB.Label Label4 
         Caption         =   "Controls:"
         Height          =   240
         Left            =   192
         TabIndex        =   9
         Top             =   240
         Width           =   620
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destination Control Info"
      Height          =   1956
      Left            =   4848
      TabIndex        =   2
      Top             =   348
      Width           =   3120
      Begin VB.ListBox lstDstControlInfo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         ItemData        =   "fmain.frx":1292
         Left            =   144
         List            =   "fmain.frx":1294
         TabIndex        =   4
         Top             =   576
         Width           =   2868
      End
      Begin VB.ComboBox cbDstControlList 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "fmain.frx":1296
         Left            =   1032
         List            =   "fmain.frx":1298
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   264
         Width           =   1836
      End
      Begin VB.Label Label3 
         Caption         =   "Controls:"
         Height          =   240
         Left            =   192
         TabIndex        =   5
         Top             =   240
         Width           =   620
      End
   End
   Begin VB.ComboBox cbMixerList 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      ItemData        =   "fmain.frx":129A
      Left            =   1755
      List            =   "fmain.frx":129C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   2568
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   12
      X2              =   8088
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   24
      X2              =   8100
      Y1              =   3372
      Y2              =   3372
   End
   Begin VB.Label lblMxList 
      Caption         =   "Audio Mixer Devices:"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mx As CMixer








Private Sub cbDstChannelList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.DstChannelID = cbDstChannelList.ListIndex
    
        PrintDstChannelInfo

End Sub
Private Sub cbSrcChannelList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.SrcChannelID = cbSrcChannelList.ListIndex
    
        PrintSrcChannelInfo

End Sub

Private Sub Form_Load()

Dim i As Long


Me.Refresh

'Create a Mixer class
Set mx = New CMixer
Me.MousePointer = vbHourglass
On Error Resume Next
If 0 = mx.Create(Me.hWnd) Then 'create() returns number of mixers on system
    'if there is no sound card just exit
    MsgBox "No Audio Mixers Detected!", vbCritical, App.Title
    Exit Sub
End If
'trap other errors here
If Err Then MsgBox Err.Description & " src: " & Err.Source
For i = 0 To mx.numMixerDevs - 1
    mx.MixerID = i
    If Err Then
        MsgBox Err.Description & " src: " & Err.Source
        cbMixerList.ListIndex = 0
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Call cbMixerList.AddItem(mx.MixerName)
Next i
cbMixerList.ListIndex = 0
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mx = Nothing

End Sub

Private Sub cbMixerList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.MixerID = cbMixerList.ListIndex
    If Err Then
        MsgBox Err.Description & " src: " & Err.Source
        cbDstLineList.ListIndex = 0
        Exit Sub
    End If
    cbDstLineList.Clear
    
    For i = 0 To mx.numDestinations - 1
        mx.DstLineID = i
        If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbDstLineList.ListIndex = 0
            Exit Sub
        End If
        cbDstLineList.AddItem (mx.DstLineName)
    Next i
        
    For i = 0 To mx.numDstLineControls - 1
        mx.DstControlID = i
                If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbDstControlList.ListIndex = 0
            Exit Sub
        End If
        cbDstControlList.AddItem (mx.DstControlName)
    Next i

    cbDstLineList.ListIndex = 0
    
End Sub
Private Sub cbDstLineList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.DstLineID = cbDstLineList.ListIndex
    
        PrintDstLineInfo
   cbDstControlList.Clear
       For i = 0 To mx.numDstLineControls - 1
        mx.DstControlID = i
                If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbDstControlList.ListIndex = 0
            Exit Sub
        End If
        cbDstControlList.AddItem (mx.DstControlName(False))
    Next i
cbDstControlList.ListIndex = 0 'add jump to a get control details function
   
    If Err Then MsgBox Err.Description & " src: " & Err.Source
    'initialize Source lines list for this Destination Line
    cbSrcLineList.Clear
    
    
    If mx.numDstLineConnections > 0 Then
        For i = 0 To mx.numDstLineConnections - 1
            mx.SrcLineID = i
            If Err Then
                'MsgBox Err.Description & " src: " & Err.Source
                cbSrcLineList.AddItem "<error>"
                Err = 0
            Else
                cbSrcLineList.AddItem mx.SrcLineName
            End If
        Next i
       
    Else
        cbSrcLineList.AddItem "<none>"
    End If
    cbSrcLineList.ListIndex = 0
      
    
End Sub
Private Sub cbDstControlList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.DstControlID = cbDstControlList.ListIndex
    
        PrintDstControlInfo
        
      cbDstChannelList.Clear
       For i = 0 To mx.numDstControlChannels - 1
        mx.DstChannelID = i
                If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbDstChannelList.ListIndex = 0
            Exit Sub
        End If
        cbDstChannelList.AddItem "channel: " & mx.DstChannelID
    Next i
cbDstChannelList.ListIndex = 0 'add jump to a get control details function


   
    If Err Then MsgBox Err.Description & " src: " & Err.Source
    'initialize Source lines list for this Destination Line
    
    
    
        
    
End Sub
Private Sub cbsrcControlList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.SrcControlID = cbSrcControlList.ListIndex
    
        PrintSrcControlInfo
         cbSrcChannelList.Clear
       For i = 0 To mx.numSrcControlChannels - 1
        mx.SrcChannelID = i
                If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbSrcChannelList.ListIndex = 0
            Exit Sub
        End If
        cbSrcChannelList.AddItem "channel: " & mx.SrcChannelID
    Next i
cbSrcChannelList.ListIndex = 0 'add jump to a get control details function

   'add jump to a get control details function
   
    If Err Then MsgBox Err.Description & " src: " & Err.Source
    'initialize Source lines list for this Destination Line
    'cbSrccontrolList.Clear
    
    
        
    
End Sub

Private Sub cbSrcLineList_Click()
    Dim i As Long
    
    On Error Resume Next
    mx.SrcLineID = cbSrcLineList.ListIndex
    If Err Then
        MsgBox Err.Description & " src: " & Err.Source
        Exit Sub
    End If
   
    
    PrintSrcLineInfo
       cbSrcControlList.Clear
       For i = 0 To mx.numSrcLineControls - 1
        mx.SrcControlID = i
                If Err Then
            MsgBox Err.Description & " src: " & Err.Source
            cbSrcControlList.ListIndex = 0
            Exit Sub
        End If
        cbSrcControlList.AddItem (mx.SrcControlName(False))
    Next i
cbSrcControlList.ListIndex = 0 'add jump to a get control details function

  'add jump to a get control details function
End Sub



Private Sub PrintDstLineInfo()
    With lstDstLineInfo
        .Clear
        .AddItem "========================="
        .AddItem mx.DstLineName
        .AddItem "========================="
        .AddItem "Line ID: " & mx.DstLineID
        .AddItem "Line is disconnected: " & mx.DstLineDisconnected
        .AddItem "Line is active: " & mx.DstLineActive
        .AddItem "Number of Channels: " & mx.numDstLineChannels
        .AddItem "Number of Controls: " & mx.numDstLineControls
        .AddItem "Number of Connections: " & mx.numDstLineConnections
        .AddItem "Audio Line Type: " & GetLineTypeString(mx.DstLineType)
        .AddItem "Line Target Type: " & GetTargetTypeString(mx.DstLineTarget)
    End With
End Sub
Private Sub PrintDstControlInfo()
    With lstDstControlInfo
        .Clear
        .AddItem "========================="
        .AddItem mx.DstControlName
        .AddItem "========================="
        .AddItem "Control ID: " & mx.DstControlID
        .AddItem "Control Type: " & GetControlTypeString(mx.DstControlType)
        .AddItem "Maximum Value: " & mx.DstChannelMaxValue
        .AddItem "Minimum Value: " & mx.DstChannelMinValue
        
    End With
End Sub
Private Sub PrintDstChannelInfo()
    With lstDstChannelInfo
    .Clear
    .AddItem "========================="
    If mx.DstChannelValue <> vbNull Then
    .AddItem "Channel: " & mx.DstChannelID & " Value: " & mx.DstChannelValue
    Else
    .AddItem "not implimented"
    End If
    .AddItem "========================="
    
End With
End Sub
Private Sub PrintSrcControlInfo()
    With lstSrcControlInfo
        .Clear
        .AddItem "========================="
        .AddItem mx.SrcControlName
        .AddItem "========================="
        .AddItem "Control ID: " & mx.SrcControlID
        .AddItem "Control Type: " & GetControlTypeString(mx.SrcControlType)
        .AddItem "Maximum Value: " & mx.SrcChannelMaxValue
        .AddItem "Minimum Value: " & mx.SrcChannelMinValue
    End With
End Sub
Private Sub PrintSrcChannelInfo()
    With lstSrcChannelInfo
    .Clear
    .AddItem "========================="
    If mx.SrcChannelValue <> vbNull Then
    .AddItem "Channel: " & mx.SrcChannelID & " Value: " & mx.SrcChannelValue
    Else
    .AddItem "not implimented"
    End If
    .AddItem "========================="
   
End With
End Sub

Private Sub PrintSrcLineInfo()
    With lstSrcLineInfo
        .Clear
        .AddItem "========================="
        .AddItem mx.SrcLineName
        .AddItem "========================="
        .AddItem "Line ID: " & mx.SrcLineID
        .AddItem "Line is disconnected: " & mx.SrcLineDisconnected
        .AddItem "Line is active: " & mx.SrcLineActive
        .AddItem "Number of Channels: " & mx.numSrcLineChannels
        .AddItem "Number of Controls: " & mx.numSrcLineControls
        .AddItem "Number of Connections: " & mx.numSrcLineConnections
        .AddItem "Audio Line Type: " & GetLineTypeString(mx.SrcLineType)
        .AddItem "Line Target Type: " & GetTargetTypeString(mx.SrcLineTarget)
    End With
End Sub

