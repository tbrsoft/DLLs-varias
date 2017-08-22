VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VolInfo sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   7170
   ClientLeft      =   2325
   ClientTop       =   885
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   5925
   Begin VB.Frame Frame1 
      Caption         =   " GetDiskFreeSpaceEx call - Windows NT and Windows 98 only "
      Height          =   1875
      Left            =   60
      TabIndex        =   17
      Top             =   5220
      Width           =   5835
      Begin VB.Label lblBF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2820
         TabIndex        =   23
         Top             =   1320
         Width           =   2835
      End
      Begin VB.Label lblBOD 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2820
         TabIndex        =   22
         Top             =   900
         Width           =   2835
      End
      Begin VB.Label lblBA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2820
         TabIndex        =   21
         Top             =   480
         Width           =   2835
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Bytes Free"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Bytes Available"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Bytes on Disk"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   2715
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Info"
      Default         =   -1  'True
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label lblFS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   2280
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Free Space"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   15
      Top             =   1980
      Width           =   2835
   End
   Begin VB.Label lblDC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   2280
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Disk Capacity"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   13
      Top             =   1980
      Width           =   2835
   End
   Begin VB.Label lblFN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "File System Name"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   11
      Top             =   1260
      Width           =   2835
   End
   Begin VB.Label lblVolFlags 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   60
      TabIndex        =   10
      Top             =   3180
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Volume Flags"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label lblMaxFilenameLen 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Max Filename Length"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   1260
      Width           =   2835
   End
   Begin VB.Label lblSerial 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Serial Number"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   540
      Width           =   2835
   End
   Begin VB.Label lblVolName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Volume Name"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Drive / Path"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VOLINFO sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample shows how to retrieve the disk volume information.  The main
' use is to see if a disk is compressed, the maximum file name length, and
' the serial number.
'
' I also demo the GetDiskFreeSpace API, which shows total drive space as well
' as available drive space.  Note that I convert the long integers returned
' by GetDiskFreeSpace to double precision.  This is to prevent overflow errors
' when the drive size calculations exceed 2 gig, which is the long integer's max
' value.
'
' 8-17-98 - I added the GetDiskFreeSpaceEx API call, which work on Windows NT
' and Windows 98.

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
    (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
    (ByVal lpRootPathName As String, lpBytesAvailable As Currency, lpTotalBytes As Currency, lpFreeBytes As Currency) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Const FILE_CASE_SENSITIVE_SEARCH = &H1
Private Const FILE_CASE_PRESERVED_NAMES = &H2
Private Const FILE_UNICODE_ON_DISK = &H4
Private Const FILE_PERSISTENT_ACLS = &H8
Private Const FILE_FILE_COMPRESSION = &H10
Private Const FILE_VOLUME_IS_COMPRESSED = &H8000

Private Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Private Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Private Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Private Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
Private Const FS_VOL_IS_COMPRESSED = FILE_VOLUME_IS_COMPRESSED
Private Const FS_FILE_COMPRESSION = FILE_FILE_COMPRESSION

Private Sub Command1_Click()
    Dim lRet As Long, aRoot$, aVN$, lSerial As Long, lMaxFileName As Long
    Dim lFlags As Long, aFN$, a$
    aRoot$ = Text1.Text
    aVN$ = Space$(255)
    aFN$ = Space$(255)
    lRet = GetVolumeInformation(aRoot$, aVN$, Len(aVN$), lSerial, lMaxFileName, lFlags, aFN$, Len(aFN$))
    aVN$ = aVN$ & Chr$(0): lblVolName.Caption = Left$(aVN$, InStr(aVN$, Chr$(0)) - 1): If lblVolName.Caption = "" Then lblVolName.Caption = "{volume has no label}"
    lblSerial.Caption = lSerial
    lblMaxFilenameLen.Caption = lMaxFileName
    lblVolFlags.Caption = ""
    If lFlags And FS_CASE_IS_PRESERVED Then a$ = "FS_CASE_IS_PRESERVED" & vbCrLf
    If lFlags And FS_CASE_SENSITIVE Then a$ = a$ & "FS_CASE_SENSITIVE" & vbCrLf
    If lFlags And FS_UNICODE_STORED_ON_DISK Then a$ = a$ & "FS_UNICODE_STORED_ON_DISK" & vbCrLf
    If lFlags And FS_PERSISTENT_ACLS Then a$ = a$ & "FS_PERSISTENT_ACLS" & vbCrLf
    If lFlags And FS_VOL_IS_COMPRESSED Then a$ = a$ & "FS_VOL_IS_COMPRESSED" & vbCrLf
    If lFlags And FS_FILE_COMPRESSION Then a$ = a$ & "FS_FILE_COMPRESSION" & vbCrLf
    If Len(a$) Then lblVolFlags.Caption = Left$(a$, Len(a$) - 2)
    aFN$ = aFN$ & Chr$(0): lblFN.Caption = Left$(aFN$, InStr(aFN$, Chr$(0)) - 1)
    
    Dim lSecPerClus As Long, lBytePerSec As Long, lNumFreeClus As Long, lTotClus As Long
    lRet = GetDiskFreeSpace(aRoot$, lSecPerClus, lBytePerSec, lNumFreeClus, lTotClus)
    Dim dDC As Double, dFS As Double
    dDC = CDbl(lBytePerSec) * CDbl(lSecPerClus) * CDbl(lTotClus)
    lblDC.Caption = Int(dDC / (1024# * 1024#)) & " meg"
    dFS = CDbl(lBytePerSec) * CDbl(lSecPerClus) * CDbl(lNumFreeClus)
    lblFS.Caption = Int(dFS / (1024# * 1024#)) & " meg"
    
    On Local Error Resume Next
    Err.Clear
    Dim cBA@, cBOD@, cBF@, m#
    lRet = GetDiskFreeSpaceEx(aRoot$, cBA@, cBOD@, cBF@)
    If Err.Number Then
        lblBA.Caption = "????"
        lblBOD.Caption = "????"
        lblBF.Caption = "????"
    Else
        m# = 10000@ / (1024# * 1024#)
        lblBA.Caption = Int(cBA@ * m#) & " meg"
        lblBOD.Caption = Int(cBOD@ * m#) & " meg"
        lblBF.Caption = Int(cBF@ * m#) & " meg"
    End If
End Sub

Private Sub Form_Load()
    Command1_Click
End Sub

