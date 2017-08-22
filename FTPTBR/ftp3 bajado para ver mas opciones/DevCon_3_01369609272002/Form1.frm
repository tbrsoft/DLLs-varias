VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "DevCon FTP "
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   9120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6246
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":669A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":70FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":72C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":74A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":78F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":797E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":816E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":81F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8282
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":839A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":849E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   4740
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6135
      ScaleWidth      =   60
      TabIndex        =   14
      Top             =   1680
      Width           =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   9615
      TabIndex        =   13
      Tag             =   "mov"
      Top             =   3450
      Width           =   9615
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   975
      Left            =   40
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "mov"
      Top             =   600
      Width           =   9495
   End
   Begin MSComctlLib.TreeView TView1 
      Height          =   1580
      Left            =   4800
      TabIndex        =   9
      Tag             =   "mov"
      Top             =   1880
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2805
      _Version        =   393217
      Indentation     =   106
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":85BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":875A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":904A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9206
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":93A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7800
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9546
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":967E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":975E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9886
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":997E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1111
      ButtonWidth     =   1561
      ButtonHeight    =   1058
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect"
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect To Server"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disconnect"
            Key             =   "Disconnect"
            Object.ToolTipText     =   "Close FTP Session"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spr1"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Reload"
            Object.ToolTipText     =   "Refresh Local Directory"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Upload"
            Key             =   "Upload"
            Object.ToolTipText     =   "Upload File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "+Folder"
            Key             =   "NewFolder"
            Object.ToolTipText     =   "Create New Local Folder"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-Folder"
            Key             =   "DelFolder"
            Object.ToolTipText     =   "Delete Local Folder"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete File"
            Key             =   "DelFile"
            Object.ToolTipText     =   "Delete Local File"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Key             =   "Rename"
            Object.ToolTipText     =   "Rename Local File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Reload2"
            Object.ToolTipText     =   "Refresh Sever Directory"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Download"
            Key             =   "DownLoad"
            Object.ToolTipText     =   "Download File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "+Folder"
            Key             =   "NewFolder2"
            Object.ToolTipText     =   "Create New Folder On Server"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-Folder"
            Key             =   "DelFolder2"
            Object.ToolTipText     =   "Delete Folder On Server"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete File"
            Key             =   "DelFile2"
            Object.ToolTipText     =   "Delete File On Server"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Key             =   "Rename2"
            Object.ToolTipText     =   "Rename File On Server"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7800
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1235
            MinWidth        =   1235
            Picture         =   "Form1.frx":A0AE
            TextSave        =   "2:09 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14023
            Picture         =   "Form1.frx":A15E
            Object.ToolTipText     =   "Current Working Directory"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Offline"
            TextSave        =   "Offline"
            Object.ToolTipText     =   "You are not connected."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   40
      TabIndex        =   3
      Tag             =   "mov"
      Top             =   1880
      Width           =   4695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4320
      Left            =   45
      TabIndex        =   2
      Tag             =   "mov"
      Top             =   3495
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7620
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList4"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   40
      TabIndex        =   1
      Tag             =   "mov"
      Top             =   2240
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   6960
      TabIndex        =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4320
      Left            =   4800
      TabIndex        =   4
      Tag             =   "mov"
      Top             =   3495
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7620
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList4"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image picMenu8 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A20E
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu7 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A282
      Top             =   4680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu6 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A2F3
      Top             =   4440
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu5 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A358
      Top             =   4080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu4 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A3C5
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu3 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A453
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu2 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "Form1.frx":A4A4
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   3
      Left            =   9480
      Picture         =   "Form1.frx":A53E
      Top             =   3000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   2
      Left            =   9480
      Picture         =   "Form1.frx":A5BA
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   1
      Left            =   9480
      Picture         =   "Form1.frx":A649
      Top             =   2520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   0
      Left            =   9480
      Picture         =   "Form1.frx":A6C0
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5840
      TabIndex        =   12
      Tag             =   "mov"
      Top             =   1640
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Tag             =   "mov"
      Top             =   1635
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000C&
      Caption         =   " Server:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   4800
      TabIndex        =   8
      Tag             =   "mov"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   " Local:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   40
      TabIndex        =   7
      Tag             =   "mov"
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Menu zSoubor 
      Caption         =   "&File"
      Begin VB.Menu zOpen 
         Caption         =   "&Open Log"
         Shortcut        =   ^O
      End
      Begin VB.Menu zSave 
         Caption         =   "&Save Log"
         Shortcut        =   ^S
      End
      Begin VB.Menu zConnect 
         Caption         =   "&Connect To Server"
         Shortcut        =   ^E
      End
      Begin VB.Menu zDisconnect 
         Caption         =   "&Disconnect From Server"
         Shortcut        =   ^Q
      End
      Begin VB.Menu zSep9 
         Caption         =   "-"
      End
      Begin VB.Menu zOdpojit 
         Caption         =   "&Close Connection"
      End
      Begin VB.Menu zSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTray 
         Caption         =   "Send To Tray"
         Shortcut        =   ^M
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu zEnd 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu zLocal 
      Caption         =   "&Local"
      Begin VB.Menu zLokRef 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu zOpenFile 
         Caption         =   "&Open File"
      End
      Begin VB.Menu zSep6 
         Caption         =   "-"
      End
      Begin VB.Menu zLokUp 
         Caption         =   "&Upload"
         Shortcut        =   ^U
      End
      Begin VB.Menu zSep5 
         Caption         =   "-"
      End
      Begin VB.Menu zLokNF 
         Caption         =   "&New Folder"
      End
      Begin VB.Menu zLokDF 
         Caption         =   "&Remove Folder"
      End
      Begin VB.Menu zLokDS 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu zLokRS 
         Caption         =   "Rename &File"
      End
      Begin VB.Menu zSep3 
         Caption         =   "-"
      End
      Begin VB.Menu zPat 
         Caption         =   "Pa&ttern"
         Begin VB.Menu zAll 
            Caption         =   "&All Files (*.*)"
            Checked         =   -1  'True
         End
         Begin VB.Menu zSep11 
            Caption         =   "-"
         End
         Begin VB.Menu zDefine 
            Caption         =   "&Define types (*.?)"
         End
      End
      Begin VB.Menu zFind 
         Caption         =   "Fin&d File"
         Shortcut        =   ^F
      End
      Begin VB.Menu zProperties 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu zRemote 
      Caption         =   "&Remote"
      Begin VB.Menu zRemRef 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu zSep7 
         Caption         =   "-"
      End
      Begin VB.Menu zRemDown 
         Caption         =   "&Download"
         Shortcut        =   ^D
      End
      Begin VB.Menu zSep8 
         Caption         =   "-"
      End
      Begin VB.Menu zRemNF 
         Caption         =   "&New Folder"
      End
      Begin VB.Menu zRemDF 
         Caption         =   "Rem&ove Folder"
      End
      Begin VB.Menu zRemDS 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu zRemRS 
         Caption         =   "Ren&ame File"
      End
      Begin VB.Menu zSep12 
         Caption         =   "-"
      End
      Begin VB.Menu zPat2 
         Caption         =   "Pa&ttern"
         Begin VB.Menu zAll2 
            Caption         =   "&All Files (*.*)"
            Checked         =   -1  'True
         End
         Begin VB.Menu zSep13 
            Caption         =   "-"
         End
         Begin VB.Menu zDefine2 
            Caption         =   "&Define Types (*.?)"
         End
      End
   End
   Begin VB.Menu zNast 
      Caption         =   "&Tools"
      Begin VB.Menu zTento 
         Caption         =   "&Local"
         Begin VB.Menu ztBigIc 
            Caption         =   "&Big Icons"
         End
         Begin VB.Menu ztSmallIc 
            Caption         =   "&Small Icons"
         End
         Begin VB.Menu ztSeznam 
            Caption         =   "&List"
         End
         Begin VB.Menu ztReport 
            Caption         =   "&Report"
         End
      End
      Begin VB.Menu zServer 
         Caption         =   "&Server"
         Begin VB.Menu zsBigIc 
            Caption         =   "&Big Icons"
         End
         Begin VB.Menu zsSmallIc 
            Caption         =   "&Small Icons"
         End
         Begin VB.Menu zsSeznam 
            Caption         =   "&List"
         End
         Begin VB.Menu zsReport 
            Caption         =   "&Report"
         End
      End
      Begin VB.Menu zSep4 
         Caption         =   "-"
      End
      Begin VB.Menu zTransf 
         Caption         =   "&Transfer"
         Begin VB.Menu zBinary 
            Caption         =   "&Binary"
            Checked         =   -1  'True
         End
         Begin VB.Menu zAscii 
            Caption         =   "&ASCII"
         End
      End
      Begin VB.Menu zPassive 
         Caption         =   "&Pasive mode"
      End
      Begin VB.Menu zSep10 
         Caption         =   "-"
      End
      Begin VB.Menu zTools 
         Caption         =   "S&how Toolbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "&Advanced"
      Begin VB.Menu mnuTelnet 
         Caption         =   "Telnet"
      End
      Begin VB.Menu mnuDNS 
         Caption         =   "DNS Query"
      End
   End
   Begin VB.Menu zHilfe 
      Caption         =   "&Help"
      Begin VB.Menu zHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu zSep2 
         Caption         =   "-"
      End
      Begin VB.Menu zAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code base from YZYFTP.
'Rewritten By Resiware, Inc. (support@resiware.com)
'Code is under GNU liscense. www.gnu.org
'This code, excluding the GUI, is open source for all!
'Please vote for this project!
'Any questions just email me! Or post to PSC forum.



Option Explicit
Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SHFindFiles Lib "shell32.dll" Alias "#90" (ByVal pidlRoot As Long, ByVal pidlSavedSearches As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private lsDrag() As ListItem
Private fPat As String
Dim iPos As Integer
Dim strExt As String
Dim tvNode As Node
Dim lsItem As ListItem
Dim RetVal As Long

Public Sub List()
Dim hFile As Long, udtWFD As WIN32_FIND_DATA
Dim strFile As String
Dim Img As Integer, r As Integer
Dim L&
Dim sTime As SYSTEMTIME, lTime As FILETIME

If session = 0 Or server = 0 Then
    MsgBox "You Are Not Connected. Please Click The Connect Button", vbInformation, App.Title
    Exit Sub
End If
StatusBar1.Panels(2).Text = Time & "  > Sending request..., wait."
    frmmain.MousePointer = 11
    ListView2.ListItems.Clear
    frmmain.txtInfo.SelText = Time & " > Transfering data..." & vbCrLf
    txtInfo.SelText = Time & " > Opening folder: " & Chr(34) & adr & Chr(34) & vbCrLf
    hFile = FtpFindFirstFile(server, adr, udtWFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
        If hFile Then
          Do
            strFile = Left(udtWFD.cFileName, InStr(1, udtWFD.cFileName, Chr(0)) - 1)
                If Len(strFile) > 0 Then
                    If udtWFD.dwFileAttributes And vbDirectory Then
                              Set tvNode = TView1.Nodes.Add(Klic, tvwChild, Klic & strFile & "/", strFile, 12, 13)
                              TView1.Nodes(1).Expanded = True
                    Else
                        Img = ImgNumber(strFile)
                        Set lsItem = ListView2.ListItems.Add(, , strFile, Img, Img)
                        lsItem.SubItems(1) = Format((udtWFD.nFileSizeLow / 1024), "### ### ###.##") & "Kb"
                          lTime = udtWFD.ftLastWriteTime
                          L = FileTimeToSystemTime(lTime, sTime)
                        lsItem.SubItems(2) = CalcFTime(sTime)
                        lsItem.SubItems(3) = udtWFD.nFileSizeLow
                    End If
                End If
            Loop While InternetFindNextFile(hFile, udtWFD)
        End If
    InternetCloseHandle hFile
    txtInfo.SelText = Time & " > Data transfer completed succesfully." & vbCrLf
ListView2.SelectedItem = Nothing
frmmain.MousePointer = 0
StatusBar1.Panels(2).Text = "Server: " & ListView2.ListItems.Count & " Files in folder: " & adr
End Sub

Private Sub Dir1_Click()
StatusBar1.Panels(2).Text = "Local: Double click folder for retreiving files"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
LoadLocal
End Sub

Private Sub Dir1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    zLokDF_Click
End If
End Sub

Private Sub Drive1_GotFocus()
StatusBar1.Panels(2).Text = "Local: Choose local drive."
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Dir1.Refresh
End Sub

Private Sub Form_Load()
Dim hMenu As Long, hSubMenu As Long
Dim RetVal As Long
Dim i As Long
hMenu = GetMenu(Me.hwnd)
hSubMenu = GetSubMenu(hMenu, 0)
For i = 0 To 3
    RetVal = SetMenuItemBitmaps(hSubMenu, i, MF_BYPOSITION, picMenu(i).Picture, picMenu(i).Picture)
Next i
RetVal = SetMenuItemBitmaps(hSubMenu, 5, MF_BYPOSITION, picMenu2.Picture, picMenu2.Picture)
hSubMenu = GetSubMenu(hMenu, 1)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu3.Picture, picMenu3.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, picMenu(0).Picture, picMenu(0).Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 3, MF_BYPOSITION, picMenu4.Picture, picMenu4.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 11, MF_BYPOSITION, picMenu6.Picture, picMenu6.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 12, MF_BYPOSITION, picMenu5.Picture, picMenu5.Picture)
hSubMenu = GetSubMenu(hMenu, 2)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu3.Picture, picMenu3.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 2, MF_BYPOSITION, picMenu4.Picture, picMenu4.Picture)
hSubMenu = GetSubMenu(hMenu, 3)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu7.Picture, picMenu7.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, picMenu(2).Picture, picMenu(2).Picture)
hSubMenu = GetSubMenu(hMenu, 4)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu8.Picture, picMenu8.Picture)
fPat = "*.*"
Dir1.Path = App.Path
LoadLocal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InternetCloseHandle server
    InternetCloseHandle session
End
End Sub

Private Sub Form_Resize()
If frmmain.WindowState = 1 Then Exit Sub
Picture1.Width = frmmain.Width - 40
Picture2.Left = frmmain.Width / 2 - 80
Picture2.Height = frmmain.ScaleHeight - Picture2.Top - 240
Dir1.Height = Picture1.Top - Dir1.Top
TView1.Height = Picture1.Top - TView1.Top
txtInfo.Width = frmmain.ScaleWidth - 80
Drive1.Width = frmmain.Width / 2 - 140
Dir1.Width = frmmain.Width / 2 - 140
TView1.Left = Drive1.Width + 120
TView1.Width = frmmain.Width / 2 - 140
ListView1.Width = frmmain.Width / 2 - 140
ListView2.Left = Drive1.Width + 120
ListView2.Width = frmmain.Width / 2 - 140
ListView1.Height = frmmain.ScaleHeight - (Picture1.Top + 260)
ListView2.Height = frmmain.ScaleHeight - (Picture1.Top + 260)
Label3.Left = TView1.Left + 1040
Label4.Left = TView1.Left
Label1.Width = ListView1.Width - 940
Label2.Width = ListView1.Width
Label3.Width = ListView2.Width - 1040
Label4.Width = ListView2.Width
ListView1.Refresh
ListView2.Refresh
End Sub
Private Sub LoadLocal()
Dim X As Integer, Img As Integer
Dim Y As Long
Drive1.Refresh
Dir1.Refresh
File1.Refresh
ListView1.ListItems.Clear
If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
       strPath = Dir1.Path
 Else: strPath = Dir1.Path & "\"
End If
Label1.Caption = strPath
     If Len(Label1.Caption) > 30 Then
        Label1.Caption = "..." & Trim(Right(Label1.Caption, 30))
     End If
For X = 0 To File1.ListCount - 1
 Img = ImgNumber(File1.List(X))
 With ListView1.ListItems.Add(, , File1.List(X), Img, Img)
   .SubItems(1) = Format((FileLen(strPath & File1.List(X)) / 1000), "### ### ###.##") & " Kb"
   .SubItems(2) = FileDateTime(strPath & File1.List(X))
   Y = Str(FileLen(strPath & File1.List(X)))
   .SubItems(3) = Str(FileLen(strPath & File1.List(X)))
End With
Next
ListView1.SelectedItem = Nothing
StatusBar1.Panels(2).Text = "Local: " & File1.ListCount & " Files in folder: " & strPath
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim strEx2 As String, strEx1 As String
Dim Msg As VbMsgBoxResult
On Error GoTo Err
strEx1 = Mid$(ListView1.SelectedItem.Text, InStrRev(ListView1.SelectedItem.Text, ".") + 1)
strEx2 = Mid$(NewString, InStrRev(NewString, ".") + 1)
If strEx1 <> strEx2 Then
    Msg = MsgBox("Are you sure to exchange the file extension from: " & Chr(34) & strEx1 & Chr(34) & " to: " & Chr(34) & strEx2 & Chr(34), vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        Cancel = 0
        Name strPath & ListView1.SelectedItem.Text As strPath & NewString
        zLokRef_Click
    Else: Cancel = 1
    End If
Else
    Cancel = 0
    Name strPath & ListView1.SelectedItem.Text As strPath & NewString
    zLokRef_Click
End If
Err: If Err.Number = 58 Then
MsgBox "More than one file with the same name in one folder? no way!", vbExclamation, App.Title
Cancel = 1
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
StatusBar1.Panels(2).Text = "Local: Renaming file: " & Chr(34) & ListView1.SelectedItem.Text & Chr(34)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
If ListView1.SortOrder = 0 Then
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = 1
 Else   ' Set Sorted to True to sort the list.
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = 0
End If
 ListView1.Sorted = True
End Sub

Private Sub ListView1_Click()
Dim i, X As Integer
Dim Y, z As Long
X = 0
z = 0
If ListView1.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
           Y = ListView1.ListItems(i).SubItems(3)
           z = z + Y
           X = X + 1
        End If
    Next i
StatusBar1.Panels(2).Text = "Local: " & X & " Files selected, " & z / 1000 & " Kb"
zProperties.Enabled = True
End Sub

Private Sub ListView1_DblClick()
zOpenFile_Click
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    zLokDS_Click
ElseIf KeyCode = vbKeyReturn Then
    zOpenFile_Click
'ElseIf (Shift And vbShiftMask) > 0 Then
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
'        ListView1_Click
'    End If
'ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
'        ListView1_Click
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu zLocal
End If
End Sub


Private Sub ListView2_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim strEx2 As String, strEx1 As String, Old As String, Nw As String
Dim Msg As VbMsgBoxResult
On Error GoTo Err
strEx1 = Mid$(ListView2.SelectedItem.Text, InStrRev(ListView2.SelectedItem.Text, ".") + 1)
strEx2 = Mid$(NewString, InStrRev(NewString, ".") + 1)
Old = Klic & ListView2.SelectedItem.Text
Nw = Klic & NewString
If strEx1 <> strEx2 Then
    Msg = MsgBox("Are you sure to exchange the file extension from: " & Chr(34) & strEx1 & Chr(34) & " to: " & Chr(34) & strEx2 & Chr(34), vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        Cancel = 0
        txtInfo.SelText = Time & " > Sending request to rename file: " & Chr(34) & Old & Chr(34) & vbCrLf
        If FtpRenameFile(server, Old, Nw) = False Then
            MsgBox "Can't accomplishe request!", vbExclamation, App.Title
            txtInfo.SelText = Time & " > request accomplished with no success at all!" & vbCrLf
            Exit Sub
        End If
        txtInfo.SelText = Time & " > File renamed to: " & Chr(34) & Nw & Chr(34) & vbCrLf
        zRemRef_Click
    Else: Cancel = 1
    End If
Else
    Cancel = 0
    txtInfo.SelText = Time & " > Sending request to rename file: " & Chr(34) & Old & Chr(34) & vbCrLf
        If FtpRenameFile(server, Old, Nw) = False Then
            MsgBox "Can't accomplish request!", vbExclamation, App.Title
            txtInfo.SelText = Time & " > Request Unsucessful!" & vbCrLf
            Exit Sub
        End If
    txtInfo.SelText = Time & " > File renamed to: " & Chr(34) & Nw & Chr(34) & vbCrLf
    zRemRef_Click
End If
Err: If Err.Number = 58 Then
MsgBox "File Aldready Exsists!", vbExclamation, App.Title
txtInfo.SelText = Time & " > Wrong file parameter!" & vbCrLf
Cancel = 1
End If
End Sub

Private Sub ListView2_BeforeLabelEdit(Cancel As Integer)
StatusBar1.Panels(2).Text = "Server: Renaming file: " & Chr(34) & ListView2.SelectedItem.Text & Chr(34)
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As ColumnHeader)
If ListView2.SortOrder = 0 Then
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.SortOrder = 1
 Else   ' Set Sorted to True to sort the list.
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.SortOrder = 0
End If
 ListView2.Sorted = True
End Sub

Private Sub ListView2_Click()
Dim i, X, d As Integer
Dim Y, z As Long
X = 0
z = 0
   For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Selected = True Then
          Y = ListView2.ListItems(i).SubItems(3)
           z = z + Y
           X = X + 1
       End If
    Next i
StatusBar1.Panels(2).Text = "Server: " & X & " Files selected, " & z / 1000 & " kb celkem"
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    zRemDS_Click
End If
End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu zRemote
End If
End Sub



Private Sub mnuDNS_Click()
Shell ("C:\Windows\System32\nslookup.exe"), vbNormalFocus
End Sub





Private Sub mnuSysTray_Click()
frmMainSys.Show
End Sub

Private Sub mnuTelnet_Click()
Shell ("C:\Windows\System32\telnet.exe"), vbNormalFocus
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Picture1.BackColor = vbRed
    Picture1.Top = Picture1.Top + Y
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
If Picture1.Top < 2500 Then Picture1.Top = 2500
If Picture1.Top > frmmain.Height - 1600 Then Picture1.Top = frmmain.Height - 1600
    Picture1.BackColor = &H8000000F
    ListView1.Top = Picture1.Top + 60
    ListView1.Height = frmmain.ScaleHeight - (Picture1.Top + 260)
    ListView2.Top = Picture1.Top + 60
    ListView2.Height = frmmain.ScaleHeight - (Picture1.Top + 260)
    Dir1.Height = Picture1.Top - Dir1.Top
    TView1.Height = Picture1.Top - TView1.Top
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Picture2.BackColor = vbRed
    Picture2.Left = Picture2.Left + X
End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
If Picture2.Left < 1980 Then Picture2.Left = 1980
If Picture2.Left > frmmain.Width - 1400 Then Picture2.Left = frmmain.Width - 1400
    Picture2.BackColor = &H8000000F
    ListView1.Width = Picture2.Left - 20
    Dir1.Width = Picture2.Left - 20
    Drive1.Width = Picture2.Left - 20
    Label2.Width = Picture2.Left - 20
    Label1.Width = Picture2.Left - 980
    ListView2.Left = Picture2.Left + 80
    ListView2.Width = frmmain.ScaleWidth - ListView2.Left
    TView1.Left = Picture2.Left + 80
    TView1.Width = frmmain.ScaleWidth - ListView2.Left
    Label3.Left = Picture2.Left + 1120
    Label3.Width = frmmain.ScaleWidth - Label3.Left
    Label4.Left = Picture2.Left + 80
    Label4.Width = frmmain.ScaleWidth - ListView2.Left
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Connect"
            FrmConnect.Show vbModal, Me
        Case "Disconnect"
            zDisconnect_Click
        Case "Reload"
            LoadLocal
        Case "Upload"
            zLokUp_Click
        Case "NewFolder"
            zLokNF_Click
        Case "DelFolder"
            zLokDF_Click
        Case "DelFile"
            zLokDS_Click
        Case "Rename"
            zLokRS_Click
        Case "Reload2"
            zRemRef_Click
        Case "DownLoad"
            zRemDown_Click
        Case "NewFolder2"
            zRemNF_Click
        Case "DelFolder2"
            zRemDF_Click
        Case "DelFile2"
            zRemDS_Click
        Case "Rename2"
            zRemRS_Click
    End Select
End Sub

Private Sub TView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    zRemDF_Click
End If
End Sub

Private Sub TView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, n As Integer
    If Not (TView1.SelectedItem Is Nothing) Then
         Set tvNode = TView1.SelectedItem
         tvNode.Expanded = True
            If tvNode.Text = ".." Then
                Klic = "/"
                adr = Klic & fPat
            Else
                Klic = tvNode.Key
                adr = Klic & fPat
            End If
                If tvNode.Children > 0 Then
                    n = tvNode.Child.Index
                    While n <> tvNode.Child.LastSibling.Index
                    TView1.Nodes.Remove (n)
                    Wend
                TView1.Nodes.Remove (n)
                End If
         FtpSetCurrentDirectory session, adr
         List
    Label3.Caption = adr
    End If
End Sub

Private Sub zAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub zAll_Click()
zAll.Checked = True
zDefine.Checked = False
File1.Pattern = "*.*"
LoadLocal
End Sub

Private Sub zAll2_Click()
Dim strPat As String
zAll2.Checked = True
zDefine2.Checked = False
fPat = "*.*"
    strPat = Mid$(adr, InStrRev(adr, "/") + 1)
    adr = Left(adr, Len(adr) - Len(strPat)) & fPat
zRemRef_Click
End Sub

Private Sub zAscii_Click()
If frmmain.zAscii.Checked = False Then
    Transfer = FTP_TRANSFER_TYPE_ASCII
    frmmain.zBinary.Checked = False
    frmmain.zAscii.Checked = True
End If
End Sub

Private Sub zBinary_Click()
If frmmain.zBinary.Checked = False Then
    Transfer = FTP_TRANSFER_TYPE_BINARY
    frmmain.zBinary.Checked = True
    frmmain.zAscii.Checked = False
End If
End Sub

Private Sub zConnect_Click()
FrmConnect.Show vbModal
End Sub

Private Sub zDefine_Click()
Dim sRet As String
    sRet = InputBox("Define file type extension (*.?):", "Pattern")
    If sRet <> "" Then
        zAll.Checked = False
        zDefine.Checked = True
        File1.Pattern = sRet
        LoadLocal
    End If
End Sub

Private Sub zDefine2_Click()
Dim sRet As String, strPat As String
    sRet = InputBox("Define file type extension (*.?):", "Pattern")
    If sRet <> "" Then
        zAll2.Checked = False
        zDefine2.Checked = True
        fPat = sRet
            strPat = Mid$(adr, InStrRev(adr, "/") + 1)
            adr = Left(adr, Len(adr) - Len(strPat)) & fPat
        zRemRef_Click
    End If
End Sub

Private Sub zDisconnect_Click()
    InternetCloseHandle server
    InternetCloseHandle session
    server = 0: session = 0
    txtInfo.SelText = Time & " > Server disconnected." & vbCrLf
End Sub

Private Sub zEnd_Click()
Unload Me
End Sub

Private Sub zFind_Click()
SHFindFiles 0, 0
End Sub

Private Sub zHelp_Click()
MsgBox "Help info not available in beta versions." & vbCrLf & "", vbInformation, "Help"
End Sub

Private Sub zLokDF_Click()
Dim Msg As VbMsgBoxResult
Dim strFol As String
Dim strCst As String

If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
    MsgBox "This is the root folder that can't be removed!", vbExclamation, App.Title
    Exit Sub
ElseIf Dir1.Path = App.Path Or Dir1.Path = App.Path & "\Logon" Then
    MsgBox "This is the program folder that can't be removed!", vbExclamation, App.Title
    Exit Sub
Else
    strFol = Mid$(Dir1.Path, InStrRev(Dir1.Path, "\") + 1)
    strCst = Left(Dir1.Path, Len(Dir1.Path) - (Len(strFol) + 1))
    Msg = MsgBox("Are you sure to remove this folder: " & Chr(34) & strFol & Chr(34) & "? Make sure it contains no files!" & vbCrLf & "If you are, don't look for it in recycle!", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        RmDir Dir1.Path
        Dir1.Path = strCst
        LoadLocal
    End If
End If
End Sub

Private Sub zLokDS_Click()
Dim Msg As VbMsgBoxResult
Dim i As Integer
If ListView1.SelectedItem Is Nothing Then
    MsgBox "Neni co vymazat!", vbExclamation
    Exit Sub
Else
    Msg = MsgBox("Are you sure to delete these files?" & vbCrLf & "These will not be in recycle!", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            Kill strPath & ListView1.ListItems(i).Text
        End If
        Next i
    LoadLocal
    End If
End If
End Sub

Private Sub zLokNF_Click()
On Error GoTo Err
    Dim sRet As String
    sRet = InputBox("Type a name of the new folder here", "New Folder")
    If sRet <> "" Then
        MkDir strPath & sRet
        Dir1.Refresh
    End If
Err: If Err.Number = 75 Then MsgBox "An error ocurred while creating the folder." & vbCrLf & "Make sure folder doesn't exist!", vbExclamation, App.Title
Exit Sub
End Sub

Private Sub zLokRef_Click()
LoadLocal
End Sub

Private Sub zLokRS_Click()
If ListView1.SelectedItem Is Nothing Then
    MsgBox "nothing to rename!", vbExclamation, App.Title
    Exit Sub
Else: ListView1.StartLabelEdit
End If
End Sub

Private Sub zLokUp_Click()
Dim i, X, d As Integer
Dim z, Y As Long
If session = 0 Or server = 0 Then
    MsgBox "You Are Not Connected! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
If ListView1.SelectedItem Is Nothing Then
MsgBox "No File Selected!"
Exit Sub
Else
X = 0
z = 0
txtInfo.SelText = Time & " > Colecting files information:" & vbCrLf
For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
          Y = ListView1.ListItems(i).SubItems(3)
          z = z + Y
           X = X + 1
               txtInfo.SelText = X & ".) " & ListView1.ListItems(i).Text & vbCrLf
               frmProg.List1.AddItem ListView1.ListItems(i).Text
               frmProg.List2.AddItem Y
               frmProg.Label2.Caption = "Soubor:       /" & X
               frmProg.Command3.Caption = "Upload"
               frmProg.Label4.Caption = "Bytes send:"
       End If
    Next i
frmProg.lbCelkem = z
frmProg.Show vbModal, Me
zRemRef_Click
End If
End Sub

Private Sub zOdpojit_Click()
Dim udtRasConn(255) As RASCONN, countConn As Long
Dim Ret As Long, b As Long

udtRasConn(0).dwSize = RAS_RASCONNSIZE
Ret = RasEnumConnections(udtRasConn(0), RAS_MAXENTRYNAME * udtRasConn(0).dwSize, countConn)
If Ret = 0 Then
    For b = 0 To countConn - 1
        Ret = RasHangUp(ByVal udtRasConn(b).hRasConn)
        If Ret = 0 Then MsgBox "Closing Connection: " & StrConv(udtRasConn(b).szEntryName(), vbUnicode), vbInformation, App.Title
    Next b
End If
End Sub

Private Sub zOpen_Click()
    Dir1.Path = App.Path & "\Logon"
    LoadLocal
End Sub

Private Sub zOpenFile_Click()
If Not ListView1.SelectedItem Is Nothing Then
ShellExecute 0, vbNullString, strPath & ListView1.SelectedItem.Text, vbNullString, strPath, SW_SHOWNORMAL
Else: MsgBox "Nothing To Open! Please Select A File.", vbExclamation, App.Title
End If
End Sub

Private Sub zPassive_Click()
    If zPassive.Checked = False Then
        zPassive.Checked = True
    Else
        zPassive.Checked = False
    End If
End Sub

Private Function CalcFTime(FTime As SYSTEMTIME) As String
Dim Datum$, Kdy$, aa$
    With FTime
      Datum = .wDay & "." & .wMonth & _
              "." & .wYear
      aa = .wMinute
      If Len(aa) = 1 Then aa = "0" & aa
      Kdy = .wHour & ":" & aa
      CalcFTime = Datum & Kdy
    End With
End Function

Private Function ImgNumber(strFileName As String) As Integer
Dim strExt As String
    strExt = Mid$(strFileName, InStrRev(strFileName, ".") + 1)
    On Error Resume Next
    Select Case LCase(strExt)
       Case "avi", "mpg", "mpeg", "mov"
            ImgNumber = 8
       Case "gif"
            ImgNumber = 4
       Case "jpg", "jpeg", "jpe", "bmp"
            ImgNumber = 1
       Case "htm", "html", "xml", "asp"
            ImgNumber = 2
       Case "js", "css", "cgi"
            ImgNumber = 5
       Case "mp3", "ram", "au", "vaw"
            ImgNumber = 6
       Case "zip", "arj"
            ImgNumber = 7
       Case "exe", "com", "bat"
           ImgNumber = 9
       Case "txt", "log", "doc", "rtf", "ftp", "ini", "dat"
           ImgNumber = 3
       Case Else
            ImgNumber = 10
    End Select
End Function
Private Sub zProperties_Click()
Dim shInfo As SHELLEXECUTEINFO
If ListView1.SelectedItem Is Nothing Then
    MsgBox "Properites Of What?"
    Exit Sub
End If
Set lsItem = ListView1.SelectedItem
    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = strPath & lsItem.Text
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With
    ShellExecuteEx shInfo
End Sub

Private Sub zRemDF_Click()
Dim Msg As VbMsgBoxResult
Dim strCst As String, strCst2 As String
Dim i As Integer
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
If Not (TView1.SelectedItem Is Nothing) Then
    Set tvNode = TView1.SelectedItem
    If tvNode.Text = ".." Then
        MsgBox "This Folder Cannot Be Removed!", vbExclamation, App.Title
        Exit Sub
    Else
        Msg = MsgBox("Are you sure to remove this folder: " & Chr(34) & tvNode.Text & Chr(34) & "? (Nesmí obsahovat žádné soubory!)" & vbCrLf & "Pokud ano, bude nenávratnì smazána!", vbQuestion + vbYesNo, App.Title)
        If Msg = vbYes Then
        strCst = Left(Klic, Len(Klic) - 1)
        strCst2 = Left(strCst, Len(strCst) - (Len(tvNode.Text)))
        txtInfo.SelText = Time & " > Sending request to remove folder: " & Chr(34) & tvNode.FullPath & Chr(34) & vbCrLf
        If FtpRemoveDirectory(server, strCst) = False Then
            MsgBox "An error occured! Make sure that folder contains no files", vbExclamation, App.Title
            txtInfo.SelText = Time & " > An error occured while removing folder!" & vbCrLf
            Exit Sub
        End If
        txtInfo.SelText = Time & " > Request Executed Sucessfully." & vbCrLf
        For i = 1 To TView1.Nodes.Count
        If TView1.Nodes(i).Key = strCst2 Then
            Set tvNode = TView1.Nodes(i)
            TView1.SelectedItem = tvNode
            TView1.SelectedItem.EnsureVisible
            Exit For
        End If
        Next i
        zRemRef_Click
        End If
    End If
End If
End Sub

Private Sub zRemDown_Click()
Dim i, X, d As Integer
Dim z, Y As Long
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
If ListView2.SelectedItem Is Nothing Then
MsgBox "No file selected!Please Select A File."
Exit Sub
Else
X = 0
z = 0
txtInfo.SelText = Time & " > Collecting information about file(s):" & vbCrLf
For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Selected = True Then
          Y = ListView2.ListItems(i).SubItems(3)
          z = z + Y
           X = X + 1
               txtInfo.SelText = X & ".) " & Klic & ListView2.ListItems(i).Text & vbCrLf
               frmProg.List1.AddItem ListView2.ListItems(i).Text
               frmProg.List2.AddItem Y
               frmProg.Label2.Caption = "File:             /" & X
               frmProg.Command3.Caption = "Download"
               frmProg.Label4.Caption = "Bytes received:"
       End If
    Next i
frmProg.lbCelkem = z
frmProg.Show vbModal, Me
LoadLocal
End If
End Sub

Private Sub zRemDS_Click()
Dim Dlt As String
Dim i As Integer
Dim Msg As VbMsgBoxResult, Cnt As Long
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
If ListView2.SelectedItem Is Nothing Then
MsgBox "No File Selected For Deletion!Please Select A File.", vbInformation, App.Title
Exit Sub
Else
    Msg = MsgBox("Are you sure you want to delete these files?", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        txtInfo.SelText = Time & " > Sending request." & vbCrLf
        For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Selected = True Then
            Dlt = Klic & ListView2.ListItems(i).Text
            txtInfo.SelText = Time & " > Deleting: " & Chr(34) & Dlt & Chr(34) & vbCrLf
            If FtpDeleteFile(server, Dlt) = False Then
                MsgBox "An error occured while deleting file!", vbExclamation, App.Title
                txtInfo.SelText = Time & " > An error occured while deleting file!" & vbCrLf
                Exit Sub
            End If
            txtInfo.SelText = Time & " > OK" & vbCrLf
        End If
        Next i
    zRemRef_Click
    End If
End If
End Sub

Private Sub zRemNF_Click()
Dim sRet As String
On Error GoTo Err
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
    sRet = InputBox("Type name for the new folder", "New Folder")
    If sRet <> "" Then
        txtInfo.SelText = Time & " > Sending request to create folder: " & Chr(34) & Klic & sRet & Chr(34) & vbCrLf
        If FtpCreateDirectory(server, Klic & sRet) = False Then
            MsgBox "An error osccured while creating folder!", vbExclamation, App.Title
            txtInfo.SelText = Time & " > An error osccured while creating folder!" & vbCrLf
            Exit Sub
        End If
        txtInfo.SelText = Time & " > Folder created." & vbCrLf
        zRemRef_Click
    End If
Err: If Err.Number = 75 Then
MsgBox "We encountered an while creating the folder!" & vbCrLf & "Make sure the folder doesn't exist", vbExclamation
txtInfo.SelText = Time & " > Wrong folder parameter." & vbCrLf
Exit Sub
End If
End Sub

Private Sub zRemRef_Click()
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
    If Not (TView1.SelectedItem Is Nothing) Then
         Set tvNode = TView1.SelectedItem
         TView1_NodeClick tvNode
    End If
End Sub

Private Sub zRemRS_Click()
If session = 0 Or server = 0 Then
    MsgBox "Not Connected To Server! Please Click The Connect Button.", vbInformation, App.Title
    Exit Sub
End If
If ListView2.SelectedItem Is Nothing Then
    MsgBox "Nothing to rename!Please select a file.", vbExclamation, App.Title
    Exit Sub
Else: ListView2.StartLabelEdit
End If
End Sub

Private Sub zSave_Click()
Dim FF As Integer
Dim Cst As String
Dim sRet As String

If txtInfo.Text <> "" Then
    sRet = InputBox("Type name for the log file:", "Log File")
    If sRet <> "" Then
    Cst = App.Path & "\" & sRet & ".txt"
        FF = FreeFile
        Open Cst For Binary As #FF
           Put #FF, , txtInfo.Text
        Close FF
        MsgBox "File saved as: " & vbCrLf & Cst, vbInformation
    Else: MsgBox "No name entered, quiting", vbInformation
    End If
Else: MsgBox "Nothing to be saved!", vbExclamation
End If
End Sub

Private Sub zsBigIc_Click()
ListView2.View = 0
End Sub

Private Sub zsReport_Click()
ListView2.View = 3
End Sub

Private Sub zsSeznam_Click()
ListView2.View = 2
End Sub

Private Sub zsSmallIc_Click()
ListView2.View = 1
End Sub

Private Sub ztBigIc_Click()
ListView1.View = 0
End Sub

Private Sub zTools_Click()
Dim Ctr As Control
If zTools.Checked = True Then
    zTools.Checked = False
    Toolbar1.Visible = False
    For Each Ctr In Controls
        If Ctr.Tag = "mov" Then
            Ctr.Top = Ctr.Top - 600
        End If
    Next Ctr
    ListView1.Height = ListView1.Height + 600
    ListView2.Height = ListView2.Height + 600
Else
    zTools.Checked = True
    Toolbar1.Visible = True
    For Each Ctr In Controls
        If Ctr.Tag = "mov" Then
            Ctr.Top = Ctr.Top + 600
        End If
    Next Ctr
    ListView1.Height = ListView1.Height - 600
    ListView2.Height = ListView2.Height - 600
End If
End Sub

Private Sub ztReport_Click()
ListView1.View = 3
End Sub

Private Sub ztSeznam_Click()
ListView1.View = 2
End Sub

Private Sub ztSmallIc_Click()
ListView1.View = 1
End Sub
