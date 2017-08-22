VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form1 
   Caption         =   "Very Easy FTP"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Download"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upload"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemoteHost      =   "www.0catch.com"
      RemotePort      =   21
      URL             =   "ftp://helloindia.0catch.com@www.0catch.com"
      UserName        =   "helloindia.0catch.com"
      Password        =   "march23"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you like this code, you can thank my by visiting my sites...
    'www.cupidsystems.com = Free Software Downloads
    'www.vexat.net = Free Tutorials
    'www.bnetsupport.com = Free Ebooks



Private Sub Command1_Click()
'Uploading file to server
'PUT method is a command in FTP to upload file to a server
Inet1.Execute , "PUT  ""c:\index.html""" & "index.html"
MsgBox Inet1.ResponseInfo
End Sub

Private Sub Command2_Click()
'Downloading a file from server
'GET method is a command in FTP to download file from a server
Inet1.Execute , "GET  ""index.html""" & "c:\index.html"
End Sub

Private Sub Form_Load()
'Server Access type
Inet1.AccessType = icUseDefault

'Protocol to be used
Inet1.Protocol = icFTP

'Remote host name
Inet1.RemoteHost = "ftp.yahoo.com"

'Server port number, usually its 21
Inet1.RemotePort = "21"

'Server password
Inet1.Password = "password"

'Server Username
Inet1.UserName = "username"

'Server session timeout
Inet1.RequestTimeout = "60"
End Sub
