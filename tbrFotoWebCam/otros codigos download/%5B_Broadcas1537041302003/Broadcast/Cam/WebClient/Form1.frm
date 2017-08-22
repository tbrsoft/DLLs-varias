VERSION 5.00
Object = "{D3BBDB60-9B18-4FBC-9A90-CCFBF4F8D491}#65.0#0"; "AVPhone3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5928
   ClientLeft      =   4440
   ClientTop       =   1776
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Height          =   336
      Left            =   2052
      TabIndex        =   1
      Top             =   180
      Width           =   1596
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   336
      Left            =   288
      TabIndex        =   0
      Top             =   180
      Width           =   1596
   End
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   1080
      Top             =   2772
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0000
   End
   Begin AVPhone3.AudRnd AudRnd1 
      Left            =   2340
      Top             =   3060
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0024
   End
   Begin AVPhone3.VidRnd VidRnd1 
      Height          =   1848
      Left            =   324
      Top             =   684
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   3260
      Control         =   "Form1.frx":0048
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    UDPSocket1.SetSendAddress InputBox("Enter remote name or IP:", "Connect to server", UDPSocket1.GetIP(UDPSocket1.LocalAddress))
    UDPSocket1.Frame 0, TM_CONNECT
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub Disconnect()
    UDPSocket1.Frame 0, TM_DISCONNECT

    StopRender
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorHandle
    Disconnect
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub Form_Load()
    Caption = "Broadcast client"
    
    UDPSocket1.Bind 1721, 1720
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    Disconnect
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub StopRender()
    AudRnd1.Format = vbNullString
    VidRnd1.Format = vbNullString
End Sub

Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal lParam As Long, vData As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DISCONNECT
        StopRender
        
    Case TM_VIDEOFORMAT
        VidRnd1.Format = vData
    Case TM_AUDIOFORMAT
        AudRnd1.Format = vData
        
    Case TM_VIDEOFRAME
        VidRnd1.Frame vData
    Case TM_AUDIOFRAME
        AudRnd1.Frame vData
        
    End Select
    Exit Sub
    
ErrorHandle:
End Sub
