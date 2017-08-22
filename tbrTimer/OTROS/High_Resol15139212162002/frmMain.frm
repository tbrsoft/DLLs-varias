VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FormLess Timer - Demo"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   3090
      TabIndex        =   13
      ToolTipText     =   "Clears the above event log"
      Top             =   2550
      Width           =   2955
   End
   Begin VB.TextBox txtResults 
      Alignment       =   2  'Center
      Height          =   2535
      Left            =   3090
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      ToolTipText     =   "This displays the Timer Event log"
      Top             =   0
      Width           =   2955
   End
   Begin VB.Frame fraTimer0 
      Caption         =   "Timer &2"
      Height          =   885
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   1890
      Width           =   3015
      Begin VB.TextBox txtInterval 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "This is the Milliseconds that the timer will yield for it's delay duration."
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Toggle Timer"
         Height          =   525
         Index           =   2
         Left            =   1635
         TabIndex        =   11
         ToolTipText     =   "Click to toggle the Timers Enabled Property"
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblInterval 
         Caption         =   "Interval:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraTimer0 
      Caption         =   "Timer &1"
      Height          =   885
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   960
      Width           =   3015
      Begin VB.TextBox txtInterval 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "This is the Milliseconds that the timer will yield for it's delay duration."
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Toggle Timer"
         Height          =   525
         Index           =   1
         Left            =   1635
         TabIndex        =   7
         ToolTipText     =   "Click to toggle the Timers Enabled Property"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblInterval 
         Caption         =   "Interval:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraTimer0 
      Caption         =   "Timer &0"
      Height          =   885
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3015
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Toggle Timer"
         Height          =   525
         Index           =   0
         Left            =   1635
         TabIndex        =   3
         ToolTipText     =   "Click to toggle the Timers Enabled Property"
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "This is the Milliseconds that the timer will yield for it's delay duration."
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblInterval 
         Caption         =   "Interval:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'** This form exists only to provide a means to show you the Timers functionality. _
    In no way dose the Timer object require any refrences to this form.

'** Define Timer Objects
    Private WithEvents MyTimer0 As tbrhTimer.hTimerCls
Attribute MyTimer0.VB_VarHelpID = -1
    Private WithEvents MyTimer1 As tbrhTimer.hTimerCls
Attribute MyTimer1.VB_VarHelpID = -1
    Private WithEvents MyTimer2 As tbrhTimer.hTimerCls
Attribute MyTimer2.VB_VarHelpID = -1

'************************ _
 ** Form Events: Start ** _
 ************************

    Private Sub Form_Load()

        '** Create Timer Objects
            Set MyTimer0 = New tbrhTimer.hTimerCls
            Set MyTimer1 = New tbrhTimer.hTimerCls
            Set MyTimer2 = New tbrhTimer.hTimerCls
        
        '** Set The initial values for the Interval Text Boxes
            txtInterval(0).Text = MyTimer0.Interval
            txtInterval(1).Text = MyTimer1.Interval
            txtInterval(2).Text = MyTimer2.Interval
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        '** Clean Up
            Set MyTimer0 = Nothing
            Set MyTimer1 = Nothing
            Set MyTimer2 = Nothing
    End Sub

'********************** _
 ** Form Events: End ** _
 **********************

Private Sub cmdToggle_Click(Index As Integer)
    '** Toggle Enabled state of the Timer(s)
        
        Select Case Index
            Case 0
                With MyTimer0
                    .Enabled = Not .Enabled
                End With
            Case 1
                With MyTimer1
                    .Enabled = Not .Enabled
                End With
            Case 2
                With MyTimer2
                    .Enabled = Not .Enabled
                End With
        End Select
End Sub

Private Sub txtInterval_Change(Index As Integer)
    '** Set the interval for the timer
        Select Case Index
            Case 0
                MyTimer0.Interval = Val(txtInterval(0).Text)
            Case 1
                MyTimer1.Interval = Val(txtInterval(1).Text)
            Case 2
                MyTimer2.Interval = Val(txtInterval(2).Text)
        End Select
End Sub

'************************* _
 ** Timer Events: Start ** _
 *************************

    Private Sub MyTimer0_Timer()
        '** Display Proof of Event in Results Log
            With txtResults
                .Text = .Text & "Timer0 Event Raised" & vbCrLf
            End With
    End Sub
    
    Private Sub MyTimer1_Timer()
        '** Display Proof of Event in Results Log
            With txtResults
                .Text = .Text & "Timer1 Event Raised" & vbCrLf
            End With
    End Sub
    
    Private Sub MyTimer2_Timer()
        '** Display Proof of Event in Results Log
            With txtResults
                .Text = .Text & "Timer2 Event Raised" & vbCrLf
            End With
    End Sub

'*********************** _
 ** Timer Events: End ** _
 ***********************

Private Sub txtResults_Change()
    '** Force Scroll to bottom of the Results textbox
        txtResults.SelStart = Len(txtResults.Text)
End Sub

Private Sub cmdClearLog_Click()
    '** Clear Results Log
        txtResults.Text = ""
End Sub
