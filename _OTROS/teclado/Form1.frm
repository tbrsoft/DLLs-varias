VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simulador de teclado"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTECLAS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      ItemData        =   "Form1.frx":0000
      Left            =   180
      List            =   "Form1.frx":012D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   10515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Presione los botones o cables del simulador de teclado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   420
      TabIndex        =   1
      Top             =   150
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim S As String
    S = CStr(KeyCode)
    Dim A As Long
    For A = 0 To cmbTECLAS.ListCount - 1
        If S = Left(cmbTECLAS.List(A), Len(S)) Then
            cmbTECLAS.ListIndex = A
            Exit For
        End If
    Next A
End Sub

