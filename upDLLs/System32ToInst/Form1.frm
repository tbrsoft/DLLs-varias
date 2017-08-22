VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      BackColor       =   &H00C0C0FF&
      Columns         =   5
      Height          =   2400
      Left            =   30
      TabIndex        =   2
      Top             =   3180
      Width           =   11265
   End
   Begin VB.ListBox List1 
      Columns         =   5
      Height          =   2400
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   11265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "comenzar"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim F As String
    Dim FSO As New Scripting.FileSystemObject
    F = Dir("c:\windows\system32\tbr*.*")
    
    Dim ORIG As String, DEST As String, ReemplaZar As Boolean
    
    Do While F <> ""
        If Len(F) > 3 Then
            List1.AddItem F
            
            ORIG = "c:\windows\system32\" + F
            DEST = "C:\Archivos de programa\Inno Setup 5\vbFiles\other\" + F
            ReemplaZar = False
            
            If FSO.FileExists(DEST) = False Then
                ReemplaZar = True
            Else 'siolo si eciste se puede comparar !
                'choreado del copiseg
                If FileDateTime(ORIG) <> FileDateTime(DEST) Then ReemplaZar = True
                If FileLen(ORIG) <> FileLen(DEST) Then ReemplaZar = True
            End If
            
            If ReemplaZar Then
                List2.AddItem "Reemplazado: " + F
                FSO.CopyFile ORIG, DEST, True
            End If
            
        End If
        F = Dir
    Loop
    
    MsgBox "termino!"
End Sub
