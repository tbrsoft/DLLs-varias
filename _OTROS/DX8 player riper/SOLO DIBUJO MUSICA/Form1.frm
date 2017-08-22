VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2640
      Top             =   2010
   End
   Begin VB.PictureBox PicVis 
      Height          =   3765
      Left            =   510
      ScaleHeight     =   3705
      ScaleWidth      =   5205
      TabIndex        =   0
      Top             =   300
      Width           =   5265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()

    Dim intSamples(FFT_SAMPLES - 1) As Integer

    If clsPlayer.bitspersample = 16 Then
        clsPlayer.CaptureSamples VarPtr(intSamples(0)), FFT_SAMPLES * 2

        Select Case udeVis
            Case VIS_BARS
                modDraw.DrawFrequencies intSamples, PicVis
            Case VIS_OSC
                modDraw.DrawOsc intSamples, PicVis
            Case VIS_PEAKS
                modDraw.DrawPeaks intSamples, PicVis
            Case VIS_NONE
                PicVis.Cls
        End Select
    End If

End Sub
