VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim graph As FilgraphManager
    Set graph = New FilgraphManager
    
    Dim gh As VBGraphHelper
    Set gh = New VBGraphHelper
    gh.graph = graph
    Dim f As IFilterInfo
    Set f = gh.FilterByClsid("{51B4ABF3-748F-4E3B-A276-C828330E926A}", "VMR9")
   
    Dim vmr As VMRConfigInfo
    Set vmr = New VMRConfigInfo
    vmr.SetFilter f
    vmr.NumberOfStreams = 2
        
    For i = 0 To 1 Step 1
        vmr.alpha(i) = 0.5
        vmr.ZOrder(i) = i
        vmr.SetOutputRect i, 0, 0, 1, 1
    Next i
           
    Dim alpha As Double
    alpha = vmr.alpha(0)
    
End Sub
