VERSION 5.00
Object = "*\A..\Proyecto1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   4890
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2430
      Width           =   2265
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   4890
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Width           =   2265
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4890
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1650
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   1620
      TabIndex        =   5
      Top             =   1530
      Width           =   3195
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   2265
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   2265
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2265
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activar Efectos"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   0
         Width           =   1545
      End
   End
   Begin tbrEspectroPT.tbrEspectro ESP 
      Height          =   4365
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   7699
      xDisp           =   0
      xModo           =   1
      xFX             =   -1  'True
      xMxL            =   0
      xBCol           =   8388736
      xCol            =   65535
      xSense          =   1
      xDoCls          =   -1  'True
      xBlur           =   -1  'True
      xLuz            =   -1  'True
      xLineas         =   -1  'True
      xRefreshRate    =   25
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   3360
      TabIndex        =   3
      Top             =   30
      Width           =   3795
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OFF"
      Height          =   825
      Left            =   750
      TabIndex        =   1
      Top             =   1500
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   825
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    ESP.FX_Enabled = Check1.Value
    
    Combo1.Enabled = Check1
    Combo2.Enabled = Check1
    Combo3.Enabled = Check1
End Sub

Private Sub Combo1_Click()
    ESP.Blur = (Combo1.ListIndex > 0)
End Sub

Private Sub Combo2_Click()
    ESP.Lineas = (Combo2.ListIndex > 0)
End Sub

Private Sub Combo3_Click()
    ESP.Luz = (Combo3.ListIndex > 0)
End Sub

Private Sub Combo4_Click()
    Select Case Combo4.ListIndex
        Case 0: ESP.Sensibilidad = 0.5
        Case 1: ESP.Sensibilidad = 1
        Case 2: ESP.Sensibilidad = 1.5
        Case 3: ESP.Sensibilidad = 2
        Case 4: ESP.Sensibilidad = 3
    End Select
End Sub

Private Sub Combo5_Click()
    ESP.Modo = Combo5.ListIndex + 1
End Sub

Private Sub Combo6_Click()
    Select Case Combo6.ListIndex
        Case 0: ESP.RefreshRate = 10
        Case 1: ESP.RefreshRate = 20
        Case 2: ESP.RefreshRate = 40
        Case 3: ESP.RefreshRate = 80
        Case 4: ESP.RefreshRate = 160
        Case 5: ESP.RefreshRate = 500
    End Select
End Sub

Private Sub Command1_Click()
    ESP.Comenzar
End Sub

Private Sub Command2_Click()
    ESP.Detener
End Sub

Private Sub Form_Load()
    
    Combo1.AddItem "Blur Desactivado"
    Combo1.AddItem "Blur Activado"
    Combo1.ListIndex = 0
    
    Combo2.AddItem "Lineas Desactivado"
    Combo2.AddItem "Lineas Activado"
    Combo2.ListIndex = 0
    
    Combo3.AddItem "Luz Desactivado"
    Combo3.AddItem "Luz Activado"
    Combo3.ListIndex = 0
    
    Combo4.AddItem "Sensibilidad x 0.50"
    Combo4.AddItem "Sensibilidad x 1.00"
    Combo4.AddItem "Sensibilidad x 1.50"
    Combo4.AddItem "Sensibilidad x 2.00"
    Combo4.AddItem "Sensibilidad x 3.00"
    Combo4.ListIndex = 1
    
    Combo5.AddItem "Modo 1"
    Combo5.AddItem "Modo 2"
    Combo5.ListIndex = 0
    
    Combo6.AddItem "Refresh Rate 10 Ms"
    Combo6.AddItem "Refresh Rate 20 Ms"
    Combo6.AddItem "Refresh Rate 40 Ms"
    Combo6.AddItem "Refresh Rate 80 Ms"
    Combo6.AddItem "Refresh Rate 160 Ms"
    Combo6.AddItem "Refresh Rate 500 Ms"
    Combo6.ListIndex = 1
    
    Check1.Value = 0
    
    ESP.CargoDispositivos List1
    ESP.CargoLineas List2
    
    List1.ListIndex = 0
    List2.ListIndex = 0

End Sub

Private Sub List1_Click()
    ESP.Dispositivo = List1.ListIndex
End Sub

Private Sub List2_Click()
    ESP.MIXERLINE = List2.ListIndex
End Sub
