VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Uninstaller Standard"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   4125
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7276
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim mIP As New tbrInstalledPrograms.tbrProgsInst
    
    LV.View = lvwReport
    LV.ListItems.Clear
    
    LV.ColumnHeaders.Add , , "Software"
    LV.ColumnHeaders.Add , , "Path"
    LV.ColumnHeaders.Add , , "HelpLink"
    LV.ColumnHeaders.Add , , "InstallDate"
    LV.ColumnHeaders.Add , , "Publisher"
    LV.ColumnHeaders.Add , , "URLInfoAbout"
    LV.ColumnHeaders.Add , , "URLUpdateInfo"
    
    'leer los programas que hay
    mIP.LoadList
    
    Dim J As Long, LI As ListItem
    For J = 1 To mIP.Cantidad
        Set LI = LV.ListItems.Add(, , mIP.GetName(J))
        LI.ListSubItems.Add , , mIP.GetPath(J)
        LI.ListSubItems.Add , , mIP.GetHelpLink(J)
        LI.ListSubItems.Add , , mIP.GetInstallDate(J)
        LI.ListSubItems.Add , , mIP.GetPublisher(J)
        LI.ListSubItems.Add , , mIP.GetURLInfoAbout(J)
        LI.ListSubItems.Add , , mIP.GetURLUpdateInfo(J)
    Next J
    
    'buscador específico
    MsgBox mIP.GetPath2("3pm", "tbrsoft")
    
End Sub
