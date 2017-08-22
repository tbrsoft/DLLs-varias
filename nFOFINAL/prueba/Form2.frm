VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "tbrSoft comunicacion"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "agregar email"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   6
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "agregar telefono"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2100
      TabIndex        =   5
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "eliminar elegido"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1080
      TabIndex        =   4
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "editar elegido"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   3
      Top             =   3270
      Width           =   975
   End
   Begin VB.ListBox lstDatos 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2940
      IntegralHeight  =   0   'False
      ItemData        =   "Form2.frx":0000
      Left            =   60
      List            =   "Form2.frx":001F
      TabIndex        =   2
      Top             =   300
      Width           =   4905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar el Archivo para enviar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6630
      TabIndex        =   1
      Top             =   780
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar archivo para enviar a tbrSoft"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6630
      TabIndex        =   0
      Top             =   180
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "datos del registro"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   2025
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

