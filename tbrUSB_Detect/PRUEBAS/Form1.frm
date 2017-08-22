VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   6405
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   6405
      Left            =   5850
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents DR As tbrDRIVES.clsDRIVES
Attribute DR.VB_VarHelpID = -1

Private Sub DR_IngresaDrive(LetterUnit As String, SBT As Long)
    Text1.Text = Text1.Text + CStr(Timer) + " Se AGREGO al unidad: " + _
        LetterUnit + " (" + GetStrSBT(SBT) + ")" + vbCrLf
    
    
    UpdateList
End Sub

Private Sub DR_SaleDrive(LetterUnit As String, SBT As Long)
    Text1.Text = Text1.Text + CStr(Timer) + _
        " Se QUITO al unidad: " + LetterUnit + " (" + GetStrSBT(SBT) + ")" + vbCrLf
    UpdateList
End Sub

Private Sub Form_Load()
    Set DR = New tbrDRIVES.clsDRIVES
    
    DR.SoloDispositivosUSB = True
    
    DR.Iniciar Me
    
    UpdateList
End Sub

Private Sub UpdateList()
    Dim H As Long
    Text2.Text = ""
    For H = 1 To DR.GetCantidadUSB
        Text2.Text = Text2.Text + DR.GetDriveList(H) + " " + DR.GetDriveInfo(H) + vbCrLf + vbCrLf
    Next H
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DR.Terminar
End Sub

Private Function GetStrSBT(i As Long) As String
    Select Case i
        Case 0: GetStrSBT = "BusTypeUnknown"
        Case 1: GetStrSBT = "BusTypeScsi"
        Case 2: GetStrSBT = "BusTypeAtapi"
        Case 3: GetStrSBT = "BusTypeAta"
        Case 4: GetStrSBT = "BusType1394"
        Case 5: GetStrSBT = "BusTypeSsa"
        Case 6: GetStrSBT = "BusTypeFibre"
        Case 7: GetStrSBT = "BusTypeUsb"
        Case 8: GetStrSBT = "BusTypeRAID"
        Case &H7F: GetStrSBT = "BusTypeMaxReserved"
        Case Else: GetStrSBT = "ERROR!"
    End Select

End Function

'DE ESTA FORMA LOS EVENTOS SE RECIBEN SOLO EN UN FORMULARIO
'SI SE NECESITAN MAS USAR
'DR.UseEventMSG Text.hWnd
'para que mande textos a esa caja de texto
