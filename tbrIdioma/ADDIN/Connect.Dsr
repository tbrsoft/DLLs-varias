VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13095
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15840
   _ExtentX        =   27940
   _ExtentY        =   23098
   _Version        =   393216
   Description     =   "Utilice este asistente para traducir cualquier proyecto."
   DisplayName     =   "Complemento traductor de proyectos"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public vbinstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'controlador de evento de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.vbinstance = vbinstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
End Sub

'------------------------------------------------------
'este método agrega el complemento a VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'guardar la instanacia de vb
    Set vbinstance = Application
'    Set frmaddin.vbInstance = Application
'    Set frmaddin.Connect = Me
'    'éste es un buen lugar para establecer un punto de interrupción y
    'y probar varios objetos, propiedades y métodos de complemento
'    MsgBox vbInstance.FullName
'
'    If ConnectMode = ext_cm_External Then
'        'Utilizado por la barra de herramientas de asistente para iniciar este asistente
'        Me.Show
'   Else
'
         vbinstance.CommandBars.Add "Martin"
         vbinstance.CommandBars("Martin").Visible = True
         'Set mcbMenuCommandBar = AddToAddInCommandBar("Mi complemento")
         Set mcbMenuCommandBar = vbinstance.CommandBars("Martin").Controls.Add
         mcbMenuCommandBar.BeginGroup = True
         mcbMenuCommandBar.Caption = "Generador de sentencias SQL"
         Clipboard.SetData LoadResPicture(101, 0)
         mcbMenuCommandBar.PasteFace
         mcbMenuCommandBar.Visible = True
'        recibir el evento
        Set Me.MenuHandler = vbinstance.Events.CommandBarEvents(mcbMenuCommandBar)
'    End If
'
'    If ConnectMode = ext_cm_AfterStartup Then
'        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'            'establecer esto para mostrar el formulario al conectar
            Me.Show
'        End If
'    End If
'frmaddin.Show
    Exit Sub

error_handler:

    MsgBox Err.Description + "cargando"
    
End Sub

'------------------------------------------------------
'este método quita el complemento de VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'eliminar la entrada de la barra de comandos
    mcbMenuCommandBar.Delete
    
    'cerrar el complemento
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'establecer esto para mostrar el formulario al conectar

        Me.Show
    End If
End Sub

'este evento se desencadena cuando se hace clic en el menú desde el IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'objeto de barra de comandos
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'ver si podemos encontrar el menú Complementos
    Set cbMenu = vbinstance.CommandBars("Complementos")
    If cbMenu Is Nothing Then
        'no disponible; error
        Exit Function
    End If
    
    'agregarlo a la barra de comandos
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'establecer el título
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

