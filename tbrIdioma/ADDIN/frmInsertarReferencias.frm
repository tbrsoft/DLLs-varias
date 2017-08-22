VERSION 5.00
Begin VB.Form frmInsertarReferencias 
   Caption         =   "Insertar referencias y clase traductor"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optModulo 
      Caption         =   "Modulo Existente"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   2640
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.TextBox txtNombreModulo 
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Text            =   "Modulo"
      ToolTipText     =   "Escriba el nombre del modulo a insertar"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.OptionButton optModulo 
      Caption         =   "Nuevo Modulo"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ListBox lstModulos 
      Height          =   1035
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Seleccione el modulo donde se insertara la declaracion del traductor"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CheckBox chkDeclaracion 
      Caption         =   "Insertar declaracion del traductor"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CheckBox chkReferencia 
      Caption         =   "Insertar Referencia a ""Microsoft Scripting Runtime"" (Necesaria para el traductor)"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.CheckBox chkTraductor 
      Caption         =   "Insertar clase traductor"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ListBox lst 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9480
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9480
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione el proyecto donde quiere incluir el traductor."
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmInsertarReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'se podria sincronizar el option modulo existente,
'si hay modulos se lo habilita, sino se setea el option nuevo modulo

Dim vbinstance As VBIDE.VBE

Public Sub LlenarFormulario(pVBinstance As VBIDE.VBE)
    Set vbinstance = pVBinstance
    Me.Show
    Dim vbPro As VBProject
    
    For Each vbPro In vbinstance.VBProjects
        lst.AddItem vbPro.Name
    Next
    'para q se actualice el estado de los option, etc
    actualizarEstadoOpciones
    
End Sub

Private Sub actualizarEstadoOpciones()
    'al checkReferencia no lo toco desde aca
    
    'la opcion modulo nuevo y existente dependen del check declaracion
    optModulo(0).Enabled = IIf(chkDeclaracion.Value = vbChecked, True, False)
    optModulo(1).Enabled = optModulo(0).Enabled
    
    'si alguno de los option esta habilitado...
    If optModulo(0).Enabled Then
    
        'si esta vacia la lista de modulos desabilito el option modulo existente
        If lstModulos.ListCount = 0 Then
            optModulo(0).Value = True
            optModulo(1).Value = False
            optModulo(1).Enabled = False
        End If
    
    End If
     'dependen del valor de los option
    txtNombreModulo.Enabled = optModulo(0).Value And optModulo(0).Enabled
    lstModulos.Enabled = optModulo(1).Value And optModulo(1).Enabled
    
End Sub

Private Sub chkDeclaracion_Click()
    actualizarEstadoOpciones
End Sub

Private Sub cmdAceptar_Click()
    On Error Resume Next 'si hay error lo siento mucho
    If datosCorrectos Then
        Dim vPro As VBProject
        Set vPro = vbinstance.VBProjects(lst.Text)
        
        'aca inserto la referencia
        If chkReferencia.Value = vbChecked Then
            InsertarReferencia vPro
        End If
        
        'aca inserto el traductor
        If chkTraductor.Value = vbChecked Then
            vPro.VBComponents.AddFromTemplate App.path + "\traductor.cls"
        End If
        
        'aca inserto la declaracion en el modulo elegido o en uno nuevo
        If chkDeclaracion.Value = vbChecked Then
            'si eligio modulo nuevo...
            Dim vbComp As VBComponent
            
            If optModulo(0).Value Then
                'habria q mejorar esta validacion
                Set vbComp = vPro.VBComponents.Add(vbext_ct_StdModule)
                vbComp.Name = IIf(txtNombreModulo = "", "Modulo", txtNombreModulo.Text)
                
            Else 'sino eligio uno existente, aca seguro llega con alguno elegido
                Set vbComp = vPro.VBComponents(lstModulos.Text)
            End If
            
            vbComp.CodeModule.InsertLines vbComp.CodeModule.CountOfDeclarationLines + 1, "Public Trans As New Translator"
        End If
    Unload Me
    End If
    
End Sub

Private Sub InsertarReferencia(pVBProject As VBProject)
    'me fijo por las dudas si ya esta incluida aunque ya me fije cuando selecciona el proyecto
    On Error GoTo e
    Dim vScrip As New Scripting.FileSystemObject
    Dim ref As Reference
    'SpecialFolderConst
    Set ref = getReferencia(pVBProject, "Scripting")
    If ref Is Nothing Then
        'aca verificar la ruta de windows
        pVBProject.References.AddFromFile vScrip.GetSpecialFolder(SystemFolder) + "\scrrun.dll"
    End If
    Exit Sub
e:
End Sub

Private Function getReferencia(pVBProject As VBProject, pNombre As String) As Reference
On Error GoTo e
Set getReferencia = pVBProject.References.Item(pNombre)

Exit Function
e:
Set getReferencia = Nothing
End Function

Private Function datosCorrectos() As Boolean
Dim msg As String

If lst.ListIndex = -1 Then msg = "Debe seleccionar un proyecto!"
If chkDeclaracion.Value = vbChecked Then If lst.ListIndex = -1 Then msg = "Debe seleccionar un Modulo para insertar la declaracion!"

If msg = "" Then
    datosCorrectos = True
Else
    MsgBox "Falta alguna informacion necesaria!" + vbCrLf + vbTab + msg, vbExclamation
    datosCorrectos = False
End If
End Function

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub lst_Click()
If lst.ListIndex <> -1 Then
    lstModulos.Clear
    Dim vPro As VBProject
    Set vPro = vbinstance.VBProjects(lst.Text)
    Dim vbComp As VBComponent
    For Each vbComp In vPro.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then lstModulos.AddItem vbComp.Name
    Next
    
    actualizarEstadoOpciones
    
    'aca me fijo si la referencia ya existe para setear el valor del checkReferencia
    Dim ref As Reference
    Set ref = getReferencia(vPro, "Scripting")
    If Not ref Is Nothing Then
        'es porq ya esta incluida
        chkReferencia.Value = vbGrayed
    Else
        'permito la seleccion
        chkReferencia.Value = vbChecked
    End If
    
End If
End Sub

Private Sub optModulo_Click(Index As Integer)
    actualizarEstadoOpciones
End Sub
