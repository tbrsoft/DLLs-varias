VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traductor de proyectos"
   ClientHeight    =   2805
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1590
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdInsertar 
      Caption         =   "Insertar Referencias y Traductor"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   645
      Width           =   3255
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Documento a traducir"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Ejecutar despues de Usar el asistente"
      Top             =   1140
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   2310
      Width           =   3255
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Asistente para Traducir"
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar PB2 
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1830
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar PB3 
      Height          =   135
      Left            =   180
      TabIndex        =   7
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Height          =   2565
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vbinstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub CancelButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "Cierra el Traductor"
End Sub

Private Sub cmdGenerar_Click()
    
    'On Error Resume Next
    
    Dim I As Integer
    Dim vPro As VBProject
    Dim vComp As VBComponent
    Dim cadena As String
    Dim palabra As String
    Dim aux() As String
    
    Dim dic As New Dictionary
    
    Dim a As Long
    Dim b As Long
    Dim c As Long
    
    'antes Trans("texto español")
    
    'nueva version de trans. Debe ser compatible con multimples renglones
    'Trans ("texto en idioma %00% con "+ _
        "&01& caracteres de texto "+ vbcrlf+ _
        "%98%Aqui se fija el idioma la variable 01 es un texto"+ vbcrlf+ _
        "como 'español' o 'frances' y el segundo es"+ _
        "la cantidad de caracteres%99%",cmbidioma.text, sCantCararcteres)
    
    'donde puede haber variables string en %00% %01% .... y en %98% _
        indico que termina la cadena y entre %98% y %99% dejo el _
        "contexto" o "descripcion".
    
    Dim pr1, pr2, pr3
    
    PB1.Max = vbinstance.VBProjects.Count
    
    
    For Each vPro In vbinstance.VBProjects
        pr1 = pr1 + 1: PB1.Value = pr1
        
        PB2.Max = vPro.VBComponents.Count
        For Each vComp In vPro.VBComponents
            pr2 = pr2 + 1: PB2.Value = pr2
            
            'una vuelta por cada formulario
            'si es que tiene lineas !!!
            If vComp.CodeModule.CountOfLines = 0 Then GoTo nxt
            cadena = vComp.CodeModule.Lines(1, vComp.CodeModule.CountOfLines)
            'todos los renglones en "cadena"
            a = 0
            PB3.Max = 100
            Do 'cada uno de los inicio de traduccion
                pr3 = pr3 + 1
                PB3.Value = pr3 Mod 100
               'a = InStr(a + 1, cadena, "Trans(")
                a = InStr(a + 1, cadena, ".Trad(")
                
                'TODO - SEGUIRAQUI si la linea esta comentada NO tiene que ir ya que el traductor trabaja al pedo
                'como es LUIS no importa, quelabure
                
                'cuando llegue al ultimo me voy
                If a = 0 Then Exit Do
                b = a
                PB1.Value = 0
                
                Dim CADENAS As Long 'cantidad de cadenas totales
                Dim CADENASF As Long 'cantidad de cadenas aceptadas
                
                Do 'busco el final de lo comenzado
                    b = InStr(b, cadena, "%99%")
                    If (b = 0) Then
                        Exit Do
                    Else
                        b = b + 3 'que llegue hasta el ultimo "%"
                        palabra = Mid(cadena, a + 7, b - a - 6)
                        
                        'ahora quitar todos los altos de carro y concatenaciones extrañas
                        'debe ser un chorizo de un renglon !!
                        
                        '************************************************************
                        'EL MAXIMO DE CARACTERES POR LINEA DE WINDOWS O DE VB ES 1275
                        '************************************************************
                        If Len(palabra) > 1000 Then
                            MsgBox "La cadena" + vbCrLf + palabra + vbCrLf + vbCrLf + "DEBE SER CORTADA, NO ENTRA"
                            Exit Sub
                        End If
                        palabra = Enchorizar(palabra)
                        
                        'RRRRR que se agrege a la segunda parte solo el texto sin ayuda!!
                        
'                        If Len(palabra) > 350 Then
'                            If MsgBox(palabra, vbYesNo) = vbNo Then
'                                Exit Sub
'                            End If
'                        End If
'
                        If InStr(palabra, ".Trad") Then
                            If MsgBox("Esto parece mal!" + vbCrLf + palabra + vbCrLf + vbCrLf + "Seguis?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        If InStr(palabra, vbCrLf) Then
                            If MsgBox("Esto parece mal!" + vbCrLf + palabra + vbCrLf + vbCrLf + "Seguis?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        CADENAS = CADENAS + 1
                        
                        'que no se repitan cadenas!!
                        If dic.Exists(palabra) = False Then
                            CADENASF = CADENASF + 1
                            dic.Add palabra, palabra
                        End If
                        
                        a = a + 5
                        Exit Do
                    End If
                Loop
            Loop
nxt:
        Next
    
    Next
    
    Dim contenidoArchivo As String
    For I = 0 To dic.Count - 1
        contenidoArchivo = contenidoArchivo + Chr$(30) + dic.Keys(I) + Chr$(31) + dic.Items(I)
    Next
    
    contenidoArchivo = Right$(contenidoArchivo, Len(contenidoArchivo) - 1)
    
    Dim cd As New CommonDialog
    cd.DialogTitle = "Grabar las cadenas (" + CStr(CADENASF) + "/" + CStr(CADENAS) + ")"
    cd.DefaultExt = ".lanSource"
    cd.Filter = "Fuente de idioma (*.lanSource)|*.lansource"
    cd.ShowSave
    If cd.FileName <> "" Then
        EscribirArchivo cd.FileName, contenidoArchivo
    End If
End Sub

Private Function Enchorizar(S As String) As String
    'toma una cadena de codigo de visual basic o con saltos de carro
    ' o concatenaciones multirenglon y la deja joia de un solo renglón
    
    Dim T As String 'temporal
    
    T = Replace(S, Chr(13), "")  'matar primero que todo los saltos de carro
    'cuando suben las lineas segun la cantidad de tabulaciones existentes en la linea
    'de abajo habra una diferente cantida de espacios
    
    T = Replace(T, vbCr, "")  'matar primero que todo los saltos de carro
    T = Replace(T, vbLf, "")  'matar primero que todo los saltos de carro
    
    Dim H As Long
    For H = 0 To 100
        'SACAR LOS                             " + vbcrlf + _"
        'buscando cualquier cantidad de espacios
        T = Replace(T, Chr(34) + " + vbCrLf + _" + Space(H) + Chr(34), "%97%")
        'el %97% representara los saltos de carro
    Next H
    
    'tambien aparecen con 2 saltos de carro
    For H = 0 To 100
        'SACAR LOS                             " + vbcrlf + vbcrlf + _"
        'buscando cualquier cantidad de espacios
        T = Replace(T, Chr(34) + " + vbCrLf + vbCrLf + _" + Space(H) + Chr(34), "%97%")
        'el %97% representara los saltos de carro
    Next H
    
    For H = 0 To 100
        'SACAR LOS                             " + _"
        'buscando cualquier cantidad de espacios
        T = Replace(T, Chr(34) + " + _" + Space(H) + Chr(34), "")
    Next H
    
    T = Replace(T, Chr(34) + " + " + Chr(34), "")
    T = Replace(T, Chr(34) + " + vbCrLf + " + Chr(34), "%97%")
    
    Enchorizar = T
    
End Function

Private Sub cmdGenerar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "Genera un documento que contiene todas las cadenas de texto que deben ser traducidas. Se debe ejecutar esta funcion al terminar de usar el traductor para asegurarse de que se incluyan en el documento todas las cadenas que se deban traducir."
End Sub

Private Sub cmdInsertar_Click()
    frmInsertarReferencias.LlenarFormulario vbinstance
End Sub

Private Sub cmdInsertar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "Inserta las referencias necesarias y la clase traductor, haciendo mas facil el uso del complemento."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = ""
End Sub

Private Sub OKButton_Click()
    'MsgBox "Operación de complemento en: " & VBInstance.FullName
    frmPrincipal.LLenarTodo vbinstance
End Sub

Private Sub OKButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "Utilice el asistente para generar y corregir el codigo para traducir un proyecto."
End Sub

Public Sub EscribirArchivo(path As String, cadena As String)
   ' On Error GoTo e
    Dim fso As FileSystemObject
    Dim f, ts
    Dim S As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
       fso.CreateTextFile path
    End If
    
    Set f = fso.GetFile(path)
    Set ts = f.OpenAsTextStream(2, 0)
     
    ts.Write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub
