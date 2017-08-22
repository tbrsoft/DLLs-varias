VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Usar registro avanzado de errores"
      Height          =   285
      Left            =   690
      TabIndex        =   1
      Top             =   330
      Width           =   2925
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAREA1"
      Height          =   465
      Left            =   1140
      TabIndex        =   0
      Top             =   690
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'el modulo de errores una sencilla herramienta para registrar
'detalles del momento en tiempo de ejecucion

'primero se debe agregar una referencia a tbrSoft Manejador de errores (tbrerr.dll)

Dim TERR As New tbrErrores.clsTbrERR

'tiene 2 formas de trabajar.
'La forma normal solo escribe registros al disco en caso de que nostros _
    hayamos interceptado un error y deseemos dejarlo registrado

'La segunda forma consume más recursos y escribe todo el tiempo todo lo _
    que hace el programa. Esto solo se activa cuando el programa se cuelga _
    en errores no interceptados
    
'el metodo mas usado es "ANOTAR" que permite indicar un codigo de referencia y
'4 parametros mas opcionales de cualquier tipo (se pasan a string en la Dll)
'estos parametros se usan para indicar valores de variables
'este solo suma un renglon al registro interno que solo se escribe

'el metodo appendlog agrega un renglon al archivo de log ademas de los xxx caracteres
'anteriores (para tener un contexto de como paso el error)
'existe tambien AppendLogSinHist para que solo escriba un renglon sin contexto
'solo lo uso para error previsibles que no necesitan contexto. Casi no lo uso


Dim ActivarErr As Boolean 'indica si se activa el modo especial de registro

Private Sub Form_Load()
    
    TERR.FileLog = App.Path + "\log.txt" 'lugar donde se graba el registro normal de errores
    TERR.LargoAcumula = 400 'cantidad de caracteres que guarda, son los ultimos _
        y los escribe todos al momento de que se lo solicite
    
    On Local Error GoTo TRE
    
    'aqui registro una ubicacion unica en el programa (aaaa) y puedo agregar _
        hasta 4 variables de cualquier tipo que soporte cstr(var) y que sirva
    TERR.Anotar "aaaa", ActivarErr
    
    Exit Sub
    
TRE:
    TERR.AppendLog "aaac", TERR.ErrToTXT(Err)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'si sale normalmente no queda el archivo grabado ya que no sera necesario
    If ActivarErr Then
        TERR.StopGrabaTodo 'cierra y borra el archivo ya que salimos OK del sistema
    End If
End Sub

Private Sub Command1_Click()
    ActivarErr = Check1
    
    'en caso de que se active el registro permanente ....-
    If ActivarErr Then
        'esto cada vez que usamos ANOTAR se escribe en disco. O sea que hace trabajar mucho _
            al disco. Usarse solo en caso de que no sepan que hacer.
            'Yo lo tengo configurable en el 3PM pero predeterminado no se usa
        
        Dim n As String
        n = CStr(Day(Date)) + "." + CStr(Month(Date)) + "." + CStr(Year(Date)) + _
            "." + CStr(Hour(Time)) + "." + CStr(Minute(Time)) + "." + CStr(Second(Time))
        TERR.FileLogGrabaTodo = AP + "REG" + CStr(n) + ".W15"
        '....quedara un registro con la fecha y hora del error con la extencion W15
        TERR.ModoGrabaTodo = True
        TERR.StartGrabaTodo 'esto hace que cada anotar escriba a disco
        'de esta forma si el programa se clava y no sale por el unload normal
        'el archivo quedara escrito (se elimina al salir normalmente)
    End If
    
    
    On Local Error GoTo CM1
    'por ejemplo leer los archivos que hay en esta carpeta
    'como tarea de ejemplo que esta sujeta a errores
    
    Dim SF As String
    SF = Dir(App.Path + "\*.*")
    
    Dim FS As New Scripting.FileSystemObject
    Dim TE As TextStream, Testo As String
    
    Do While SF <> ""
        TERR.Anotar "aaab", SF, FileLen(App.Path + "\" + SF) 'puedo registra el paso con el valor único (para _
            encontrarlo despues). Ademas indico por el archivo que paso y para _
            mostrar que va cualquier tipo de valor el tamaño que tiene
            
        Set TE = FS.OpenTextFile(App.Path + "\" + SF, ForReading)
            Testo = TE.ReadAll
            'supongamos que los archivos deben decir "hola" en los primeros 4 bytes
            If Left(Testo, 4) <> "hola" Then
                'aca hice uno mas largo
                'la ubicacion unica aaad
                'una segunda que explica (esto podría ir comentado en el codigo
                'para esconderle al usuario que paso)
                TERR.AppendLog "aaae", "No dice hola, dice:" + Left(Testo, 4)
                'se pueden hacer las combinaciones que se quiera
            End If
        TE.Close
        SF = Dir
    Loop
    
    Exit Sub
    
CM1:
    TERR.AppendLog "aaad", TERR.ErrToTXT(Err)
    'aqui hay una desicion del programador. ¿sigo con la linea que sigue?
    Resume Next
End Sub
