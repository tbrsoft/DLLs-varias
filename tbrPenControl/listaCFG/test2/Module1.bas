Attribute VB_Name = "Module1"
Public Ap As String
Public FSO As New Scripting.FileSystemObject


Public Sub Main()
    On Local Error GoTo ERRMAIN
    
    Load frmTest 'necesito que este el ccfg
    
    Ap = App.Path
    If Right(Ap, 1) <> "\" Then Ap = Ap + "\"
    
    Dim CM As String
    CM = Command 'parametro (solo uno, un numero al azar que se verifica)
    Dim CM2 As String 'parametro que supuestamente se deberia haber mandado
    
    'hay un archivo supertemporal que me da el programa que me llama que dice las cosas basicas que tengo que hacer
    'SEGUIRAQUI
    Dim SuperTemp As String
    SuperTemp = Ap + "now.ifo"
    
    If FSO.FileExists(SuperTemp) = False Then
        Exit Sub 'nunca se hara visible
    End If
    
    Dim TX As TextStream, R As String
    Set TX = FSO.OpenTextFile(SuperTemp, ForReading, False)
        Dim cfg As String, val As String, pos2P As Long
        Do While Not TX.AtEndOfStream
            R = TX.ReadLine
            pos2P = InStr(R, ":")
            If pos2P > 0 Then
                cfg = Mid(R, 1, pos2P - 1)
                val = Mid(R, pos2P + 1)
                
                Select Case LCase(cfg)
                    Case "sv" ' CV.Juamai 'el software que lo llama (para diferenciar si es martino o mprock o e2games o 3pm!!!)
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.SoftNow = val
                        
                        
                    Case "qii" ' CV.Qii 'la clave del software que lo esta llamando
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.SetPcKy val
                        
                        
                    Case "ex" 'path al archivo de la lista con nodos y ejecuciones
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.Load val
                        
                        
                    Case "pthlog" 'PARA LOG del pendrive se grabe aqui!
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.setPathError val
                        
                        
                    Case "orig" ' PartOrig(0)
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.AddOrigMusica val 'CASI SIEMPRE ES MAS DE UNO
                        
                        
                    Case "pthmusic"
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.SetPathMusic val  'CASI SIEMPRE ES MAS DE UNO
                    
                    
                    Case "ptsysrockola"
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        frmTest.CCFG.SetPathSysRockola val
                        
                    Case "perm" 'debe estar antes del load para que no se carguen los nodos que no deben cargarse!!!
                        frmTest.CCFG.wError R, False 'es solo ANOTAR
                        Dim spK() As String
                        spK = Split(val, "|")
                        Dim kj As Long
                        For kj = 0 To UBound(spK)
                            Dim spk2() As String
                            spk2 = Split(spK(kj), ":")
                            
                            frmTest.CCFG.GetPerms.AddPerm CLng(spk2(0)), CLng(spk2(1))
                            
                        Next kj
                    
                    Case "az"
                        CM2 = val
                    Case "varaibol"
                        Dim spyu() As String
                        spyu = Split(val)
                        frmTest.CCFG.VS.SetV spyu(0), spyu(1)
                End Select
            End If
        Loop
    TX.Close
    
    If FSO.FileExists(SuperTemp) Then
        'borrar el archivo temporal para no dejar rastros
        FSO.DeleteFile SuperTemp, True 'que no queden registros, recien se creo en el mprock (como si fuera un parametro gigante)
    End If
    
    'si no es de un archivo "now" para ejecutar ya no se ejecuta
    If Trim(CM) <> Trim(CM2) Then 'para que cada uso sea unico este ejecutable se abre pasando un parametro que debe existir en el archivo temporal "now". Aqui se verifica que sean iguales
        frmTest.CCFG.wError "No CM-CM2", True
        Exit Sub
    End If
    
    'leer la info que me deben dar para que me abra!
    frmTest.Visible = True
    frmTest.Show
    
    Exit Sub
    
ERRMAIN:
    Dim S33 As String
    S33 = CStr(Err.Number) + "= " + Err.Description
    frmTest.CCFG.wError S33, True
    MsgBox "Error al iniciar complemento" + vbCrLf + S33
    Unload frmTest
    End
End Sub

Public Sub APAGAR_PC()
    Dim VW As New clsWindowsVERSION
    Dim v As vWindows
    v = VW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        Shell "rundll32 user.exe,exitwindows"
    Case Win2000, WinNT4, WinXp, WinXP2
        Shell "Shutdown -s -t 0"
    End Select
End Sub

Public Sub REINICIAR_PC()
    Dim VW As New clsWindowsVERSION
    Dim v As vWindows
    v = VW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        Shell "rundll32 user.exe,exitwindowsexec"
    Case Win2000, WinNT4, WinXp, WinVista
        Shell "Shutdown -r -t 0" 'el -s es shutdowsn y el -r restart
    End Select
End Sub

