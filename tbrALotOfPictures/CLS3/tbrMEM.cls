VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrMEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MEMORYSTATUS
         dwLength As Long
         dwMemoryLoad As Long
         dwTotalPhys As Long
         dwAvailPhys As Long
         dwTotalPageFile As Long
         dwAvailPageFile As Long
         dwTotalVirtual As Long
         dwAvailVirtual As Long
 End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private MemUsadaMomentos() As Long
Private Detalles() As String

'Para cantidad de memoria

Public Sub SetMomento(sDESCR As String) 'dejar un comentario de que esta haciendo el programa
    
    Dim tmpUSA As Long, TOT As Long
    GetMem TOT, tmpUSA
    
    ReDim Preserve MemUsadaMomentos(UBound(MemUsadaMomentos) + 1)
    MemUsadaMomentos(UBound(MemUsadaMomentos)) = tmpUSA
    
    ReDim Preserve Detalles(UBound(Detalles) + 1)
    Detalles(UBound(Detalles)) = CStr(Timer) + " " + sDESCR
End Sub

Public Function GetMemMoment(i As Long)
    If i > GetMaxMoment Then
        GetMemMoment = -1
    Else
        GetMemMoment = MemUsadaMomentos(i)
    End If
End Function

Public Function GetDescrMoment(i As Long)
    If i > GetMaxMoment Then
        GetDescrMoment = -1
    Else
        GetDescrMoment = Detalles(i)
    End If
End Function

Public Function GetFullDetalles() As String
    Dim J As Long, TMP As String
    TMP = ""
    For J = 0 To UBound(Detalles)
        TMP = TMP + CStr(MemUsadaMomentos(J)) + " MB " + Detalles(J) + vbCrLf
    Next J
    GetFullDetalles = TMP
End Function

Public Function GetMaxMoment() As Long
    GetMaxMoment = UBound(MemUsadaMomentos)
End Function

Private Sub Class_Initialize()
    ReDim MemUsadaMomentos(0)
    ReDim Detalles(0)
    
    'dejo asentado como arranco
    Dim tmpUSA As Long, TOT As Long
    GetMem TOT, tmpUSA
    
    MemUsadaMomentos(0) = tmpUSA
    Detalles(0) = CStr(Timer) + " INICIO tbrMEM con " + CStr(TOT) + " MB totales"
End Sub

Public Function GetMem(ByRef MemTotalInMB As Long, ByRef MemUsadaInMB As Long) As Long
    
    On Local Error GoTo ErMEM
    
    Dim MemStat As MEMORYSTATUS
    GlobalMemoryStatus MemStat

    
    MemTotalInMB = MemStat.dwTotalPhys / (1024 ^ 2)
    MemUsadaInMB = MemTotalInMB - (MemStat.dwAvailPhys / (1024 ^ 2))


    GetMem = 0
    Exit Function
    
ErMEM:
    GetMem = 1
End Function

'    List1.AddItem "Total RAM: " + CStr(Int(((MemStat.dwTotalPhys) / 1024 ^ 2) * 100) / 100) + " Mb."
'    List1.AddItem "dwAvailPageFile: " + CStr(MemStat.dwAvailPageFile)
'    List1.AddItem "dwAvailPhys: " + CStr(MemStat.dwAvailPhys)
'    List1.AddItem "dwAvailVirtual: " + CStr(MemStat.dwAvailVirtual)
'    List1.AddItem "dwLength: " + CStr(MemStat.dwLength)
'    List1.AddItem "dwMemoryLoad: " + CStr(MemStat.dwMemoryLoad)
'    List1.AddItem "dwTotalPageFile: " + CStr(MemStat.dwTotalPageFile)
'    List1.AddItem "dwTotalVirtual: " + CStr(MemStat.dwTotalVirtual)

'***************************
'***************************
'******* SDK DICE ********************
'***************************
'***************************
'dwLength
'Size, in bytes, of the MEMORYSTATUS data structure. _
    You do not need to set this member before calling the GlobalMemoryStatus function; the _
    function sets it.
    
'dwMemoryLoad
    'Windows NT 3.1 – NT 4.0: The percentage of approximately the last 1000 pages of _
    physical memory that is in use.
    
    'Windows 2000: The approximate percentage of total physical memory that is in use.

'dwTotalPhys
    'Total size, in bytes, of physical memory.
'dwAvailPhys
    'Size, in bytes, of physical memory available.
'dwTotalPageFile
    'Total possible size, in bytes, of the paging file. Note that this number does not represent the actual physical size of the paging file on disk.
'dwAvailPageFile
    'Size, in bytes, of space available in the paging file. The operating system can enlarge the paging file from time to time. The dwAvailPageFile member shows the difference between the size of current committed memory and the current size of the paging file — it does not show the largest possible size of the paging file.
'dwTotalVirtual
    'Total size, in bytes, of the user mode portion of the virtual address space of the calling process.
'dwAvailVirtual
    'Size, in bytes, of unreserved and uncommitted memory in the user mode portion of the virtual address space of the calling process.

'Remarks
    'MEMORYSTATUS reflects the state of memory at the time of the call. _
    It reflects the size of the paging file at that time. The operating system can enlarge _
    the paging file up to the maximum size set by the administrator.

'On computers with more than 4 GB of memory, the MEMORYSTATUS structure can return incorrect _
    information. Windows 2000 reports a value of -1 to indicate an overflow. Earlier versions _
    of Windows NT report a value that is the real amount of memory, modulo 4 GB. If your _
    application is at risk for this behavior, use the GlobalMemoryStatusEx function instead _
    of the GlobalMemoryStatus function.

'***************************
'***************************
'mi notebook da
'dwTotalPhys:        234 209 280
'dwAvailPageFile:    295 292 928
'dwAvailPhys:         44 843 008
'dwAvailVirtual:   2 059 149 312
'dwLength:                    32
'dwMemoryLoad:                80
'dwTotalPageFile:    570 748 928
'dwTotalVirtual:   2 147 352 576
'***************************
'***************************

