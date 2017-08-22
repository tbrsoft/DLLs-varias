Attribute VB_Name = "AVICAP"
'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------

Public lwndC As Long       ' Handle to the Capture Windows
Public Const WM_USER = &H400

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
     ByVal lParam As Long) As Long

'// Defines start of the message range

Public Const WM_CAP_START = WM_USER
Public Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14

'// Following added post VFW 1.1

'// ------------------------------------------------------------------
'//  Structures
'// ------------------------------------------------------------------
Type CAPDRIVERCAPS
    wDeviceIndex As Long                   '// Driver index in system.ini
    fHasOverlay As Long                    '// Can device overlay?
    fHasDlgVideoSource As Long             '// Has Video source dlg?
    fHasDlgVideoFormat As Long             '// Has Format dlg?
    fHasDlgVideoDisplay As Long            '// Has External out dlg?
    fCaptureInitialized As Long            '// Driver ready to capture?
    fDriverSuppliesPalettes As Long        '// Can driver make palettes?
    hVideoIn As Long                       '// Driver In channel
    hVideoOut As Long                      '// Driver Out channel
    hVideoExtIn As Long                    '// Driver Ext In channel
    hVideoExtOut As Long                   '// Driver Ext Out channel
End Type


'// The two functions exported by AVICap
Declare Function capGetDriverDescriptionA Lib "avicap32.dll" ( _
    ByVal wDriver As Integer, ByVal lpszName As String, ByVal cbName As Long, _
    ByVal lpszVer As String, ByVal cbVer As Long) As Boolean



Function capDriverGetCaps(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean

   capDriverGetCaps = SendMessage(lwnd, WM_CAP_DRIVER_GET_CAPS, wSize, s)

End Function





