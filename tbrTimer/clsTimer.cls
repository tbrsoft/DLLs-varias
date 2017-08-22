VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

'Version 1.00, 15.01.03
'(c) by Goetz Reinecke 01/2003
'    reinecke@activevb.de

Option Explicit

Public Event Timer()

Private m_Interval As Long
Private m_Enabled As Long
Private m_hTimer As Long
Private m_Added As Boolean

Private Sub Class_Terminate()
    Call StopTimer
End Sub

Public Sub TimerEvent()
    RaiseEvent Timer
End Sub

Public Property Let Enabled(n_Enabled As Boolean)
    m_Enabled = n_Enabled
    If m_Enabled Then
        If m_Interval > 0 Then 'agregado feb 2010
            Call StartTimer
        End If
    Else
        Call StopTimer
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Interval(n_Interval As Long)
    m_Interval = n_Interval
    'lo puse yop para que no empieze si es cero!!!
    If m_Interval > 0 Then
        If m_Enabled Then Call StartTimer
    Else
        StopTimer 'agregado por mi
    End If
End Property

Public Property Get Interval() As Long
    Interval = m_Interval
End Property

Private Sub StartTimer()
    If m_Interval > 0 Then
        Call StopTimer
        m_hTimer = modTimer.AddObject(Me, m_Interval)
        m_Added = True
    End If
End Sub

Private Sub StopTimer()
    If m_Added Then Call modTimer.RemoveObject(m_hTimer)
    m_Added = False
End Sub
