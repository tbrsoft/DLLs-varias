VERSION 5.00
Begin VB.Form frmKeyboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmKeyboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstKeyboard 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Get List"
      Height          =   350
      Left            =   7800
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.ListBox lstPowerManagementCapabilities 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4215
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   8655
   End
   Begin VB.Label lblList 
      Caption         =   "List"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblPowerManagementCapabilities 
      Caption         =   "Power Management Capabilities"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Keyboard As SWbemObject
   
    'Clear current
    lstKeyboard.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim KeyboardSet As SWbemObjectSet
    Set KeyboardSet = Namespace.InstancesOf("Win32_Keyboard")
    
    For Each Keyboard In KeyboardSet
        ' Use the RelPath property of the instance path to display the disk
        lstKeyboard.AddItem Keyboard.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstKeyboard_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Keyboard As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstPowerManagementCapabilities.Clear

    Me.MousePointer = vbHourglass
    
    SelectedItem = lstKeyboard.List(lstKeyboard.ListIndex)
    Set Keyboard = Namespace.Get(SelectedItem)
    
    Value = Keyboard.Availability
    lstData.AddItem Left("Availability" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(0) = lstData.List(0) & "Other"
        Case 2: lstData.List(0) = lstData.List(0) & "Unknown"
        Case 3: lstData.List(0) = lstData.List(0) & "Running/Full Power"
        Case 4: lstData.List(0) = lstData.List(0) & "Warning"
        Case 5: lstData.List(0) = lstData.List(0) & "In Test"
        Case 6: lstData.List(0) = lstData.List(0) & "Not Applicable"
        Case 7: lstData.List(0) = lstData.List(0) & "Power Off"
        Case 8: lstData.List(0) = lstData.List(0) & "Off Line"
        Case 9: lstData.List(0) = lstData.List(0) & "Off Duty"
        Case 10: lstData.List(0) = lstData.List(0) & "Degraded"
        Case 11: lstData.List(0) = lstData.List(0) & "Not Installed"
        Case 12: lstData.List(0) = lstData.List(0) & "Install Error"
        Case 13: lstData.List(0) = lstData.List(0) & "Power Save - Unknown"
        Case 14: lstData.List(0) = lstData.List(0) & "Power Save - Low Power Mode"
        Case 15: lstData.List(0) = lstData.List(0) & "Power Save - Standby"
        Case 16: lstData.List(0) = lstData.List(0) & "Power Cycle"
        Case 17: lstData.List(0) = lstData.List(0) & "Power Save - Warning"
    End Select
    
    Value = Keyboard.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = Keyboard.ConfigManagerErrorCode
    lstData.AddItem Left("Config Manager Error Code" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(2) = lstData.List(2) & "This device is working properly."
        Case 1: lstData.List(2) = lstData.List(2) & "This device is not configured correctly."
        Case 2: lstData.List(2) = lstData.List(2) & "Windows cannot load the driver for this device."
        Case 3: lstData.List(2) = lstData.List(2) & "The driver for this device might be corrupted, or your system may be running low on memory or other resources."
        Case 4: lstData.List(2) = lstData.List(2) & "This device is not working properly. One of its drivers or your registry might be corrupted."
        Case 5: lstData.List(2) = lstData.List(2) & "The driver for this device needs a resource that Windows cannot manage."
        Case 6: lstData.List(2) = lstData.List(2) & "The boot configuration for this device conflicts with other devices."
        Case 7: lstData.List(2) = lstData.List(2) & "Cannot filter."
        Case 8: lstData.List(2) = lstData.List(2) & "The driver loader for the device is missing."
        Case 9: lstData.List(2) = lstData.List(2) & "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."
        Case 10: lstData.List(2) = lstData.List(2) & "This device cannot start."
        Case 11: lstData.List(2) = lstData.List(2) & "This device failed."
        Case 12: lstData.List(2) = lstData.List(2) & "This device cannot find enough free resources that it can use."
        Case 13: lstData.List(2) = lstData.List(2) & "Windows cannot verify this device's resources."
        Case 14: lstData.List(2) = lstData.List(2) & "This device cannot work properly until you restart your computer."
        Case 15: lstData.List(2) = lstData.List(2) & "This device is not working properly because there is probably a re-enumeration problem. "
        Case 16: lstData.List(2) = lstData.List(2) & "Windows cannot identify all the resources this device uses."
        Case 17: lstData.List(2) = lstData.List(2) & "This device is asking for an unknown resource type."
        Case 18: lstData.List(2) = lstData.List(2) & "Reinstall the drivers for this device."
        Case 19: lstData.List(2) = lstData.List(2) & "Your registry might be corrupted."
        Case 20: lstData.List(2) = lstData.List(2) & "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation."
        Case 21: lstData.List(2) = lstData.List(2) & "Windows is removing this device."
        Case 22: lstData.List(2) = lstData.List(2) & "This device is disabled."
        Case 23: lstData.List(2) = lstData.List(2) & "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."
        Case 24: lstData.List(2) = lstData.List(2) & "This device is not present, is not working properly, or does not have all its drivers installed."
        Case 25: lstData.List(2) = lstData.List(2) & "Windows is still setting up this device."
        Case 26: lstData.List(2) = lstData.List(2) & "Windows is still setting up this device."
        Case 27: lstData.List(2) = lstData.List(2) & "This device does not have valid log configuration."
        Case 28: lstData.List(2) = lstData.List(2) & "The drivers for this device are not installed."
        Case 29: lstData.List(2) = lstData.List(2) & "This device is disabled because the firmware of the device did not give it the required resources."
        Case 30: lstData.List(2) = lstData.List(2) & "This device is using an Interrupt Request (IRQ) resource that another device is using."
        Case 31: lstData.List(2) = lstData.List(2) & "This device is not working properly because Windows cannot load the drivers required for this device."
    End Select

    Value = Keyboard.ConfigManagerUserConfig
    lstData.AddItem Left("Config Manager User Config" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = Keyboard.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)

    Value = Keyboard.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = Keyboard.DeviceID
    lstData.AddItem Left("Device ID" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = Keyboard.ErrorCleared
    lstData.AddItem Left("Error Cleared" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = Keyboard.ErrorDescription
    lstData.AddItem Left("Error Description" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = Keyboard.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)

    Value = Keyboard.IsLocked
    lstData.AddItem Left("Is Locked" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = Keyboard.LastErrorCode
    lstData.AddItem Left("Last Error Code" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = Keyboard.Layout
    lstData.AddItem Left("Layout" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = Keyboard.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = Keyboard.NumberOfFunctionKeys
    lstData.AddItem Left("Number Of Function Keys" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = Keyboard.Password
    lstData.AddItem Left("Password" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(15) = lstData.List(15) & "Other"
        Case 2: lstData.List(15) = lstData.List(15) & "Unknown"
        Case 3: lstData.List(15) = lstData.List(15) & "Disabled"
        Case 4: lstData.List(15) = lstData.List(15) & "Enabled"
        Case 5: lstData.List(15) = lstData.List(15) & "Not Implemented"
    End Select

    Value = Keyboard.PNPDeviceID
    lstData.AddItem Left("PNP Device ID" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = Keyboard.PowerManagementCapabilities
    For tmpInt = 0 To 7
        Select Case Value(tmpInt) 'Put in data
            Case 0: lstPowerManagementCapabilities.AddItem "Unknown"
            Case 1: lstPowerManagementCapabilities.AddItem "Not Supported"
            Case 2: lstPowerManagementCapabilities.AddItem "Disabled"
            Case 3: lstPowerManagementCapabilities.AddItem "Enabled"
            Case 4: lstPowerManagementCapabilities.AddItem "Power Saving Modes Entered Automatically"
            Case 5: lstPowerManagementCapabilities.AddItem "Power State Settable"
            Case 6: lstPowerManagementCapabilities.AddItem "Power Cycling Supported"
            Case 7: lstPowerManagementCapabilities.AddItem "Timed Power On Supported"
            Case Else: Exit For
        End Select
    Next tmpInt

    Value = Keyboard.PowerManagementSupported
    lstData.AddItem Left("Power Management Supported" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = Keyboard.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = Keyboard.StatusInfo
    lstData.AddItem Left("StatusInfo" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(19) = lstData.List(19) & "Other"
        Case 2: lstData.List(19) = lstData.List(19) & "Unknown"
        Case 3: lstData.List(19) = lstData.List(19) & "Enabled"
        Case 4: lstData.List(19) = lstData.List(19) & "Disabled"
        Case 5: lstData.List(19) = lstData.List(19) & "Not Applicable"
    End Select

    Value = Keyboard.SystemCreationClassName
    lstData.AddItem Left("System Creation Class Name" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)
    
    Value = Keyboard.SystemName
    lstData.AddItem Left("System Name" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
