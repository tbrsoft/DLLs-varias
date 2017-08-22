VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List2 
      Height          =   6300
      Left            =   5040
      TabIndex        =   2
      Top             =   390
      Width           =   6795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5820
      TabIndex        =   1
      Top             =   0
      Width           =   1035
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   4905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error GoTo MUestraERR
    
    Dim ObjSet As SWbemObjectSet
    Dim SERV As SWbemServices
    
    Set SERV = GetObject("WinMgmts:")
    
    'datos de los discos
    List1.AddItem "INFORMACION DE DISCOS"
    Set ObjSet = SERV.InstancesOf("Win32_LogicalDisk")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each Drive In ObjSet
            If Not IsNull(Drive.Size) And Left(Drive.DeviceID, 1) <> "A" Then
               List1.AddItem ("Drive " & Drive.DeviceID & " contains " & Drive.Size & " bytes")
               List1.AddItem ("VolueSerialNumbre: " + CStr(Drive.volumeserialnumber))
            Else
                List1.AddItem ("Drive " & Drive.DeviceID & " is not available.")
                List1.AddItem ("VolueSerialNumbre: " + CStr(Drive.volumeserialnumber))
            End If
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    
    'ME SALTEO
    'Win32_Fan  Represents the properties of a fan device in the computer system.
    'Win32_HeatPipe  Represents the properties of a heat pipe cooling device.
    'Win32_Refrigeration  Represents the properties of a refrigeration device.
    'Win32_TemperatureProbe  Represents the properties of a temperature sensor (electronic thermometer).


    Set ObjSet = Nothing
    'datos del ventilador
    List1.AddItem "-----------------------"
    List1.AddItem "INFORMACION DEL TECLADO"
    Set ObjSet = SERV.InstancesOf("Win32_KeyBoard")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each tec In ObjSet
            List1.AddItem "Caption: " + CStr(tec.Caption)
            List1.AddItem "Availability: " + CStr(tec.Availability)
            List1.AddItem "CreationClassName: " + CStr(tec.CreationClassName)
            List1.AddItem "Description: " + CStr(tec.Description)
            List1.AddItem "DeviceID: " + CStr(tec.DeviceID)
            List1.AddItem "InstallDate: " + CStr(tec.InstallDate)
            List1.AddItem "IsLocked: " + CStr(tec.IsLocked)
            List1.AddItem "Layout: " + CStr(tec.Layout)
            List1.AddItem "Name: " + CStr(tec.Name)
            List1.AddItem "NumberOfFunctionKeys: " + CStr(tec.NumberOfFunctionKeys)
            List1.AddItem "Password: " + CStr(tec.Password)
            List1.AddItem "PNPDeviceID: " + CStr(tec.PNPDeviceID)
            List1.AddItem "Status: " + CStr(tec.Status)
            List1.AddItem "StatusInfo: " + CStr(tec.StatusInfo)
            List1.AddItem "SystemCreationClassName: " + CStr(tec.SystemCreationClassName)
            List1.AddItem "SystemName: " + CStr(tec.SystemName)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    'SALTEADO
    'Win32_PointingDevice Represents an input device used to point to and select regions on the display of a Win32 computer system.
    'Win32_CDROMDrive Represents a CD-ROM drive on a Win32 computer system.
    'Win32_DiskDrive Represents a physical disk drive as seen by a computer running the Win32 operating system.
    'Win32_FloppyDrive  Manages the capabilities of a floppy disk drive.
    'Win32_LogicalDisk Represents a data source that resolves to an actual local storage device on a Win32 system.
    'Win32_TapeDrive Represents a tape drive on a Win32 computer.
    'Win32_1394Controller  Represents the capabilities and management of a 1394 controller.
    'Win32_1394ControllerDevice  Association class that relates the high-speed serial bus (IEEE 1394 Firewire) Controller and the CIM_LogicalDevice instance connected to it.
    'Win32_AllocatedResource Association class that relates a logical device to a system resource.
    'Win32_AssociatedProcessorMemory  Association class that relates a processor and its cache memory.
    'Win32_BaseBoard Represents a base board (also known as a motherboard or system board).
    
    
    Set ObjSet = Nothing
    'datos del ventilador
    List1.AddItem "-----------------------"
    List1.AddItem "INFORMACION DE LA BIOS"
    Set ObjSet = SERV.InstancesOf("Win32_Bios")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each BIOS In ObjSet
            List1.AddItem "BiosCharacteristics: " + CStr(BIOS.BiosCharacteristics)
            List1.AddItem "BuildNumber: " + CStr(BIOS.BuildNumber)
            List1.AddItem "Caption: " + CStr(BIOS.Caption)
            List1.AddItem "CodeSet: " + CStr(BIOS.CodeSet)
            List1.AddItem "CurrentLanguage: " + CStr(BIOS.CurrentLanguage)
            List1.AddItem "Description: " + CStr(BIOS.Description)
            List1.AddItem "IdentificationCode: " + CStr(BIOS.IdentificationCode)
            List1.AddItem "InstallableLanguages: " + CStr(BIOS.InstallableLanguages)
            List1.AddItem "InstallDate: " + CStr(BIOS.InstallDate)
            List1.AddItem "LanguageEdition: " + CStr(BIOS.LanguageEdition)
            List1.AddItem "ListOfLanguages: " + CStr(BIOS.ListOfLanguages)
            List1.AddItem "Manufacturer: " + CStr(BIOS.Manufacturer)
            List1.AddItem "Name: " + CStr(BIOS.Name)
            List1.AddItem "OtherTargetOS: " + CStr(BIOS.OtherTargetOS)
            List1.AddItem "PrimaryBIOS: " + CStr(BIOS.PrimaryBIOS)
            List1.AddItem "ReleaseDate: " + CStr(BIOS.ReleaseDate)
            List1.AddItem "SerialNumber: " + CStr(BIOS.SerialNumber)
            List1.AddItem "BuildNumber: " + CStr(BIOS.BuildNumber)
            List1.AddItem "SMBIOSBIOSVersion: " + CStr(BIOS.SMBIOSBIOSVersion)
            List1.AddItem "SMBIOSMajorVersion: " + CStr(BIOS.SMBIOSMajorVersion)
            List1.AddItem "SMBIOSMinorVersion: " + CStr(BIOS.SMBIOSMinorVersion)
            List1.AddItem "SMBIOSPresent: " + CStr(BIOS.SMBIOSPresent)
            List1.AddItem "SoftwareElementID: " + CStr(BIOS.SoftwareElementID)
            List1.AddItem "SoftwareElementState: " + CStr(BIOS.SoftwareElementState)
            List1.AddItem "Status: " + CStr(BIOS.Status)
            List1.AddItem "TargetOperatingSystem: " + CStr(BIOS.TargetOperatingSystem)
            List1.AddItem "Version: " + CStr(BIOS.Version)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    
    Set ObjSet = Nothing
    'Win32_Processor Represents a device capable of interpreting a sequence of machine instructions on a Win32 computer system.
    'datos del ventilador
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION DEL PROCESADOR"
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each MICRO In ObjSet
            List1.AddItem "Availability: " + CStr(MICRO.Availability)
            List1.AddItem "AddressWidth: " + CStr(MICRO.AddressWidth)
            List1.AddItem "Architecture: " + CStr(MICRO.Architecture)
            List1.AddItem "CpuStatus: " + CStr(MICRO.CpuStatus)
            List1.AddItem "CreationClassName: " + CStr(MICRO.CreationClassName)
            List1.AddItem "CurrentClockSpeed: " + CStr(MICRO.CurrentClockSpeed)
            List1.AddItem "CurrentVoltage: " + CStr(MICRO.CurrentVoltage)
            List1.AddItem "DataWidth: " + CStr(MICRO.DataWidth)
            List1.AddItem "Description: " + CStr(MICRO.Description)
            List1.AddItem "DeviceID: " + CStr(MICRO.DeviceID)
            List1.AddItem "ExtClock: " + CStr(MICRO.ExtClock)
            List1.AddItem "Family: " + CStr(MICRO.Family)
            List1.AddItem "L2CacheSize: " + CStr(MICRO.L2CacheSize)
            List1.AddItem "L2CacheSpeed: " + CStr(MICRO.L2CacheSpeed)
            List1.AddItem "Level: " + CStr(MICRO.Level)
            List1.AddItem "LoadPercentage: " + CStr(MICRO.LoadPercentage)
            List1.AddItem "Manufacturer: " + CStr(MICRO.Manufacturer)
            List1.AddItem "MaxClockSpeed: " + CStr(MICRO.MaxClockSpeed)
            List1.AddItem "Name: " + CStr(MICRO.Name)
            List1.AddItem "OtherFamilyDescription: " + CStr(MICRO.OtherFamilyDescription)
            List1.AddItem "PNPDeviceID: " + CStr(MICRO.PNPDeviceID)
            List1.AddItem "ProcessorId: " + CStr(MICRO.ProcessorId)
            List1.AddItem "ProcessorType: " + CStr(MICRO.ProcessorType)
            List1.AddItem "Revision: " + CStr(MICRO.Revision)
            List1.AddItem "Role: " + CStr(MICRO.Role)
            List1.AddItem "SocketDesignation: " + CStr(MICRO.SocketDesignation)
            List1.AddItem "Status: " + CStr(MICRO.Status)
            List1.AddItem "StatusInfo: " + CStr(MICRO.StatusInfo)
            List1.AddItem "Stepping: " + CStr(MICRO.Stepping)
            List1.AddItem "SystemCreationClassName: " + CStr(MICRO.SystemCreationClassName)
            List1.AddItem "SystemName: " + CStr(MICRO.SystemName)
            List1.AddItem "UniqueId: " + CStr(MICRO.UniqueId)
            List1.AddItem "UpgradeMethod: " + CStr(MICRO.UpgradeMethod)
            List1.AddItem "Version: " + CStr(MICRO.Version)
            List1.AddItem "VoltageCaps: " + CStr(MICRO.VoltageCaps)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_OnBoardDevice  Represents common adapter devices built into the motherboard (system board).
    'datos de DISPOSITIVOS ON BOARD
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION ONBOARD"
    Set ObjSet = SERV.InstancesOf("Win32_OnBoardDevice")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each ONBOARD In ObjSet
            List1.AddItem "Caption: " + CStr(ONBOARD.Caption)
            List1.AddItem "CreationClassName: " + CStr(ONBOARD.CreationClassName)
            List1.AddItem "Description: " + CStr(ONBOARD.Description)
            List1.AddItem "DeviceType: " + CStr(ONBOARD.DeviceType)
            List1.AddItem "Enabled: " + CStr(ONBOARD.Enabled)
            List1.AddItem "HotSwappable: " + CStr(ONBOARD.HotSwappable)
            List1.AddItem "InstallDate: " + CStr(ONBOARD.InstallDate)
            List1.AddItem "Manufacturer: " + CStr(ONBOARD.Manufacturer)
            List1.AddItem "Model: " + CStr(ONBOARD.Model)
            List1.AddItem "Name: " + CStr(ONBOARD.Name)
            List1.AddItem "OtherIdentifyingInfo: " + CStr(ONBOARD.OtherIdentifyingInfo)
            List1.AddItem "PartNumber: " + CStr(ONBOARD.PartNumber)
            List1.AddItem "PoweredOn: " + CStr(ONBOARD.PoweredOn)
            List1.AddItem "Removable: " + CStr(ONBOARD.Removable)
            List1.AddItem "Replaceable: " + CStr(ONBOARD.Replaceable)
            List1.AddItem "SerialNumber: " + CStr(ONBOARD.SerialNumber)
            List1.AddItem "SKU: " + CStr(ONBOARD.SKU)
            List1.AddItem "Status: " + CStr(ONBOARD.Status)
            List1.AddItem "Tag: " + CStr(ONBOARD.Tag)
            List1.AddItem "Version: " + CStr(ONBOARD.Version)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_MotherboardDevice Represents a device that contains the central components of the Win32 computer system.
    'datos de DISPOSITIVOS ON BOARD
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION MOTHR DEVICE"
    Set ObjSet = SERV.InstancesOf("Win32_MotherboardDevice")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each ONBOARD In ObjSet
            List1.AddItem "Availability: " + CStr(ONBOARD.Availability)
            List1.AddItem "Caption: " + CStr(ONBOARD.Caption)
            List1.AddItem "Description: " + CStr(ONBOARD.Description)
            List1.AddItem "DeviceID: " + CStr(ONBOARD.DeviceID)
            List1.AddItem "InstallDate: " + CStr(ONBOARD.InstallDate)
            List1.AddItem "Name: " + CStr(ONBOARD.Name)
            List1.AddItem "PNPDeviceID: " + CStr(ONBOARD.PNPDeviceID)
            List1.AddItem "PowerManagementCapabilities: " + CStr(ONBOARD.PowerManagementCapabilities)
            List1.AddItem "PowerManagementSupported: " + CStr(ONBOARD.PowerManagementSupported)
            List1.AddItem "PrimaryBusType: " + CStr(ONBOARD.PrimaryBusType)
            List1.AddItem "RevisionNumber: " + CStr(ONBOARD.RevisionNumber)
            List1.AddItem "SecondaryBusType: " + CStr(ONBOARD.SecondaryBusType)
            List1.AddItem "Status: " + CStr(ONBOARD.Status)
            List1.AddItem "StatusInfo: " + CStr(ONBOARD.StatusInfo)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_PortConnector Represents physical connection ports, such as DB-25 pin male, Centronics, and PS/2.
    'datos de PUERTOS CONECTADOS
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION PUERTOS CONECTADOS"
    Set ObjSet = SERV.InstancesOf("Win32_PortConnector")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each PORTS In ObjSet
            List1.AddItem "Caption: " + CStr(PORTS.Caption)
            List1.AddItem "ConnectorPinout: " + CStr(PORTS.ConnectorPinout)
            List1.AddItem "ConnectorType: " + CStr(PORTS.ConnectorType)
            List1.AddItem "Description: " + CStr(PORTS.Description)
            List1.AddItem "ExternalReferenceDesignator: " + CStr(PORTS.ExternalReferenceDesignator)
            List1.AddItem "InstallDate: " + CStr(PORTS.InstallDate)
            List1.AddItem "InternalReferenceDesignator: " + CStr(PORTS.InternalReferenceDesignator)
            List1.AddItem "Manufacturer: " + CStr(PORTS.Manufacturer)
            List1.AddItem "Model: " + CStr(PORTS.Model)
            List1.AddItem "Name: " + CStr(PORTS.Name)
            List1.AddItem "OtherIdentifyingInfo: " + CStr(PORTS.OtherIdentifyingInfo)
            List1.AddItem "PartNumber: " + CStr(PORTS.PartNumber)
            List1.AddItem "PortType: " + CStr(PORTS.PortType)
            List1.AddItem "PoweredOn: " + CStr(PORTS.PoweredOn)
            List1.AddItem "SerialNumber: " + CStr(PORTS.SerialNumber)
            List1.AddItem "SKU: " + CStr(PORTS.SKU)
            List1.AddItem "Status: " + CStr(PORTS.Status)
            List1.AddItem "Tag: " + CStr(PORTS.Tag)
            List1.AddItem "Version: " + CStr(PORTS.Version)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_SystemSlot Represents physical connection points including ports, motherboard slots and peripherals, and proprietary connections points.
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION SLOTS"
    
    Set ObjSet = SERV.InstancesOf("Win32_SystemSlot")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
    
        For Each slots In ObjSet
            List1.AddItem "Caption: " + CStr(slots.Caption)
            List1.AddItem "ConnectorPinout: " + CStr(slots.ConnectorPinout)
            List1.AddItem "ConnectorType: " + CStr(slots.ConnectorType)
            List1.AddItem "CurrentUsage: " + CStr(slots.CurrentUsage)
            List1.AddItem "Description: " + CStr(slots.Description)
            List1.AddItem "HeightAllowed: " + CStr(slots.HeightAllowed)
            List1.AddItem "InstallDate: " + CStr(slots.InstallDate)
            List1.AddItem "LengthAllowed: " + CStr(slots.LengthAllowed)
            List1.AddItem "Manufacturer: " + CStr(slots.Manufacturer)
            List1.AddItem "MaxDataWidth: " + CStr(slots.MaxDataWidth)
            List1.AddItem "Model: " + CStr(slots.Model)
            List1.AddItem "Name: " + CStr(slots.Name)
            List1.AddItem "Number: " + CStr(slots.Number)
            List1.AddItem "OtherIdentifyingInfo: " + CStr(slots.OtherIdentifyingInfo)
            List1.AddItem "PartNumber: " + CStr(slots.PartNumber)
            List1.AddItem "PMESignal: " + CStr(slots.PMESignal)
            List1.AddItem "PoweredOn: " + CStr(slots.PoweredOn)
            List1.AddItem "PurposeDescription: " + CStr(slots.PurposeDescription)
            List1.AddItem "SerialNumber: " + CStr(slots.SerialNumber)
            List1.AddItem "Shared: " + CStr(slots.Shared)
            List1.AddItem "SKU: " + CStr(slots.SKU)
            List1.AddItem "SlotDesignation: " + CStr(slots.SlotDesignation)
            List1.AddItem "SpecialPurpose: " + CStr(slots.SpecialPurpose)
            List1.AddItem "Status: " + CStr(slots.Status)
            List1.AddItem "SupportsHotPlug: " + CStr(slots.SupportsHotPlug)
            List1.AddItem "Tag: " + CStr(slots.Tag)
            List1.AddItem "ThermalRating: " + CStr(slots.ThermalRating)
            List1.AddItem "VccMixedVoltageSupport: " + CStr(slots.VccMixedVoltageSupport)
            List1.AddItem "Version: " + CStr(slots.Version)
            List1.AddItem "VppMixedVoltageSupport: " + CStr(slots.VppMixedVoltageSupport)
        
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    List1.Height = Me.Height - 1200
    List2.Height = Me.Height - 1200
    WriteLSTinFIle List1
    Exit Sub
MUestraERR:
    List2.AddItem CStr(Err.Number) + "-" + Err.Description
    Resume Next
End Sub

Public Sub WriteLSTinFIle(LST As ListBox)
    libre = FreeFile
    Open "c:\lst.txt" For Output As libre
        c = 0
        Do While c < LST.ListCount
            Write #libre, LST.List(c)
            c = c + 1
        Loop
    Close #libre
End Sub

