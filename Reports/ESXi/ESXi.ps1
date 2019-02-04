#requires -Modules @{ModuleName="PScribo";ModuleVersion="0.7.24"},VMware.VimAutomation.Core

<#
.SYNOPSIS  
    PowerShell script to document the configuration of VMware vSphere infrastucture in Word/HTML/XML/Text formats
.DESCRIPTION
    Documents the configuration of VMware vSphere infrastucture in Word/HTML/XML/Text formats using PScribo.
.NOTES
    Version:        0.3.0
    Author:         Tim Carman
    Twitter:        @tpcarman
    Github:         tpcarman
    Credits:        Iain Brighton (@iainbrighton) - PScribo module
                    Jake Rutski (@jrutski) - VMware vSphere Documentation Script Concept
.LINK
    https://github.com/tpcarman/As-Built-Report
    https://github.com/iainbrighton/PScribo
#>

#region Configuration Settings
#---------------------------------------------------------------------------------------------#
#                                    CONFIG SETTINGS                                          #
#---------------------------------------------------------------------------------------------#
# Clear variables
$ESXiHost = @()
$VIServer = @()

# If custom style not set, use VMware style
if (!$StyleName) {
    & "$PSScriptRoot\..\..\Styles\VMware.ps1"
}

#endregion Configuration Settings

#region Script Functions
#---------------------------------------------------------------------------------------------#
#                                    SCRIPT FUNCTIONS                                         #
#---------------------------------------------------------------------------------------------#

function Get-License {
    <#
    .SYNOPSIS
    Function to retrieve vSphere product licensing information.
    .DESCRIPTION
    Function to retrieve vSphere product licensing information.
    .NOTES
    Version:        0.1.0
    Author:         Tim Carman
    Twitter:        @tpcarman
    Github:         tpcarman
    .PARAMETER VMHost
    A vSphere ESXi Host object
    .PARAMETER vCenter
    A vSphere vCenter Server object
    .PARAMETER Licenses
    All vSphere product licenses
    .INPUTS
    System.Management.Automation.PSObject.
    .OUTPUTS
    System.Management.Automation.PSObject.
    .EXAMPLE
    PS> Get-License -VMHost ESXi01
    .EXAMPLE
    PS> Get-License -vCenter VCSA
    .EXAMPLE
    PS> Get-License -Licenses
    #>
    [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]

    Param
    (
        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$vCenter, [PSObject]$VMHost,
        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [Switch]$Licenses
    ) 

    $LicenseObject = @()
    $ServiceInstance = Get-View ServiceInstance -Server $ESXiHost
    $LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
    $LicenseManagerAssign = Get-View $LicenseManager.LicenseAssignmentManager 
    if ($VMHost) {
        $VMHostId = $VMHost.Extensiondata.Config.Host.Value
        $VMHostAssignedLicense = $LicenseManagerAssign.QueryAssignedLicenses($VMHostId)    
        $VMHostLicense = $VMHostAssignedLicense | Where-Object {$_.EntityId -eq $VMHostId}
        if ($Options.ShowLicenses) {
            $VMHostLicenseKey = $VMHostLicense.AssignedLicense.LicenseKey
        } else {
            $VMHostLicenseKey = "*****-*****-*****" + $VMHostLicense.AssignedLicense.LicenseKey.Substring(17)
        }
        $LicenseObject = [PSCustomObject]@{                               
            Product = $VMHostLicense.AssignedLicense.Name 
            LicenseKey = $VMHostLicenseKey                   
        }
    }
    if ($vCenter) {
        $vCenterAssignedLicense = $LicenseManagerAssign.QueryAssignedLicenses($vCenter.InstanceUuid.AssignedLicense)
        $vCenterLicense = $vCenterAssignedLicense | Where-Object {$_.EntityId -eq $vCenter.InstanceUuid}
        if ($vCenterLicense -and $Options.ShowLicenses) { 
            $vCenterLicenseKey = $vCenterLicense.AssignedLicense.LicenseKey
        } elseif ($vCenterLicense) { 
            $vCenterLicenseKey = "*****-*****-*****" + $vCenterLicense.AssignedLicense.LicenseKey.Substring(17)
        } else {
            $vCenterLicenseKey = 'No License Key'
        }
        $LicenseObject = [PSCustomObject]@{                               
            Product = $vCenterLicense.AssignedLicense.Name
            LicenseKey = $vCenterLicenseKey                    
        }
    }
    if ($Licenses) {
        foreach ($License in $LicenseManager.Licenses) {
            if ($Options.ShowLicenses) {
                $LicenseKey = $License.LicenseKey
            } else {
                $LicenseKey = "*****-*****-*****" + $License.LicenseKey.Substring(17)
            }
            $Object = [PSCustomObject]@{                               
                'Product' = $License.Name
                'LicenseKey' = $LicenseKey
                'Total' = $License.Total
                'Used' = $License.Used                     
            }
            $LicenseObject += $Object
        }
    }
    Write-Output $LicenseObject
}

function Get-VMHostNetworkAdapterCDP {
    <#
    .SYNOPSIS
    Function to retrieve the Network Adapter CDP info of a vSphere host.
    .DESCRIPTION
    Function to retrieve the Network Adapter CDP info of a vSphere host.
    .PARAMETER VMHost
    A vSphere ESXi Host object
    .INPUTS
    System.Management.Automation.PSObject.
    .OUTPUTS
    System.Management.Automation.PSObject.
    .EXAMPLE
    PS> Get-VMHostNetworkAdapterCDP -VMHost ESXi01,ESXi02
    .EXAMPLE
    PS> Get-VMHost ESXi01,ESXi02 | Get-VMHostNetworkAdapterCDP
    #>
    [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]

    Param
    (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]$VMHosts   
    )    

    begin {
        $CDPObject = @()
    }

    process {
        try {
            foreach ($VMHost in $VMHosts) {
                $ConfigManagerView = Get-View $VMHost.ExtensionData.ConfigManager.NetworkSystem
                $pNics = $ConfigManagerView.NetworkInfo.Pnic
                foreach ($pNic in $pNics) {
                    $PhysicalNicHintInfo = $ConfigManagerView.QueryNetworkHint($pNic.Device)
                    $Object = [PSCustomObject]@{                            
                        'VMHost' = $VMHost.Name
                        'Device' = $pNic.Device
                        'Status' = if ($PhysicalNicHintInfo.ConnectedSwitchPort) {
                            'Connected'
                        } else {
                            'Disconnected'
                        }
                        'SwitchId' = $PhysicalNicHintInfo.ConnectedSwitchPort.DevId
                        'Address' = $PhysicalNicHintInfo.ConnectedSwitchPort.Address
                        'VLAN' = $PhysicalNicHintInfo.ConnectedSwitchPort.Vlan
                        'MTU' = $PhysicalNicHintInfo.ConnectedSwitchPort.Mtu
                        'SystemName' = $PhysicalNicHintInfo.ConnectedSwitchPort.SystemName
                        'Location' = $PhysicalNicHintInfo.ConnectedSwitchPort.Location
                        'HardwarePlatform' = $PhysicalNicHintInfo.ConnectedSwitchPort.HardwarePlatform
                        'SoftwareVersion' = $PhysicalNicHintInfo.ConnectedSwitchPort.SoftwareVersion
                        'ManagementAddress' = $PhysicalNicHintInfo.ConnectedSwitchPort.MgmtAddr
                        'PortId' = $PhysicalNicHintInfo.ConnectedSwitchPort.PortId
                    }
                    $CDPObject += $Object
                }
            }
        } catch [Exception] {
            throw 'Unable to retrieve CDP info'
        }
    }
    end {
        Write-Output $CDPObject
    }
}

function Get-InstallDate {
    $esxcli = Get-EsxCli -VMHost $VMHost -V2 -Server $ESXiHost
    $thisUUID = $esxcli.system.uuid.get.Invoke()
    $decDate = [Convert]::ToInt32($thisUUID.Split("-")[0], 16)
    $installDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($decDate))
    [PSCustomObject][Ordered]@{
        Name = $VMHost.Name
        InstallDate = $installDate
    }
}

function Get-Uptime {
    [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]
    Param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$VMHost, [PSObject]$VM
    )
    $UptimeObject = @()
    $Date = (Get-Date).ToUniversalTime() 
    If ($VMHost) {
        $UptimeObject = Get-View -ViewType hostsystem -Property Name, Runtime.BootTime -Filter @{
            "Name" = "^$($VMHost.Name)$"
            "Runtime.ConnectionState" = "connected"
        } | Select-Object Name, @{L = 'UptimeDays'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalDays), 2)}}, @{L = 'UptimeHours'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalHours), 2)}}, @{L = 'UptimeMinutes'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalMinutes), 2)}}
    }

    if ($VM) {
        $UptimeObject = Get-View -ViewType VirtualMachine -Property Name, Runtime.BootTime -Filter @{
            "Name" = "^$($VM.Name)$"
            "Runtime.PowerState" = "poweredOn"
        } | Select-Object Name, @{L = 'UptimeDays'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalDays), 2)}}, @{L = 'UptimeHours'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalHours), 2)}}, @{L = 'UptimeMinutes'; E = {[math]::round(((($Date) - ($_.Runtime.BootTime)).TotalMinutes), 2)}}
    }
    Write-Output $UptimeObject
}

function Get-ESXiBootDevice {
    <#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function identifies how an ESXi host was booted up along with its boot
        device (if applicable). This supports both local installation to Auto Deploy as
        well as Boot from SAN.
    .PARAMETER VMHostname
        The name of an individual ESXi host managed by vCenter Server
    .EXAMPLE
        Get-ESXiBootDevice
    .EXAMPLE
        Get-ESXiBootDevice -VMHost esxi-01
    #>
    param(
        [Parameter(Mandatory = $false)][PSObject]$VMHost
    )

    $results = @()
    $esxcli = Get-EsxCli -V2 -VMHost $vmhost -Server $ESXiHost
    $bootDetails = $esxcli.system.boot.device.get.Invoke()

    # Check to see if ESXi booted over the network
    $networkBoot = $false
    if ($bootDetails.BootNIC) {
        $networkBoot = $true
        $bootDevice = $bootDetails.BootNIC
    } elseif ($bootDetails.StatelessBootNIC) {
        $networkBoot = $true
        $bootDevice = $bootDetails.StatelessBootNIC
    }

    # If ESXi booted over network, check to see if deployment
    # is Stateless, Stateless w/Caching or Stateful
    if ($networkBoot) {
        $option = $esxcli.system.settings.advanced.list.CreateArgs()
        $option.option = "/UserVars/ImageCachedSystem"
        try {
            $optionValue = $esxcli.system.settings.advanced.list.Invoke($option)
        } catch {
            $bootType = "stateless"
        }
        $bootType = $optionValue.StringValue
    }

    # Loop through all storage devices to identify boot device
    $devices = $esxcli.storage.core.device.list.Invoke()
    $foundBootDevice = $false
    foreach ($device in $devices) {
        if ($device.IsBootDevice -eq $true) {
            $foundBootDevice = $true

            if ($device.IsLocal -eq $true -and $networkBoot -and $bootType -ne "stateful") {
                $bootType = "stateless caching"
            } elseif ($device.IsLocal -eq $true -and $networkBoot -eq $false) {
                $bootType = "local"
            } elseif ($device.IsLocal -eq $false -and $networkBoot -eq $false) {
                $bootType = "remote"
            }

            $bootDevice = $device.Device
            $bootModel = $device.Model
            $bootVendor = $device.VEndor
            $bootSize = $device.Size
            $bootIsSAS = $device.IsSAS
            $bootIsSSD = $device.IsSSD
            $bootIsUSB = $device.IsUSB
        }
    }

    # Pure Stateless (e.g. No USB or Disk for boot)
    if ($networkBoot -and $foundBootDevice -eq $false) {
        $bootModel = "N/A"
        $bootVendor = "N/A"
        $bootSize = "N/A"
        $bootIsSAS = "N/A"
        $bootIsSSD = "N/A"
        $bootIsUSB = "N/A"
    }

    $tmp = [PSCustomObject]@{
        Host = $vmhost.Name;
        Device = $bootDevice;
        BootType = $bootType;
        Vendor = $bootVendor;
        Model = $bootModel;
        SizeMB = $bootSize;
        IsSAS = $bootIsSAS;
        IsSSD = $bootIsSSD;
        IsUSB = $bootIsUSB;
    }
    $results += $tmp
    $results
}

function Get-ScsiDeviceDetail {
    <#
        .SYNOPSIS
        Helper function to return Scsi device information for a specific host and a specific datastore.
        .PARAMETER VMHosts
        This parameter accepts a list of host objects returned from the Get-VMHost cmdlet
        .PARAMETER VMHostMoRef
        This parameter specifies, by MoRef Id, the specific host of interest from with the $VMHosts array.
        .PARAMETER DatastoreDiskName
        This parameter specifies, by disk name, the specific datastore of interest.
        .EXAMPLE
        $VMHosts = Get-VMHost
        Get-ScsiDeviceDetail -AllVMHosts $VMHosts -VMHostMoRef 'HostSystem-host-131' -DatastoreDiskName 'naa.6005076801810082480000000001d9fe'

        DisplayName      : IBM Fibre Channel Disk (naa.6005076801810082480000000001d9fe)
        Ssd              : False
        LocalDisk        : False
        CanonicalName    : naa.6005076801810082480000000001d9fe
        Vendor           : IBM
        Model            : 2145
        Multipath Policy : Round Robin
        CapacityGB       : 512
        .NOTES
        Author: Ryan Kowalewski
    #>

    [CmdLetBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $VMHosts,
        [Parameter(Mandatory = $true)]
        $VMHostMoRef,
        [Parameter(Mandatory = $true)]
        $DatastoreDiskName
    )

    $VMHostObj = $VMHosts | Where-Object {$_.Id -eq $VMHostMoRef}
    $ScsiDisk = $VMHostObj.ExtensionData.Config.StorageDevice.ScsiLun | Where-Object {
        $_.CanonicalName -eq $DatastoreDiskName
    }
    $Multipath = $VMHostObj.ExtensionData.Config.StorageDevice.MultipathInfo.Lun | Where-Object {
        $_.Lun -eq $ScsiDisk.Key
    }
    switch ($Multipath.Policy.Policy) {
        'VMW_PSP_RR' { $MultipathPolicy = 'Round Robin' }
        'VMW_PSP_FIXED' { $MultipathPolicy = 'Fixed' }
        'VMW_PSP_MRU' { $MultipathPolicy = 'Most Recently Used'}
        default { $MultipathPolicy = $Multipath.Policy.Policy }
    }
    $CapacityGB = [math]::Round((($ScsiDisk.Capacity.BlockSize * $ScsiDisk.Capacity.Block) / 1024 / 1024 / 1024), 2)

    [PSCustomObject]@{
        'DisplayName' = $ScsiDisk.DisplayName
        'Ssd' = $ScsiDisk.Ssd
        'LocalDisk' = $ScsiDisk.LocalDisk
        'CanonicalName' = $ScsiDisk.CanonicalName
        'Vendor' = $ScsiDisk.Vendor
        'Model' = $ScsiDisk.Model
        'MultipathPolicy' = $MultipathPolicy
        'CapacityGB' = $CapacityGB
    }
}

Function Get-PciDeviceDetail {
    <#
    .SYNOPSIS
    Helper function to return PCI Devices Drivers & Firmware information for a specific host.
    .PARAMETER Server
    vCenter VISession object.
    .PARAMETER esxcli
    Esxcli session object associated to the host.
    .EXAMPLE
    $Credentials = Get-Crendentials
    $Server = Connect-VIServer -Server vcenter01.example.com -Credentials $Credentials
    $VMHost = Get-VMHost -Server $Server -Name esx01.example.com
    $esxcli = Get-EsxCli -Server $Server -VMHost $VMHost -V2
    Get-PciDeviceDetail -Server $ESXiHost -esxcli $esxcli
    VMkernel Name    : vmhba0
    Device Name      : Sunrise Point-LP AHCI Controller
    Driver           : vmw_ahci
    Driver Version   : 1.0.0-34vmw.650.0.14.5146846
    Firmware Version : NA
    VIB Name         : vmw-ahci
    VIB Version      : 1.0.0-34vmw.650.0.14.5146846
    .NOTES
    Author: Erwan Quelin heavily based on the work of the vDocumentation team - https://github.com/arielsanchezmora/vDocumentation/blob/master/powershell/vDocumentation/Public/Get-ESXIODevice.ps1
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        $Server,
        [Parameter(Mandatory = $true)]
        $esxcli
    )
    Begin {}
    
    Process {
        # Set default results
        $firmwareVersion = "N/A"
        $vibName = "N/A"
        $driverVib = @{
            Name = "N/A"
            Version = "N/A"
        }
        $pciDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -or $_.VMKernelName -like "vmnic*" -or $_.VMKernelName -like "vmgfx*"} | Sort-Object -Property VMKernelName 
        foreach ($pciDevice in $pciDevices) {
            $driverVersion = $esxcli.system.module.get.Invoke(@{module = $pciDevice.ModuleName}) | Select-Object -ExpandProperty Version
            # Get NIC Firmware version
            if ($pciDevice.VMKernelName -like 'vmnic*') {
                $vmnicDetail = $esxcli.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
                $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
                # Get NIC driver VIB package version
                $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net-" + $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net55-" + $vmnicDetail.DriverInfo.Driver}
                <#
                    If HP Smart Array vmhba* (scsi-hpsa driver) then get Firmware version
                    else skip if VMkernnel is vmhba*. Can't get HBA Firmware from 
                    Powercli at the moment only through SSH or using Putty Plink+PowerCli.
                #>
            } elseif ($pciDevice.VMKernelName -like 'vmhba*') {
                if ($pciDevice.DeviceName -match "smart array") {
                    $hpsa = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "HP Smart Array"}
                    if ($hpsa) {
                        $firmwareVersion = (($hpsa.Name -split "firmware")[1]).Trim()
                    }
                }
                # Get HBA driver VIB package version
                $vibName = $pciDevice.ModuleName -replace "_", "-"
                $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq "scsi-" + $VibName -or $_.Name -eq "sata-" + $VibName -or $_.Name -eq $VibName}
            }
            # Output collected data
            [PSCustomObject]@{
                'VMkernel Name' = $pciDevice.VMKernelName
                'Device Name' = $pciDevice.DeviceName
                'Driver' = $pciDevice.ModuleName
                'Driver Version' = $driverVersion
                'Firmware Version' = $firmwareVersion
                'VIB Name' = $driverVib.Name
                'VIB Version' = $driverVib.Version
            } 
        } 
    }
    End {}
    
}
#endregion Script Functions

#region Script Body
#---------------------------------------------------------------------------------------------#
#                                         SCRIPT BODY                                         #
#---------------------------------------------------------------------------------------------#

# Connect to ESXi Host using supplied credentials
foreach ($VIServer in $Target) { 
    $ESXiHost = Connect-VIServer $VIServer -Credential $Credentials -ErrorAction SilentlyContinue
    
    # Create a lookup hashtable to quickly link VM MoRefs to Names
    # Exclude VMware Site Recovery Manager placeholder VMs
    $VMs = Get-VM -Server $ESXiHost | Where-Object {
        $_.ExtensionData.Config.ManagedBy.ExtensionKey -notlike 'com.vmware.vcDr*'
    } | Sort-Object Name
    $VMLookup = @{}
    foreach ($VM in $VMs) {
        $VMLookup.($VM.Id) = $VM.Name
    }

    #region ESXi VMHost Section
    if ($InfoLevel.VMHost -ge 1) {
        $VMHost = Get-VMHost -Server $ESXiHost
        Section -Style Heading2 "$($VMHost.Name)" {
            Paragraph ("The following section provides information on the configuration of VMware ESXi host $($VMHost.Name).")

            #region ESXi Host Detailed Information
            ### TODO: Host Certificate, Swap File Location
            #region ESXi Host Hardware Section
            Section -Style Heading4 'Hardware' {
                Paragraph ("The following section provides information on the host hardware configuration of $($VMHost.Name).")
                BlankLine

                #region ESXi Host Specifications
                $VMHostUptime = Get-Uptime -VMHost $VMHost
                $esxcli = Get-EsxCli -VMHost $VMHost -V2 -Server $ESXiHost
                $VMHostHardware = Get-VMHostHardware -VMHost $VMHost
                #$VMHostLicense = Get-License -VMHost $VMHost
                $ScratchLocation = Get-AdvancedSetting -Entity $VMHost | Where-Object {$_.Name -eq 'ScratchConfig.CurrentScratchLocation'}
                $VMHostDetail = [PSCustomObject]@{
                    'Name' = $VMHost.Name
                    'Connection State' = Switch ($VMHost.ConnectionState) {
                        'NotResponding' {'Not Responding'}
                        default {$VMHost.ConnectionState}
                    }
                    'ID' = $VMHost.Id
                    'Parent' = $VMHost.Parent
                    'Manufacturer' = $VMHost.Manufacturer
                    'Model' = $VMHost.Model
                    'Serial Number' = $VMHostHardware.SerialNumber 
                    'Asset Tag' = $VMHostHardware.AssetTag 
                    'Processor Type' = $VMHost.Processortype
                    'HyperThreading' = Switch ($VMHost.HyperthreadingActive) {
                        $true {'Enabled'}
                        $false {'Disabled'}
                    }
                    'Number of CPU Sockets' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuPackages 
                    'Number of CPU Cores' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuCores 
                    'Number of CPU Threads' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuThreads
                    'CPU Speed' = "$([math]::Round(($VMHost.ExtensionData.Hardware.CpuInfo.Hz) / 1000000000, 2)) GHz" 
                    'Memory' = "$([math]::Round($VMHost.MemoryTotalGB, 0)) GB" 
                    'NUMA Nodes' = $VMHost.ExtensionData.Hardware.NumaInfo.NumNodes 
                    'Number of NICs' = $VMHostHardware.NicCount 
                    'Number of Datastores' = $VMHost.ExtensionData.Datastore.Count 
                    'Number of VMs' = $VMHost.ExtensionData.VM.Count 
                    'Power Management Policy' = $VMHost.ExtensionData.Hardware.CpuPowerManagementInfo.CurrentPolicy 
                    'Scratch Location' = $ScratchLocation.Value 
                    'Bios Version' = $VMHost.ExtensionData.Hardware.BiosInfo.BiosVersion 
                    'Bios Release Date' = $VMHost.ExtensionData.Hardware.BiosInfo.ReleaseDate 
                    'ESXi Version' = $VMHost.Version 
                    'ESXi Build' = $VMHost.build 
                    #'Product' = $VMHostLicense.Product 
                    #'License Key' = $VMHostLicense.LicenseKey 
                    'Boot Time' = $VMHost.ExtensionData.Runtime.Boottime 
                    'Uptime Days' = $VMHostUptime.UptimeDays
                }
                if ($Healthcheck.VMHost.ConnectionState) {
                    $VMHostDetail | Where-Object {$_.'Connection State' -eq 'Maintenance'} | Set-Style -Style Warning -Property 'Connection State'
                }
                if ($Healthcheck.VMHost.HyperThreading) {
                    $VMHostDetail | Where-Object {$_.'HyperThreading' -eq 'Disabled'} | Set-Style -Style Warning -Property 'Disabled'
                }
                if ($Healthcheck.VMHost.Licensing) {
                    $VMHostDetail | Where-Object {$_.'Product' -like '*Evaluation*'} | Set-Style -Style Warning -Property 'Product'
                    $VMHostDetail | Where-Object {$_.'License Key' -like '*-00000-00000'} | Set-Style -Style Warning -Property 'License Key'
                }
                if ($Healthcheck.VMHost.ScratchLocation) {
                    $VMHostDetail | Where-Object {$_.'Scratch Location' -eq '/tmp/scratch'} | Set-Style -Style Warning -Property 'Scratch Location'
                }
                if ($Healthcheck.VMHost.UpTimeDays) {
                    $VMHostDetail | Where-Object {$_.'Uptime Days' -ge 275 -and $_.'Uptime Days' -lt 365} | Set-Style -Style Warning -Property 'Uptime Days'
                    $VMHostDetail | Where-Object {$_.'Uptime Days' -ge 365} | Set-Style -Style Warning -Property 'Uptime Days'
                }
                $VMHostDetail | Table -Name "$VMHost ESXi Host Detailed Information" -List -ColumnWidths 50, 50 
                #endregion ESXi Host Specifications

                #region ESXi Host Boot Device
                Section -Style Heading5 'Boot Device' {
                    $ESXiBootDevice = Get-ESXiBootDevice -VMHost $VMHost
                    $VMHostBootDevice = [PSCustomObject]@{
                        'Host' = $ESXiBootDevice.Host
                        'Device' = $ESXiBootDevice.Device
                        'Boot Type' = $ESXiBootDevice.BootType
                        'Vendor' = $ESXiBootDevice.Vendor
                        'Model' = $ESXiBootDevice.Model
                        'Size' = "$([math]::Round($ESXiBootDevice.SizeMB / 1024), 2) GB"
                        'Is SAS' = Switch ($ESXiBootDevice.IsSAS) {
                            $true {'Yes'}
                            $false {'No'}
                        }
                        'Is SSD' = Switch ($ESXiBootDevice.IsSSD) {
                            $true {'Yes'}
                            $false {'No'}
                        }
                        'Is USB' = Switch ($ESXiBootDevice.IsUSB) {
                            $true {'Yes'}
                            $false {'No'}
                        }
                    }
                    $VMHostBootDevice | Table -Name "$VMHost Boot Device" -List -ColumnWidths 50, 50 
                }
                #endregion ESXi Host Boot Devices

                #region ESXi Host PCI Devices
                Section -Style Heading5 'PCI Devices' {
                    $PciHardwareDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -OR $_.VMKernelName -like "vmnic*" -OR $_.VMKernelName -like "vmgfx*"} 
                    $VMHostPciDevices = foreach ($PciHardwareDevice in $PciHardwareDevices) {
                        [PSCustomObject]@{
                            'VMkernel Name' = $PciHardwareDevice.VMkernelName 
                            'PCI Address' = $PciHardwareDevice.Address 
                            'Device Class' = $PciHardwareDevice.DeviceClassName 
                            'Device Name' = $PciHardwareDevice.DeviceName 
                            'Vendor Name' = $PciHardwareDevice.VendorName 
                            'Slot Description' = $PciHardwareDevice.SlotDescription
                        }
                    }
                    $VMHostPciDevices | Sort-Object 'VMkernel Name' | Table -Name "$VMHost PCI Devices" 
                }
                #endregion ESXi Host PCI Devices
                            
                #region ESXi Host PCI Devices Drivers & Firmware
                Section -Style Heading5 'PCI Devices Drivers & Firmware' {
                    $VMHostPciDevicesDetails = Get-PciDeviceDetail -Server $ESXiHost -esxcli $esxcli 
                    $VMHostPciDevicesDetails | Sort-Object 'VMkernel Name' | Table -Name "$VMHost PCI Devices Drivers & Firmware" 
                }
                #endregion ESXi Host PCI Devices Drivers & Firmware
            }
            #endregion ESXi Host Hardware Section

            #region ESXi Host System Section
            Section -Style Heading4 'System' {
                Paragraph ("The following section provides information on the host system configuration of $($VMHost.Name).")

                #region ESXi Host Image Profile Information
                Section -Style Heading5 'Image Profile' {
                    $installdate = Get-InstallDate
                    $esxcli = Get-ESXCli -VMHost $VMHost -V2 -Server $ESXiHost
                    $ImageProfile = $esxcli.software.profile.get.Invoke()
                    $SecurityProfile = [PSCustomObject]@{
                        'Image Profile' = $ImageProfile.Name
                        'Vendor' = $ImageProfile.Vendor
                        'Installation Date' = $InstallDate.InstallDate
                    }
                    $SecurityProfile | Table -Name "$VMHost Image Profile" -ColumnWidths 50, 25, 25 
                }
                #endregion ESXi Host Image Profile Information

                #region ESXi Host Time Configuration
                Section -Style Heading5 'Time Configuration' {
                    $VMHostTimeSettings = [PSCustomObject]@{
                        'Time Zone' = $VMHost.timezone
                        'NTP Service' = Switch ((Get-VMHostService -VMHost $VMHost | Where-Object {$_.key -eq 'ntpd'}).Running) {
                            $true {'Running'}
                            $false {'Stopped'}
                        }
                        'NTP Server(s)' = (Get-VMHostNtpServer -VMHost $VMHost | Sort-Object) -join ', '
                    }
                    if ($Healthcheck.VMHost.TimeConfig) {
                        $VMHostTimeSettings | Where-Object {$_.'NTP Service' -eq 'Stopped'} | Set-Style -Style Critical -Property 'NTP Service'
                    }
                    $VMHostTimeSettings | Table -Name "$VMHost Time Configuration" -ColumnWidths 30, 30, 40
                }
                #endregion ESXi Host Time Configuration

                #region ESXi Host Syslog Configuration
                $SyslogConfig = $VMHost | Get-VMHostSysLogServer
                if ($SyslogConfig) {
                    Section -Style Heading5 'Syslog Configuration' {
                        ### TODO: Syslog Rotate & Size, Log Directory (Adv Settings)
                        $SyslogConfig = $SyslogConfig | Select-Object @{L = 'SysLog Server'; E = {$_.Host}}, Port
                        $SyslogConfig | Table -Name "$VMHost Syslog Configuration" -ColumnWidths 50, 50 
                    }
                }
                #endregion ESXi Host Syslog Configuration

                # Set InfoLevel to 5 to provide advanced system information for VMHosts
                if ($InfoLevel.VMHost -ge 5) {
                    #region ESXi Host Advanced System Settings
                    Section -Style Heading5 'Advanced System Settings' {
                        $AdvSettings = $VMHost | Get-AdvancedSetting | Select-Object Name, Value
                        $AdvSettings | Sort-Object Name | Table -Name "$VMHost Advanced System Settings" -ColumnWidths 50, 50 
                    }
                    #endregion ESXi Host Advanced System Settings

                    #region ESXi Host Software VIBs
                    Section -Style Heading5 'Software VIBs' {
                        $esxcli = Get-ESXCli -VMHost $VMHost -V2 -Server $ESXiHost
                        $VMHostVibs = $esxcli.software.vib.list.Invoke()
                        $VMHostVibs = foreach ($VMHostVib in $VMHostVibs) {
                            [PSCustomObject]@{
                                'Name' = $VMHostVib.Name
                                'ID' = $VMHostVib.Id
                                'Version' = $VMHostVib.Version
                                'Acceptance Level' = $VMHostVib.AcceptanceLevel
                                'Creation Date' = $VMHostVib.CreationDate
                                'Install Date' = $VMHostVib.InstallDate
                            }
                        } 
                        $VMHostVibs | Sort-Object 'Install Date' -Descending | Table -Name "$VMHost Software VIBs" -ColumnWidths 10, 25, 20, 10, 15, 10, 10
                    }
                    #endregion ESXi Host Software VIBs
                }
            }
            #endregion ESXi Host System Section

            #region ESXi Host Storage Section
            Section -Style Heading4 'Storage' {
                Paragraph ("The following section provides information on the host " +
                    "storage configuration of $($VMHost.Name).")
            
                #region ESXi Host Datastore Specifications
                Section -Style Heading5 'Datastores' {
                    $VMHostDatastores = $VMHost | Get-Datastore          
                    $VMHostDsSpecs = foreach ($VMHostDatastore in $VMHostDatastores) {
                        [PSCustomObject]@{
                            'Name' = $VMHostDatastore.Name
                            'Type' = $VMHostDatastore.Type
                            'Version' = $VMHostDatastore.FileSystemVersion
                            '# of VMs' = $VMHostDatastore.ExtensionData.VM.Count
                            'Total Capacity GB' = [math]::Round($VMHostDatastore.CapacityGB, 2)
                            'Used Capacity GB' = [math]::Round((($VMHostDatastore.CapacityGB) - ($VMHostDatastore.FreeSpaceGB)), 2)
                            'Free Space GB' = [math]::Round($VMHostDatastore.FreeSpaceGB, 2)
                            '% Used' = [math]::Round((100 - (($VMHostDatastore.FreeSpaceGB) / ($VMHostDatastore.CapacityGB) * 100)), 2)
                        }
                    }
                    if ($Healthcheck.Datastore.CapacityUtilization) {
                        $VMHostDsSpecs | Where-Object {$_.'% Used' -ge 90} | Set-Style -Style Critical
                        $VMHostDsSpecs | Where-Object {$_.'% Used' -ge 75 -and $_.'% Used' -lt 90} | Set-Style -Style Warning
                    }
                    $VMHostDsSpecs | Sort-Object Name | Table -Name "$VMHost Datastores" #-ColumnWidths 20,10,10,10,10,10,10,10,10
                }
                #endregion ESXi Host Datastore Specifications
            
                #region ESXi Host Storage Adapter Information
                $VMHostHba = $VMHost | Get-VMHostHba | Where-Object {$_.type -eq 'FibreChannel' -or $_.type -eq 'iSCSI' }
                if ($VMHostHba) {
                    Section -Style Heading5 'Storage Adapters' {
                        $VMHostHbaFC = $VMHost | Get-VMHostHba -Type FibreChannel
                        if ($VMHostHbaFC) {
                            Paragraph ("The following table details the fibre channel storage adapters for $($VMHost.Name).")
                            Blankline
                            $VMHostHbaFC = $VMHost | Get-VMHostHba -Type FibreChannel | Select-Object Device, Type, Model, Driver, 
                            @{L = 'Node WWN'; E = {([String]::Format("{0:X}", $_.NodeWorldWideName) -split "(\w{2})" | Where-Object {$_ -ne ""}) -join ":" }}, 
                            @{L = 'Port WWN'; E = {([String]::Format("{0:X}", $_.PortWorldWideName) -split "(\w{2})" | Where-Object {$_ -ne ""}) -join ":" }}, speed, status
                            $VMHostHbaFC | Sort-Object Device | Table -Name "$VMHost FC Storage Adapters"
                        }

                        $VMHostHbaIScsi = $VMHost | Get-VMHostHba -Type iSCSI
                        if ($VMHostHbaFC -and $VMHostHbaIScsi) {
                            Blankline
                        }
                        if ($VMHostHbaIScsi) {
                            Paragraph ("The following table details the iSCSI storage adapters for $($VMHost.Name).")
                            Blankline
                            $VMHostHbaIScsi = $VMHost | Get-VMHostHba -Type iSCSI | Select-Object Device, @{L = 'iSCSI Name'; E = {$_.IScsiName}}, Model, Driver, @{L = 'Speed'; E = {$_.CurrentSpeedMb}}, status
                            $VMHostHbaIScsi | Sort-Object Device | Table -Name "$VMHost iSCSI Storage Adapters" -List -ColumnWidths 25, 75
                        }
                    }
                }
                #endregion ESXi Host Storage Adapater Information
            }
            #endregion ESXi Host Storage Section

            #region ESXi Host Network Section
            Section -Style Heading4 'Network' {
                Paragraph ("The following section provides information on the host network configuration of $($VMHost.Name).")
                BlankLine
                #region ESXi Host Network Configuration
                $VMHostNetwork = $VMHost.ExtensionData.Config.Network
                $VMHostNetworkDetail = [PSCustomObject]@{
                    'VMHost' = $VMHost.Name 
                    'Virtual Switches' = ($VMHostNetwork.Vswitch.Name | Sort-Object) -join ', '
                    'VMKernel Adapters' = ($VMHostNetwork.Vnic.Device | Sort-Object) -join ', '
                    'Physical Adapters' = ($VMHostNetwork.Pnic.Device | Sort-Object) -join ', '
                    'VMKernel Gateway' = $VMHostNetwork.IpRouteConfig.DefaultGateway
                    'IPv6 Enabled' = $VMHostNetwork.IPv6Enabled
                    'VMKernel IPv6 Gateway' = $VMHostNetwork.IpRouteConfig.IpV6DefaultGateway
                    'DNS Servers' = ($VMHostNetwork.DnsConfig.Address | Sort-Object) -join ', ' 
                    'Host Name' = $VMHostNetwork.DnsConfig.HostName
                    'Domain Name' = $VMHostNetwork.DnsConfig.DomainName 
                    'Search Domain' = ($VMHostNetwork.DnsConfig.SearchDomain | Sort-Object) -join ', '
                }
                if ($Healthcheck.VMHost.IPv6Enabled) {
                    $VMHostNetworkDetail | Where-Object {$_.'IPv6 Enabled' -eq $false} | Set-Style -Style Warning -Property 'IPv6 Enabled'
                }
                $VMHostNetworkDetail | Table -Name "$VMHost Network Configuration" -List -ColumnWidths 50, 50
                #endregion ESXi Host Network Configuration

                #region ESXi Host Physical Adapters
                Section -Style Heading5 'Physical Adapters' {
                    Paragraph ("The following table details the physical network adapters for $($VMHost.Name).")
                    BlankLine

                    $PhysicalNetAdapters = $VMHost.ExtensionData.Config.Network.Pnic
                    $VMHostPhysicalNetAdapter = foreach ($PhysicalNetAdapter in $PhysicalNetAdapters) {
                        [PSCustomObject]@{
                            'Device' = $PhysicalNetAdapter.Device
                            'Status' = Switch ($PhysicalNetAdapter.Linkspeed) {
                                $null {'Disconnected'}
                                default {'Connected'}
                            }
                            'vSwitch' = foreach ($vSwitch in $VMHost.ExtensionData.Config.Network.Vswitch) {
                                foreach ($pNic in $vSwitch.Pnic) {
                                    if ($pNic -eq $PhysicalNetAdapter.Key) {
                                        $vSwitch.Name
                                    }
                                }
                            }
                            'MAC Address' = $PhysicalNetAdapter.Mac
                            'Actual Speed, Duplex' = Switch ($PhysicalNetAdapter.LinkSpeed.SpeedMb) {
                                $null {'Down'}
                                default {
                                    if ($PhysicalNetAdapter.LinkSpeed.Duplex) {
                                        "$($PhysicalNetAdapter.LinkSpeed.SpeedMb) Mb, Full Duplex"
                                    } else {
                                        'Auto negotiate'
                                    }
                                }
                            }
                            'Configured Speed, Duplex' = Switch ($PhysicalNetAdapter.Spec.LinkSpeed) {
                                $null {'Auto negotiate'}
                                default {
                                    if ($PhysicalNetAdapter.Spec.LinkSpeed.Duplex) {
                                        "$($PhysicalNetAdapter.Spec.LinkSpeed.SpeedMb) Mb, Full Duplex"
                                    } else {
                                        "$($PhysicalNetAdapter.Spec.LinkSpeed.SpeedMb) Mb"
                                    }
                                }
                            }
                            'Wake on LAN' = Switch ($PhysicalNetAdapter.WakeOnLanSupported) {
                                $true {'Supported'}
                                $false {'Not Supported'}
                            }
                        }
                    }
                    if ($InfoLevel.VMHost -ge 4) {
                        $VMHostPhysicalNetAdapter | Sort-Object 'Device' | Table -List -Name "$VMHost Network Physical Adapters" -ColumnWidths 50, 50
                    } else {
                        $VMHostPhysicalNetAdapter | Sort-Object 'Device' | Table -Name "$VMHost Network Physical Adapters"
                    }
                }
                #endregion ESXi Host Physical Adapters
                                
                #region ESXi Host Cisco Discovery Protocol
                $VMHostNetworkAdapterCDP = @()
                $VMHostNetworkAdapterCDP = $VMHost | Get-VMHostNetworkAdapterCDP | Where-Object {$_.Status -eq 'Connected'}
                if ($VMHostNetworkAdapterCDP) {
                    Section -Style Heading5 'Cisco Discovery Protocol' {
                        if ($InfoLevel.VMHost -ge 4) {
                            $VMHostCDP = $VMHostNetworkAdapterCDP | Select-Object Device, Status, @{L = 'Hardware Platform'; E = {$_.HardwarePlatform}},
                            @{L = 'Software Version'; E = {$_.SoftwareVersion}}, @{L = 'Switch'; E = {$_.SwitchId}}, @{L = 'Management Address'; E = {$_.ManagementAddress}}, @{L = 'Switch ID'; E = {$_.SwitchId}}, Address, @{L = 'Port ID'; E = {$_.PortId}}, VLAN, MTU
                            $VMHostCDP | Sort-Object Device | Table -List -Name "$VMHost Network Adapter CDP Information" -ColumnWidths 50, 50
                        } else {
                            $VMHostCDP = $VMHostNetworkAdapterCDP | Select-Object Device, Status, @{L = 'Hardware Platform'; E = {$_.HardwarePlatform}},
                            @{L = 'Switch'; E = {$_.SwitchId}}, @{L = 'Management Address'; E = {$_.ManagementAddress}}, @{L = 'Port ID'; E = {$_.PortId}}
                            $VMHostCDP | Sort-Object Device | Table -Name "$VMHost Network Adapter CDP Information" #-ColumnWidths 20, 20, 20, 20, 20
                        }
                    }
                }
                #endregion ESXi Host Cisco Discovery Protocol

                #region ESXi Host VMkernel Adapaters
                Section -Style Heading5 'VMkernel Adapters' {
                    Paragraph "The following table details the VMkernel adapters for $($VMHost.Name)."
                    BlankLine

                    $VMkernelAdapters = $VMHost | Get-VMHostNetworkAdapter -VMKernel
                    $VMHostVmkAdapters = foreach ($VMkernelAdapter in $VMkernelAdapters) {
                        [PSCustomObject]@{
                            'Device' = $VMkernelAdapter.DeviceName 
                            'Port Group' = $VMkernelAdapter.PortGroupName 
                            'MTU' = $VMkernelAdapter.Mtu 
                            'MAC Address' = $VMkernelAdapter.Mac
                            'IP Address' = $VMkernelAdapter.IP 
                            'Subnet Mask' = $VMkernelAdapter.SubnetMask 
                            'vMotion Traffic' = Switch ($VMkernelAdapter.vMotionEnabled) {
                                $true {'Enabled'}
                                $false {'Disabled'}
                            }
                            'FT Logging' = Switch ($VMkernelAdapter.FaultToleranceLoggingEnabled) {
                                $true {'Enabled'}
                                $false {'Disabled'}
                            }
                            'Management Traffic' = Switch ($VMkernelAdapter.ManagementTrafficEnabled) {
                                $true {'Enabled'}
                                $false {'Disabled'}
                            }
                            'vSAN Traffic' = Switch ($VMkernelAdapter.VsanTrafficEnabled) {
                                $true {'Enabled'}
                                $false {'Disabled'}
                            }
                        }
                    }
                    $VMHostVmkAdapters | Sort-Object 'Device' | Table -Name "$VMHost VMkernel Adapters" -List -ColumnWidths 50, 50 
                                    
                    <#
                                    $VMkernelAdapters = $VMHost.ExtensionData.Config.Network.Vnic
                                    $VMHostVmkAdapters = foreach ($VMkernelAdapter in $VMkernelAdapters) {
                                        [PSCustomObject]@{
                                            'Device' = $VMkernelAdapter.Device
                                            'Port Group' = Switch ($VMkernelAdapter.Spec.PortGroup) {
                                                $null {$VMkernelAdapter.Spec.DistributedVirtualPort}
                                                default {$VMkernelAdapter.Spec.PortGroup}
                                            }
                                            'TCP/IP stack' = Switch ($VMkernelAdapter.Spec.NetStackInstanceKey) {
                                                'defaultTcpipStack' {'Default'}
                                                'vmotion' {'vMotion'}
                                                'vSphereProvisioning' {'Provisioning'}
                                                default {$VMkernelAdapter.Spec.NetStackInstanceKey}
                                            }
                                            'MTU' = $VMkernelAdapter.Spec.Mtu
                                            'MAC Address' = $VMkernelAdapter.Mac
                                            'DHCP' = Switch ($VMkernelAdapter.Spec.IP.Dhcp) {
                                                $true {'Enabled'}
                                                $false {'Disabled'}
                                            }
                                            'IP Address' = $VMkernelAdapter.Spec.IP.IPAddress
                                            'Subnet Mask' = $VMkernelAdapter.Spec.IP.SubnetMask
                                        }
                                    }
                                    $VMHostVmkAdapters | Sort-Object 'Device' | Table -Name "$VMHost VMkernel Adapters" -List -ColumnWidths 50, 50 
                                    #>
                }
                #endregion ESXi Host VMkernel Adapaters

                #region ESXi Host Virtual Switches
                $VSSwitches = $VMHost | Get-VirtualSwitch -Standard | Sort-Object Name
                if ($VSSwitches) {
                    Section -Style Heading5 'Standard Virtual Switches' {
                        Paragraph ("The following sections detail the standard virtual " +
                            "switch configuration for $($VMHost.Name).")
                        BlankLine
                        $VSSwitchNicTeaming = $VSSwitches | Get-NicTeamingPolicy
                        $VSSGeneral = foreach ($VSSwitchNicTeam in $VSSwitchNicTeaming) {
                            [PSCustomObject]@{
                                'Name' = $VSSwitchNicTeam.VirtualSwitch 
                                'MTU' = $VSSwitchNicTeam.VirtualSwitch.Mtu 
                                'Number of Ports' = $VSSwitchNicTeam.VirtualSwitch.NumPorts
                                'Number of Ports Available' = $VSSwitchNicTeam.VirtualSwitch.NumPortsAvailable 
                                'Load Balancing' = Switch ($VSSwitchNicTeam.LoadBalancingPolicy) {
                                    'LoadbalanceSrcId' {'Route based on the originating port ID'}
                                    'LoadbalanceSrcMac' {'Route based on source MAC hash'}
                                    'LoadbalanceIP' {'Route based on IP hash'}
                                    'ExplicitFailover' {'Explicit Failover'}
                                }
                                'Failover Detection' = Switch ($VSSwitchNicTeam.NetworkFailoverDetectionPolicy) {
                                    'LinkStatus' {'Link Status'}
                                    'BeaconProbing' {'Beacon Probing'}
                                } 
                                'Notify Switches' = Switch ($VSSwitchNicTeam.NotifySwitches) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                }
                                'Failback' = Switch ($VSSwitchNicTeam.FailbackEnabled) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                } 
                                'Active NICs' = (($VSSwitchNicTeam.ActiveNic | Sort-Object) -join ', ') 
                                'Standby NICs' = (($VSSwitchNicTeam.StandbyNic | Sort-Object) -join ', ')
                                'Unused NICs' = (($VSSwitchNicTeam.UnusedNic | Sort-Object) -join ', ')
                            }
                        }
                        $VSSGeneral | Table -Name "$VMHost Standard Virtual Switches" -List -ColumnWidths 50, 50
                    }
                    #region ESXi Host Virtual Switch Security Policy
                    $VssSecurity = $VSSwitches | Get-SecurityPolicy
                    if ($VssSecurity) {
                        Section -Style Heading5 'Virtual Switch Security Policy' {
                            $VssSecurity = foreach ($VssSec in $VssSecurity) {
                                [PSCustomObject]@{
                                    'vSwitch' = $VssSec.VirtualSwitch 
                                    'MAC Address Changes' = Switch ($VssSec.MacChanges) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    } 
                                    'Forged Transmits' = Switch ($VssSec.ForgedTransmits) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    } 
                                    'Promiscuous Mode' = Switch ($VssSec.AllowPromiscuous) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                }
                            }
                            $VssSecurity | Sort-Object 'vSwitch' | Table -Name "$VMHost vSwitch Security Policy" #-ColumnWidths 25, 25, 25, 25
                        }
                    }
                    #endregion ESXi Host Virtual Switch Security Policy                  

                    #region ESXi Host Virtual Switch NIC Teaming
                    $VssPortgroupNicTeaming = $VSSwitches | Get-NicTeamingPolicy
                    if ($VssPortgroupNicTeaming) {
                        Section -Style Heading5 'Virtual Switch NIC Teaming' {
                            $VssPortgroupNicTeaming = foreach ($VssPortgroupNicTeam in $VssPortgroupNicTeaming) {
                                [PSCustomObject]@{
                                    'vSwitch' = $VssPortgroupNicTeam.VirtualSwitch 
                                    'Load Balancing' = Switch ($VssPortgroupNicTeam.LoadBalancingPolicy) {
                                        'LoadbalanceSrcId' {'Route based on the originating port ID'}
                                        'LoadbalanceSrcMac' {'Route based on source MAC hash'}
                                        'LoadbalanceIP' {'Route based on IP hash'}
                                        'ExplicitFailover' {'Explicit Failover'}
                                    }
                                    'Failover Detection' = Switch ($VssPortgroupNicTeam.NetworkFailoverDetectionPolicy) {
                                        'LinkStatus' {'Link Status'}
                                        'BeaconProbing' {'Beacon Probing'}
                                    } 
                                    'Notify Switches' = Switch ($VssPortgroupNicTeam.NotifySwitches) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                    'Failback' = Switch ($VssPortgroupNicTeam.FailbackEnabled) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                    'Active NICs' = (($VssPortgroupNicTeam.ActiveNic | Sort-Object) -join [Environment]::NewLine)
                                    'Standby NICs' = (($VssPortgroupNicTeam.StandbyNic | Sort-Object) -join [Environment]::NewLine)
                                    'Unused NICs' = (($VssPortgroupNicTeam.UnusedNic | Sort-Object) -join [Environment]::NewLine)
                                }
                            }
                            $VssPortgroupNicTeaming | Sort-Object 'vSwitch' | Table -Name "$VMHost vSwitch NIC Teaming"
                        }
                    }
                    #endregion ESXi Host Virtual Switch NIC Teaming                       
                    
                    #region ESXi Host Virtual Switch Port Groups
                    $VssPortgroups = $VSSwitches | Get-VirtualPortGroup -Standard 
                    if ($VssPortgroups) {
                        Section -Style Heading5 'Virtual Port Groups' {
                            $VssPortgroups = foreach ($VssPortgroup in $VssPortgroups) {
                                [PSCustomObject]@{
                                    'vSwitch' = $VssPortgroup.VirtualSwitchName 
                                    'Port Group' = $VssPortgroup.Name 
                                    'VLAN ID' = $VssPortgroup.VLanId 
                                    '# of VMs' = ($VssPortgroup | Get-VM).Count
                                }
                            }
                            $VssPortgroups | Sort-Object 'vSwitch', 'Port Group' | Table -Name "$VMHost vSwitch Port Group Information"
                        }
                    }
                    #endregion ESXi Host Virtual Switch Port Groups                
                    
                    #region ESXi Host Virtual Switch Port Group Security Poilicy
                    $VssPortgroupSecurity = $VSSwitches | Get-VirtualPortGroup | Get-SecurityPolicy 
                    if ($VssPortgroupSecurity) {
                        Section -Style Heading5 'Virtual Port Group Security Policy' {
                            $VssPortgroupSecurity = foreach ($VssPortgroupSec in $VssPortgroupSecurity) {
                                [PSCustomObject]@{
                                    'vSwitch' = $VssPortgroupSec.virtualportgroup.virtualswitchname 
                                    'Port Group' = $VssPortgroupSec.VirtualPortGroup 
                                    'MAC Changes' = Switch ($VssPortgroupSec.MacChanges) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                    'Forged Transmits' = Switch ($VssPortgroupSec.ForgedTransmits) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    } 
                                    'Promiscuous Mode' = Switch ($VssPortgroupSec.AllowPromiscuous) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                }
                            }
                            $VssPortgroupSecurity | Sort-Object 'vSwitch', 'Port Group' | Table -Name "$VMHost vSwitch Port Group Security Policy" 
                        }
                    }
                    #endregion ESXi Host Virtual Switch Port Group Security Poilicy                 

                    #region ESXi Host Virtual Switch Port Group NIC Teaming
                    $VssPortgroupNicTeaming = $VSSwitches | Get-VirtualPortGroup  | Get-NicTeamingPolicy 
                    if ($VssPortgroupNicTeaming) {
                        Section -Style Heading5 'Virtual Port Group NIC Teaming' {
                            $VssPortgroupNicTeaming = foreach ($VssPortgroupNicTeam in $VssPortgroupNicTeaming) {
                                [PSCustomObject]@{
                                    'vSwitch' = $VssPortgroupNicTeam.virtualportgroup.virtualswitchname 
                                    'Port Group' = $VssPortgroupNicTeam.VirtualPortGroup 
                                    'Load Balancing' = Switch ($VssPortgroupNicTeam.LoadBalancingPolicy) {
                                        'LoadbalanceSrcId' {'Route based on the originating port ID'}
                                        'LoadbalanceSrcMac' {'Route based on source MAC hash'}
                                        'LoadbalanceIP' {'Route based on IP hash'}
                                        'ExplicitFailover' {'Explicit Failover'}
                                    }
                                    'Failover Detection' = Switch ($VssPortgroupNicTeam.NetworkFailoverDetectionPolicy) {
                                        'LinkStatus' {'Link Status'}
                                        'BeaconProbing' {'Beacon Probing'}
                                    }  
                                    'Notify Switches' = Switch ($VssPortgroupNicTeam.NotifySwitches) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    }
                                    'Failback' = Switch ($VssPortgroupNicTeam.FailbackEnabled) {
                                        $true {'Enabled'}
                                        $false {'Disabled'}
                                    } 
                                    'Active NICs' = (($VssPortgroupNicTeam.ActiveNic | Sort-Object) -join [Environment]::NewLine)
                                    'Standby NICs' = (($VssPortgroupNicTeam.StandbyNic | Sort-Object) -join [Environment]::NewLine)
                                    'Unused NICs' = (($VssPortgroupNicTeam.UnusedNic | Sort-Object) -join [Environment]::NewLine)
                                }
                            }
                            $VssPortgroupNicTeaming | Sort-Object 'vSwitch', 'Port Group' | Table -Name "$VMHost vSwitch Port Group NIC Teaming"
                        }
                    }
                    #endregion ESXi Host Virtual Switch Port Group NIC Teaming                      
                }
                #endregion ESXi Host Standard Virtual Switches
            }                
            #endregion ESXi Host Network Configuration

            #region ESXi Host Security Section
            Section -Style Heading4 'Security' {
                Paragraph ("The following section provides information on the host " +
                    "security configuration of $($VMHost.Name).")
                                
                #region ESXi Host Lockdown Mode
                if ($VMHost.ExtensionData.Config.LockdownMode -ne $null) {
                    Section -Style Heading5 'Lockdown Mode' {
                        $LockdownMode = [PSCustomObject]@{
                            'Lockdown Mode' = Switch ($VMHost.ExtensionData.Config.LockdownMode) {
                                'lockdownDisabled' {'Disabled'}
                                'lockdownNormal' {'Enabled (Normal)'}
                                'lockdownStrict' {'Enabled (Strict)'}
                            }
                        }
                        if ($Healthcheck.VMHost.LockdownMode) {
                            $LockdownMode | Where-Object {$_.'Lockdown Mode' -eq 'Disabled'} | Set-Style -Style Warning -Property 'Lockdown Mode'
                        }
                        $LockdownMode | Table -Name "$VMHost Lockdown Mode" -List -ColumnWidths 50, 50
                    }
                }
                #endregion ESXi Host Lockdown Mode

                #region ESXi Host Services
                Section -Style Heading5 'Services' {
                    $VMHostServices = $VMHost | Get-VMHostService
                    $Services = foreach ($VMHostService in $VMHostServices) {
                        [PSCustomObject]@{
                            'Name' = $VMHostService.Label
                            'Daemon' = Switch ($VMHostService.Running) {
                                $true {'Running'}
                                $false {'Stopped'}
                            }
                            'Startup Policy' = Switch ($VMHostService.Policy) {
                                'automatic' {'Start and stop with port usage'}
                                'on' {'Start and stop with host'}
                                'off' {'Start and stop manually'}
                            }
                        }
                    }
                    if ($Healthcheck.VMHost.Services) {
                        $Services | Where-Object {$_.'Name' -eq 'SSH' -and $_.Daemon -eq 'Running'} | Set-Style -Style Warning -Property 'Daemon'
                        $Services | Where-Object {$_.'Name' -eq 'ESXi Shell' -and $_.Daemon -eq 'Running'} | Set-Style -Style Warning -Property 'Daemon'
                        $Services | Where-Object {$_.'Name' -eq 'NTP Daemon' -and $_.Daemon -eq 'Stopped'} | Set-Style -Style Critical -Property 'Daemon'
                        $Services | Where-Object {$_.'Name' -eq 'NTP Daemon' -and $_.'Startup Policy' -ne 'Start and stop with host'} | Set-Style -Style Critical -Property 'Startup Policy'
                    }
                    $Services | Sort-Object Name | Table -Name "$VMHost Services" 
                }
                #endregion ESXi Host Services

                if ($InfoLevel.VMHost -ge 4) {
                    #region ESXi Host Firewall
                    Section -Style Heading5 'Firewall' {
                        $Firewall = $VMHost | Get-VMHostFirewallException | Sort-Object Name | Select-Object Name, Enabled, @{L = 'Incoming Ports'; E = {$_.IncomingPorts}}, @{L = 'Outgoing Ports'; E = {$_.OutgoingPorts}}, Protocols, @{L = 'Service Running'; E = {$_.ServiceRunning}}
                        $Firewall | Table -Name "$VMHost Firewall Configuration" 
                    }
                    #endregion ESXi Host Firewall
                }
                
                #region ESXi Host Authentication
                $AuthServices = $VMHost | Get-VMHostAuthentication
                if ($AuthServices.DomainMembershipStatus) {
                    Section -Style Heading5 'Authentication Services' {
                        $AuthServices = $AuthServices | Select-Object Domain, @{L = 'Domain Membership'; E = {$_.DomainMembershipStatus}}, @{L = 'Trusted Domains'; E = {$_.TrustedDomains}}
                        $AuthServices | Table -Name "$VMHost Authentication Services" -ColumnWidths 25, 25, 50 
                    }    
                }
                #endregion ESXi Host Authentication
            }
            #endregion ESXi Host Security Section

            #region ESXi Host Virtual Machines Section
            Section -Style Heading2 'Virtual Machines' {
                Paragraph ("The following section provides information on Virtual Machines " +
                    "managed by $($VMHost.Name).")

                #region Virtual Machine Informative Information
                if ($InfoLevel.VM -eq 2) {
                    BlankLine
                    $VMSummary = foreach ($VM in $VMs) {
                        [PSCustomObject]@{
                            'Name' = $VM.Name
                            'Power State' = Switch ($VM.PowerState) {
                                'PoweredOn' {'Powered On'}
                                'PoweredOff' {'Powered Off'}
                            }
                            'vCPUs' = $VM.NumCpu
                            'Cores per Socket' = $VM.CoresPerSocket
                            'Memory GB' = [math]::Round(($VM.MemoryGB), 2)
                            'Provisioned GB' = [math]::Round(($VM.ProvisionedSpaceGB), 2)
                            'Used GB' = [math]::Round(($VM.UsedSpaceGB), 2)
                            'HW Version' = $VM.HardwareVersion
                            'VM Tools Status' = Switch ($VM.ExtensionData.Guest.ToolsStatus) {
                                'toolsOld' {'Tools Old'}
                                'toolsOk' {'Tools OK'}
                                'toolsNotRunning' {'Tools Not Running'}
                                'toolsNotInstalled' {'Tools Not Installed'}
                            }         
                        }
                    }
                    if ($Healthcheck.VM.VMToolsOK) {
                        $VMSummary | Where-Object {$_.'VM Tools Status' -ne 'Tools OK'} | Set-Style -Style Warning -Property 'VM Tools Status'
                    }
                    if ($Healthcheck.VM.PoweredOn) {
                        $VMSummary | Where-Object {$_.'Power State' -ne 'Powered On'} | Set-Style -Style Warning -Property 'Power State'
                    }
                    $VMSummary | Table -Name 'VM Informative Information'
                }
                #endregion Virtual Machine Informative Information

                #region Virtual Machine Detailed Information
                if ($InfoLevel.VM -ge 3) {
                    ## TODO: More VM Details to Add
                    foreach ($VM in $VMs) {
                        Section -Style Heading3 $VM.name {
                            $VMUptime = Get-Uptime -VM $VM
                            $VMDetail = [PSCustomObject]@{
                                'Name' = $VM.Name
                                'ID' = $VM.Id 
                                'Operating System' = $VM.ExtensionData.Summary.Config.GuestFullName
                                'Hardware Version' = $VM.HardwareVersion
                                'Power State' = Switch ($VM.PowerState) {
                                    'PoweredOn' {'Powered On'}
                                    'PoweredOff' {'Powered Off'}
                                }
                                'VM Tools Status' = Switch ($VM.ExtensionData.Guest.ToolsStatus) {
                                    'toolsOld' {'Tools Old'}
                                    'toolsOk' {'Tools OK'}
                                    'toolsNotRunning' {'Tools Not Running'}
                                    'toolsNotInstalled' {'Tools Not Installed'}
                                }
                                'Fault Tolerance State' = Switch ($VM.ExtensionData.Runtime.FaultToleranceState) {
                                    'notConfigured' {'Not Configured'}
                                    'needsSecondary' {'Needs Secondary'}
                                    'running' {'Running'}
                                    'disabled' {'Disabled'}
                                    'starting' {'Starting'}
                                    'enabled' {'Enabled'}
                                } 
                                'Host' = $VM.VMHost.Name
                                'Parent' = $VM.VMHost.Parent.Name
                                'Parent Folder' = $VM.Folder.Name
                                'Parent Resource Pool' = $VM.ResourcePool.Name
                                'vCPUs' = $VM.NumCpu
                                'Cores per Socket' = $VM.CoresPerSocket
                                'CPU Resources' = "$($VM.VMResourceConfiguration.CpuSharesLevel) / $($VM.VMResourceConfiguration.NumCpuShares)"
                                'CPU Reservation' = $VM.VMResourceConfiguration.CpuReservationMhz
                                'CPU Limit' = "$($VM.VMResourceConfiguration.CpuReservationMhz) MHz" 
                                'CPU Hot Add' = Switch ($VM.ExtensionData.Config.CpuHotAddEnabled) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                }
                                'CPU Hot Remove' = Switch ($VM.ExtensionData.Config.CpuHotRemoveEnabled) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                } 
                                'Memory Allocation' = "$([math]::Round(($VM.memoryGB), 2)) GB" 
                                'Memory Resources' = "$($VM.VMResourceConfiguration.MemSharesLevel) / $($VM.VMResourceConfiguration.NumMemShares)"
                                'Memory Hot Add' = Switch ($VM.ExtensionData.Config.MemoryHotAddEnabled) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                }
                                'vDisks' = $VM.ExtensionData.Summary.Config.NumVirtualDisks
                                'Used Space' = "$([math]::Round(($VM.UsedSpaceGB), 2)) GB"
                                'Provisioned Space' = "$([math]::Round(($VM.ProvisionedSpaceGB), 2)) GB"
                                'Changed Block Tracking' = Switch ($VM.ExtensionData.Config.ChangeTrackingEnabled) {
                                    $true {'Enabled'}
                                    $false {'Disabled'}
                                }
                                'vNICs' = $VM.ExtensionData.Summary.Config.NumEthernetCards
                                'Notes' = $VM.Notes
                                'Boot Time' = $VM.ExtensionData.Runtime.BootTime
                                'Uptime Days' = $VMUptime.UptimeDays
                            }
                            
                            if ($Healthcheck.VM.VMToolsOK) {
                                $VMDetail | Where-Object {$_.'VM Tools Status' -ne 'Tools OK'} | Set-Style -Style Warning -Property 'VM Tools Status'
                            }
                            if ($Healthcheck.VM.PoweredOn) {
                                $VMDetail | Where-Object {$_.'Power State' -ne 'Powered On'} | Set-Style -Style Warning -Property 'Power State'
                            }
                            if ($Healthcheck.VM.CpuHotAddEnabled) {
                                $VMDetail | Where-Object {$_.'CPU Hot Add' -eq 'Enabled'} | Set-Style -Style Warning -Property 'CPU Hot Add'
                            }
                            if ($Healthcheck.VM.CpuHotRemoveEnabled) {
                                $VMDetail | Where-Object {$_.'CPU Hot Remove' -eq 'Enabled'} | Set-Style -Style Warning -Property 'CPU Hot Remove'
                            } 
                            if ($Healthcheck.VM.MemoryHotAddEnabled) {
                                $VMDetail | Where-Object {$_.'Memory Hot Add' -eq 'Enabled'} | Set-Style -Style Warning -Property 'Memory Hot Add'
                            } 
                            if ($Healthcheck.VM.ChangeBlockTrackingEnabled) {
                                $VMDetail | Where-Object {$_.'Changed Block Tracking' -eq 'Disabled'} | Set-Style -Style Warning -Property 'Changed Block Tracking'
                            } 
                            $VMDetail | Table -Name 'VM Detailed Information' -List -ColumnWidths 50, 50
                        }
                    } 
                }
                #endregion Virtual Machine Detailed Information

                #region VM Snapshot Information
                $Script:VMSnapshots = $VMs | Get-Snapshot 
                if ($VMSnapshots) {
                    Section -Style Heading3 'VM Snapshots' {
                        $VMSnapshotInfo = foreach ($VMSnapshot in $VMSnapshots) {
                            [PSCustomObject]@{
                                'Virtual Machine' = $VMSnapshot.VM
                                'Name' = $VMSnapshot.Name
                                'Description' = $VMSnapshot.Description
                                'Days Old' = ((Get-Date).ToUniversalTime() - $VMSnapshot.Created).Days
                            } 
                        }
                        if ($Healthcheck.VM.VMSnapshots) {
                            $VMSnapshotInfo | Where-Object {$_.'Days Old' -ge 7} | Set-Style -Style Warning 
                            $VMSnapshotInfo | Where-Object {$_.'Days Old' -ge 14} | Set-Style -Style Critical
                        }
                        $VMSnapshotInfo | Table -Name 'VM Snapshot Information'
                    }
                }
                #endregion VM Snapshot Information
            }

            #endregion ESXi Host Virtual Machines Section
        }
        #endregion ESXi Host Detailed Information

    }
    #endregion ESXi VMHost Section 

}
# Disconnect ESXi Host
$Null = Disconnect-VIServer -Server $ESXiHost -Confirm:$false -ErrorAction SilentlyContinue
#endregion Script Body