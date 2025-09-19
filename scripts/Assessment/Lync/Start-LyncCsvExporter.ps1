<#
.SYNOPSIS
    Launches the Lync CSV Export Tool with Phone Inventory Support
.DESCRIPTION
    This script provides an interactive menu for exporting Lync/Skype for Business data to CSV files.
    It allows users to select the type of data they want to export and the output format.
    
    Export Categories:
    - User data (Summary, Voice users, SBA users, Complete)
    - Phone/Device inventory (Common area phones, Analog devices, USB devices)
    - Infrastructure (Pools, Policies)
    - Bulk exports (All data types)
    
    The exported CSV files will be saved in the specified output directory and can be imported
    into Excel, Power BI, or other analysis tools for reporting and inventory management.
    
.PARAMETER OutputDirectory
    The directory where the exported CSV files will be saved.
.PARAMETER OrganizationName
    The name of the organization for which the report is being generated.
.PARAMETER SBAPattern
    The pattern used to identify Survivable Branch Appliances (SBA).
.EXAMPLE
    Start-LyncCsvExporter.ps1 -OutputDirectory "C:\Reports\CSV_Exports" -OrganizationName "Contoso" -SBAPattern "*MSSBA*"
    This command starts the Lync CSV Export Tool with the specified output directory, organization name, and SBA pattern.
.NOTES
    Author: W. Ford
    Date: 2025-09-17
    Version: 2.0
    
    Version 2.0 Changes:
    - Added phone inventory export functionality
    - Support for Common Area Phones, Analog Devices, and USB Devices
    - Enhanced menu structure with phone inventory section
    - Updated bulk export to include all device types
#>
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = "C:\Reports\CSV_Exports",
    
    [Parameter(Mandatory=$false)]
    [string]$OrganizationName = "Organization",
    
    [Parameter(Mandatory=$false)]
    [string]$SBAPattern = "*MSSBA*"
)

$Separator = "=" * 60

# Create output directory
if (!(Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force
    Write-Host "Created directory: $OutputDirectory" -ForegroundColor Green
}

# Function to export users
function Export-LyncUsers {
    param([string]$Type)
    
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFile = "$OutputDirectory\Lync_Users_$Type`_$Timestamp.csv"
    
    try {
        Write-Host "Exporting Lync Users ($Type)..." -ForegroundColor Yellow
        $Users = Get-CsUser -ErrorAction Stop
        
        switch ($Type) {
            "Summary" {
                $ExportData = $Users | Select-Object @{
                    Name = 'DisplayName'; Expression = { $_.DisplayName }
                }, @{
                    Name = 'SIPAddress'; Expression = { $_.SipAddress }
                }, @{
                    Name = 'UPN'; Expression = { $_.UserPrincipalName }
                }, @{
                    Name = 'Enabled'; Expression = { $_.Enabled }
                }, @{
                    Name = 'Pool'; Expression = { $_.RegistrarPool }
                }, @{
                    Name = 'VoiceEnabled'; Expression = { $_.EnterpriseVoiceEnabled }
                }, @{
                    Name = 'LineURI'; Expression = { $_.LineURI }
                }, @{
                    Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                }
            }
            
            "Voice" {
                $VoiceUsers = $Users | Where-Object { $_.EnterpriseVoiceEnabled -eq $true }
                $ExportData = $VoiceUsers | Select-Object @{
                    Name = 'DisplayName'; Expression = { $_.DisplayName }
                }, @{
                    Name = 'SIPAddress'; Expression = { $_.SipAddress }
                }, @{
                    Name = 'LineURI'; Expression = { $_.LineURI }
                }, @{
                    Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                }, @{
                    Name = 'VoiceRoutingPolicy'; Expression = { $_.VoiceRoutingPolicy }
                }, @{
                    Name = 'HostedVoiceMail'; Expression = { $_.HostedVoiceMail }
                }, @{
                    Name = 'Pool'; Expression = { $_.RegistrarPool }
                }, @{
                    Name = 'LocationPolicy'; Expression = { $_.LocationPolicy }
                }
            }
            
            "Complete" {
                $ExportData = $Users | Select-Object DisplayName, FirstName, LastName, SipAddress, UserPrincipalName, 
                    SamAccountName, Enabled, RegistrarPool, HomeServer, EnterpriseVoiceEnabled, LineURI, 
                    HostedVoiceMail, VoicePolicy, VoiceRoutingPolicy, ConferencingPolicy, ClientPolicy, 
                    LocationPolicy, MobilityPolicy, ExternalAccessPolicy, EnabledForFederation, 
                    EnabledForInternetAccess, EnabledForRichPresence, PublicNetworkEnabled, 
                    DistinguishedName, WhenCreated, WhenChanged, Guid, Sid
            }
            
            "SBA" {
                $SBAUsers = $Users | Where-Object { $_.RegistrarPool -like $SBAPattern }
                $ExportData = $SBAUsers | Select-Object @{
                    Name = 'DisplayName'; Expression = { $_.DisplayName }
                }, @{
                    Name = 'SIPAddress'; Expression = { $_.SipAddress }
                }, @{
                    Name = 'SchoolSite'; Expression = { ($_.RegistrarPool -split '-')[0] }
                }, @{
                    Name = 'SBAPool'; Expression = { $_.RegistrarPool }
                }, @{
                    Name = 'VoiceEnabled'; Expression = { $_.EnterpriseVoiceEnabled }
                }, @{
                    Name = 'LineURI'; Expression = { $_.LineURI }
                }, @{
                    Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                }
            }
        }
        
        $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Exported $($ExportData.Count) records to: $OutputFile" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ Error exporting users: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to export pools
function Export-LyncPools {
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFile = "$OutputDirectory\Lync_Pools_$Timestamp.csv"
    
    try {
        Write-Host "Exporting Lync Pools..." -ForegroundColor Yellow
        $Pools = Get-CsPool -ErrorAction Stop
        
        $ExportData = $Pools | Select-Object @{
            Name = 'PoolFQDN'; Expression = { $_.Fqdn }
        }, @{
            Name = 'Identity'; Expression = { $_.Identity }
        }, @{
            Name = 'Site'; Expression = { $_.Site }
        }, @{
            Name = 'PoolType'; Expression = { 
                if ($_.Fqdn -like $SBAPattern) { "SBA Branch Appliance" }
                elseif ($_.Fqdn -like "*ivr*") { "IVR Pool" }
                elseif ($_.Fqdn -like "*edge*") { "Edge Pool" }
                elseif ($_.Fqdn -like "*lync*") { "Lync Pool" }
                elseif ($_.Fqdn -match '^\d+\.\d+\.\d+\.\d+$') { "IP Address Pool" }
                else { "Standard Pool" }
            }
        }, @{
            Name = 'Services'; Expression = { $_.Services -join '; ' }
        }, @{
            Name = 'Computers'; Expression = { $_.Computers -join '; ' }
        }, @{
            Name = 'BackupPoolFqdn'; Expression = { $_.BackupPoolFqdn }
        }
        
        $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Exported $($ExportData.Count) pools to: $OutputFile" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ Error exporting pools: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to export phone inventory
function Export-PhoneInventory {
    param([string]$Type)
    
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    try {
        switch ($Type) {
            "CommonArea" {
                Write-Host "Exporting Common Area Phones..." -ForegroundColor Yellow
                $OutputFile = "$OutputDirectory\Lync_CommonAreaPhones_$Timestamp.csv"
                $CommonAreaPhones = Get-CsCommonAreaPhone -ErrorAction SilentlyContinue
                if ($CommonAreaPhones) {
                    $ExportData = $CommonAreaPhones | Select-Object @{
                        Name = 'DisplayName'; Expression = { $_.DisplayName }
                    }, @{
                        Name = 'SIPAddress'; Expression = { $_.SipAddress }
                    }, @{
                        Name = 'LineURI'; Expression = { $_.LineURI }
                    }, @{
                        Name = 'Pool'; Expression = { $_.RegistrarPool }
                    }, @{
                        Name = 'Description'; Expression = { $_.Description }
                    }, @{
                        Name = 'Location'; Expression = { $_.Location }
                    }, @{
                        Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                    }, @{
                        Name = 'Enabled'; Expression = { $_.Enabled }
                    }, @{
                        Name = 'DeviceModel'; Expression = { 
                            if ($_.UserAgent) { $_.UserAgent }
                            elseif ($_.ClientVersionFilter) { $_.ClientVersionFilter }
                            else { "Unknown" }
                        }
                    }, @{
                        Name = 'MACAddress'; Expression = { $_.MACAddress }
                    }, @{
                        Name = 'IPAddress'; Expression = { $_.IPAddress }
                    }, @{
                        Name = 'Manufacturer'; Expression = { 
                            if ($_.UserAgent -match "Polycom") { "Polycom" }
                            elseif ($_.UserAgent -match "Cisco") { "Cisco" }
                            elseif ($_.UserAgent -match "Yealink") { "Yealink" }
                            elseif ($_.UserAgent -match "AudioCodes") { "AudioCodes" }
                            elseif ($_.UserAgent -match "Snom") { "Snom" }
                            elseif ($_.UserAgent -match "Grandstream") { "Grandstream" }
                            else { "Unknown" }
                        }
                    }, @{
                        Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }
                    }
                    
                    $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                    Write-Host "✅ Exported $($ExportData.Count) common area phones to: $OutputFile" -ForegroundColor Green
                } else {
                    Write-Host "⚠️  No common area phones found" -ForegroundColor Yellow
                }
            }
            
            "Analog" {
                Write-Host "Exporting Analog Devices..." -ForegroundColor Yellow
                $OutputFile = "$OutputDirectory\Lync_AnalogDevices_$Timestamp.csv"
                $AnalogDevices = Get-CsAnalogDevice -ErrorAction SilentlyContinue
                if ($AnalogDevices) {
                    $ExportData = $AnalogDevices | Select-Object @{
                        Name = 'DisplayName'; Expression = { $_.DisplayName }
                    }, @{
                        Name = 'SIPAddress'; Expression = { $_.SipAddress }
                    }, @{
                        Name = 'LineURI'; Expression = { $_.LineURI }
                    }, @{
                        Name = 'Gateway'; Expression = { $_.Gateway }
                    }, @{
                        Name = 'Port'; Expression = { $_.Port }
                    }, @{
                        Name = 'AnalogFaxEnabled'; Expression = { $_.AnalogFaxEnabled }
                    }, @{
                        Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                    }, @{
                        Name = 'Enabled'; Expression = { $_.Enabled }
                    }, @{
                        Name = 'DeviceType'; Expression = { 
                            if ($_.AnalogFaxEnabled -eq $true) { "Fax Machine" }
                            else { "Analog Phone" }
                        }
                    }, @{
                        Name = 'Location'; Expression = { $_.Location }
                    }, @{
                        Name = 'Contact'; Expression = { $_.Contact }
                    }, @{
                        Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }
                    }
                    
                    $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                    Write-Host "✅ Exported $($ExportData.Count) analog devices to: $OutputFile" -ForegroundColor Green
                } else {
                    Write-Host "⚠️  No analog devices found" -ForegroundColor Yellow
                }
            }
            
            "USB" {
                Write-Host "Exporting USB Audio/Communication Devices..." -ForegroundColor Yellow
                $OutputFile = "$OutputDirectory\Lync_UsbDevices_$Timestamp.csv"
                
                try {
                    # Get USB devices that are likely to be communication devices (headsets, speakerphones, etc.)
                    $UsbDevices = Get-WmiObject -Class Win32_PnPEntity | Where-Object { 
                        ($_.Name -match "audio|headset|speakerphone|webcam|camera|microphone|USB.*phone") -and
                        ($_.DeviceID -like "USB*") -and
                        ($_.Status -eq "OK")
                    }
                    
                    if ($UsbDevices) {
                        $ExportData = $UsbDevices | Select-Object @{
                            Name = 'DeviceName'; Expression = { $_.Name }
                        }, @{
                            Name = 'Description'; Expression = { $_.Description }
                        }, @{
                            Name = 'DeviceID'; Expression = { $_.DeviceID }
                        }, @{
                            Name = 'Status'; Expression = { $_.Status }
                        }, @{
                            Name = 'Manufacturer'; Expression = { $_.Manufacturer }
                        }, @{
                            Name = 'Service'; Expression = { $_.Service }
                        }, @{
                            Name = 'ClassGuid'; Expression = { $_.ClassGuid }
                        }, @{
                            Name = 'VendorId'; Expression = { 
                                if ($_.DeviceID -match "VID_([0-9A-F]{4})") { $matches[1] } else { "Unknown" }
                            }
                        }, @{
                            Name = 'ProductId'; Expression = { 
                                if ($_.DeviceID -match "PID_([0-9A-F]{4})") { $matches[1] } else { "Unknown" }
                            }
                        }, @{
                            Name = 'DeviceCategory'; Expression = { 
                                if ($_.Name -match "headset|headphone") { "Headset" }
                                elseif ($_.Name -match "speakerphone") { "Speakerphone" }
                                elseif ($_.Name -match "webcam|camera") { "Webcam" }
                                elseif ($_.Name -match "microphone|mic") { "Microphone" }
                                elseif ($_.Name -match "phone") { "USB Phone" }
                                else { "Audio Device" }
                            }
                        }
                        
                        $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                        Write-Host "✅ Exported $($ExportData.Count) USB communication devices to: $OutputFile" -ForegroundColor Green
                    } else {
                        Write-Host "⚠️  No USB communication devices found" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "⚠️  Error querying USB devices: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "⚠️  Trying alternative method..." -ForegroundColor Yellow
                    
                    # Alternative method using Get-PnpDevice if available (Windows 8+)
                    try {
                        $PnpDevices = Get-PnpDevice -Class "AudioEndpoint","Media" -Status OK -ErrorAction SilentlyContinue | 
                                     Where-Object { $_.InstanceId -like "USB*" -and $_.FriendlyName -match "audio|headset|speakerphone|microphone" }
                        
                        if ($PnpDevices) {
                            $ExportData = $PnpDevices | Select-Object @{
                                Name = 'DeviceName'; Expression = { $_.FriendlyName }
                            }, @{
                                Name = 'Description'; Expression = { $_.Description }
                            }, @{
                                Name = 'DeviceID'; Expression = { $_.InstanceId }
                            }, @{
                                Name = 'Status'; Expression = { $_.Status }
                            }, @{
                                Name = 'Class'; Expression = { $_.Class }
                            }, @{
                                Name = 'Service'; Expression = { $_.Service }
                            }
                            
                            $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                            Write-Host "✅ Exported $($ExportData.Count) USB communication devices to: $OutputFile" -ForegroundColor Green
                        } else {
                            Write-Host "⚠️  No USB communication devices found via alternative method" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "⚠️  Alternative USB device query also failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            }
            
            "ClientVersions" {
                Write-Host "Exporting Registered Device/Client Information..." -ForegroundColor Yellow
                $OutputFile = "$OutputDirectory\Lync_RegisteredDevices_$Timestamp.csv"
                
                try {
                    # Try to get actual registered endpoint information from users
                    $Users = Get-CsUser -ErrorAction SilentlyContinue
                    if ($Users) {
                        $DeviceData = @()
                        
                        foreach ($User in ($Users | Select-Object -First 100)) {  # Limit to first 100 users for performance
                            try {
                                # Try to get user sessions/endpoints if available
                                $UserSessions = Get-CsUserSession -User $User.SipAddress -ErrorAction SilentlyContinue
                                if ($UserSessions) {
                                    foreach ($Session in $UserSessions) {
                                        $DeviceData += [PSCustomObject]@{
                                            UserDisplayName = $User.DisplayName
                                            UserSipAddress = $User.SipAddress
                                            DeviceType = $Session.ClientType
                                            UserAgent = $Session.UserAgent
                                            Endpoint = $Session.Endpoint
                                            ConnectionTime = $Session.ConnectionTime
                                            Status = $Session.Status
                                        }
                                    }
                                }
                            } catch {
                                # Skip individual user errors
                            }
                        }
                        
                        if ($DeviceData.Count -gt 0) {
                            $DeviceData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                            Write-Host "✅ Exported $($DeviceData.Count) registered device sessions to: $OutputFile" -ForegroundColor Green
                        } else {
                            Write-Host "⚠️  No active user sessions/devices found via Get-CsUserSession" -ForegroundColor Yellow
                        }
                    }
                } catch {
                    Write-Host "⚠️  Error querying user sessions: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                
                # Fallback: Get client version configuration as additional info
                try {
                    $ClientVersionFile = "$OutputDirectory\Lync_ClientVersionConfig_$Timestamp.csv"
                    $ClientVersions = Get-CsClientVersionConfiguration -ErrorAction SilentlyContinue
                    if ($ClientVersions) {
                        $ExportData = $ClientVersions | Select-Object @{
                            Name = 'Identity'; Expression = { $_.Identity }
                        }, @{
                            Name = 'Enabled'; Expression = { $_.Enabled }
                        }, @{
                            Name = 'DefaultAction'; Expression = { $_.DefaultAction }
                        }, @{
                            Name = 'DefaultURL'; Expression = { $_.DefaultURL }
                        }
                        
                        $ExportData | Export-Csv -Path $ClientVersionFile -NoTypeInformation -Encoding UTF8
                        Write-Host "✅ Exported $($ExportData.Count) client version configurations to: $ClientVersionFile" -ForegroundColor Green
                    } else {
                        Write-Host "⚠️  No client version configuration found" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "⚠️  Error getting client version configuration: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            "IPPhones" {
                Write-Host "Exporting IP Phone Devices..." -ForegroundColor Yellow
                $OutputFile = "$OutputDirectory\Lync_IPPhones_$Timestamp.csv"
                
                try {
                    # Get all Lync users and filter for potential phone devices
                    $AllUsers = Get-CsUser -ErrorAction SilentlyContinue
                    if ($AllUsers) {
                        Write-Host "  Analyzing $($AllUsers.Count) user accounts for phone devices..." -ForegroundColor Gray
                        
                        # Look for users that are likely phone devices based on various criteria
                        $PhoneDevices = $AllUsers | Where-Object { 
                            # Check display name patterns
                            ($_.DisplayName -match "phone|device|room|conference|lobby|reception|meeting|board|kitchen|break|common|area|front.*desk|copier|printer|fax|emergency|elevator|security|kiosk") -or
                            # Check SIP address patterns  
                            ($_.SipAddress -match "phone|device|room|conference|lobby|reception|meeting|board|kitchen|break|common|area|front.*desk|copier|printer|fax|emergency|elevator|security|kiosk") -or
                            # Check if they have LineURI but no typical user attributes
                            (($null -ne $_.LineURI) -and ([string]::IsNullOrEmpty($_.FirstName) -or [string]::IsNullOrEmpty($_.LastName))) -or
                            # Check for generic/service account patterns
                            ($_.SamAccountName -match "svc|service|phone|device|room|common") -or
                            # Check for accounts without mailbox (typically devices)
                            ($_.EnterpriseVoiceEnabled -eq $true -and [string]::IsNullOrEmpty($_.EmailAddress))
                        }
                        
                        if ($PhoneDevices) {
                            $ExportData = $PhoneDevices | Select-Object @{
                                Name = 'DisplayName'; Expression = { $_.DisplayName }
                            }, @{
                                Name = 'SIPAddress'; Expression = { $_.SipAddress }
                            }, @{
                                Name = 'LineURI'; Expression = { $_.LineURI }
                            }, @{
                                Name = 'Pool'; Expression = { $_.RegistrarPool }
                            }, @{
                                Name = 'DeviceType'; Expression = { 
                                    if ($_.DisplayName -match "conference|meeting|board") { "Conference Room Phone" }
                                    elseif ($_.DisplayName -match "lobby|reception|front") { "Reception/Lobby Phone" }
                                    elseif ($_.DisplayName -match "common|area|kitchen|break") { "Common Area Phone" }
                                    elseif ($_.DisplayName -match "emergency|security") { "Emergency/Security Phone" }
                                    elseif ($_.DisplayName -match "elevator|kiosk") { "Public Access Phone" }
                                    else { "IP Phone Device" }
                                }
                            }, @{
                                Name = 'Location'; Expression = { 
                                    # Try to extract location from display name
                                    if ($_.DisplayName -match "(\w+\s*\d+|\w+\s+room|\w+\s+floor)") { $matches[1] } 
                                    else { "Unknown" }
                                }
                            }, @{
                                Name = 'VoiceEnabled'; Expression = { $_.EnterpriseVoiceEnabled }
                            }, @{
                                Name = 'VoicePolicy'; Expression = { $_.VoicePolicy }
                            }, @{
                                Name = 'Enabled'; Expression = { $_.Enabled }
                            }, @{
                                Name = 'SamAccountName'; Expression = { $_.SamAccountName }
                            }, @{
                                Name = 'FirstName'; Expression = { $_.FirstName }
                            }, @{
                                Name = 'LastName'; Expression = { $_.LastName }
                            }, @{
                                Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }
                            }
                            
                            $ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                            Write-Host "✅ Exported $($ExportData.Count) potential IP phone devices to: $OutputFile" -ForegroundColor Green
                            
                            # Display summary by device type
                            $DeviceTypes = $ExportData | Group-Object DeviceType
                            Write-Host "  📊 Device Types Found:" -ForegroundColor Cyan
                            foreach ($Type in $DeviceTypes) {
                                Write-Host "    $($Type.Name): $($Type.Count)" -ForegroundColor White
                            }
                        } else {
                            Write-Host "⚠️  No IP phone devices found based on naming patterns and criteria" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "❌ Unable to retrieve user accounts" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "❌ Error exporting IP phones: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            "All" {
                Write-Host "Exporting All Phone/Device Inventory..." -ForegroundColor Cyan
                Export-PhoneInventory -Type "CommonArea"
                Export-PhoneInventory -Type "Analog" 
                Export-PhoneInventory -Type "IPPhones"
                Export-PhoneInventory -Type "USB"
                Export-PhoneInventory -Type "ClientVersions"
                Write-Host "All phone inventory exports completed!" -ForegroundColor Green
            }
        }
        
    } catch {
        Write-Host "❌ Error exporting phone inventory ($Type): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to export phone numbers
function Export-PhoneNumbers {
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    try {
        Write-Host "Exporting All Phone Numbers (LineURIs)..." -ForegroundColor Yellow
        $PhoneNumberFile = "$OutputDirectory\Lync_AllPhoneNumbers_$Timestamp.csv"
        
        $AllPhoneNumbers = @()
        
        # Get phone numbers from regular users
        Write-Host "  Collecting phone numbers from users..." -ForegroundColor Gray
        $Users = Get-CsUser -ErrorAction SilentlyContinue
        if ($Users) {
            foreach ($User in $Users) {
                if (![string]::IsNullOrEmpty($User.LineURI)) {
                    $PhoneNumber = $User.LineURI -replace "tel:", "" -replace ";.*", ""
                    $Extension = if ($User.LineURI -match ";ext=(\d+)") { $matches[1] } else { "" }
                    
                    $AllPhoneNumbers += [PSCustomObject]@{
                        PhoneNumber = $PhoneNumber
                        Extension = $Extension
                        FullLineURI = $User.LineURI
                        AssignedTo = $User.DisplayName
                        AssignedType = "User"
                        SIPAddress = $User.SipAddress
                        UserPrincipalName = $User.UserPrincipalName
                        Pool = $User.RegistrarPool
                        VoicePolicy = $User.VoicePolicy
                        VoiceEnabled = $User.EnterpriseVoiceEnabled
                        Enabled = $User.Enabled
                        Site = if ($User.RegistrarPool) {
                            # Try to extract site from pool name
                            if ($User.RegistrarPool -match "^(.+?)-" -or $User.RegistrarPool -match "^(.+?)\.") {
                                $matches[1]
                            } else {
                                "Unknown"
                            }
                        } else { "Unknown" }
                        FirstName = $User.FirstName
                        LastName = $User.LastName
                        Department = $User.Department
                        Title = $User.Title
                        DistinguishedName = $User.DistinguishedName
                    }
                }
            }
        }
        
        # Get phone numbers from common area phones
        Write-Host "  Collecting phone numbers from common area phones..." -ForegroundColor Gray
        $CommonAreaPhones = Get-CsCommonAreaPhone -ErrorAction SilentlyContinue
        if ($CommonAreaPhones) {
            foreach ($Phone in $CommonAreaPhones) {
                if (![string]::IsNullOrEmpty($Phone.LineURI)) {
                    $PhoneNumber = $Phone.LineURI -replace "tel:", "" -replace ";.*", ""
                    $Extension = if ($Phone.LineURI -match ";ext=(\d+)") { $matches[1] } else { "" }
                    
                    $AllPhoneNumbers += [PSCustomObject]@{
                        PhoneNumber = $PhoneNumber
                        Extension = $Extension
                        FullLineURI = $Phone.LineURI
                        AssignedTo = $Phone.DisplayName
                        AssignedType = "Common Area Phone"
                        SIPAddress = $Phone.SipAddress
                        UserPrincipalName = ""
                        Pool = $Phone.RegistrarPool
                        VoicePolicy = $Phone.VoicePolicy
                        VoiceEnabled = $true
                        Enabled = $Phone.Enabled
                        Site = if ($Phone.RegistrarPool) {
                            if ($Phone.RegistrarPool -match "^(.+?)-" -or $Phone.RegistrarPool -match "^(.+?)\.") {
                                $matches[1]
                            } else {
                                "Unknown"
                            }
                        } else { "Unknown" }
                        FirstName = ""
                        LastName = ""
                        Department = ""
                        Title = ""
                        DistinguishedName = $Phone.DistinguishedName
                    }
                }
            }
        }
        
        # Get phone numbers from analog devices
        Write-Host "  Collecting phone numbers from analog devices..." -ForegroundColor Gray
        $AnalogDevices = Get-CsAnalogDevice -ErrorAction SilentlyContinue
        if ($AnalogDevices) {
            foreach ($Device in $AnalogDevices) {
                if (![string]::IsNullOrEmpty($Device.LineURI)) {
                    $PhoneNumber = $Device.LineURI -replace "tel:", "" -replace ";.*", ""
                    $Extension = if ($Device.LineURI -match ";ext=(\d+)") { $matches[1] } else { "" }
                    
                    $AllPhoneNumbers += [PSCustomObject]@{
                        PhoneNumber = $PhoneNumber
                        Extension = $Extension
                        FullLineURI = $Device.LineURI
                        AssignedTo = $Device.DisplayName
                        AssignedType = if ($Device.AnalogFaxEnabled) { "Fax Machine" } else { "Analog Device" }
                        SIPAddress = $Device.SipAddress
                        UserPrincipalName = ""
                        Pool = ""
                        VoicePolicy = $Device.VoicePolicy
                        VoiceEnabled = $true
                        Enabled = $Device.Enabled
                        Site = if ($Device.Gateway) {
                            # Try to extract site from gateway name
                            if ($Device.Gateway -match "^(.+?)-" -or $Device.Gateway -match "^(.+?)\.") {
                                $matches[1]
                            } else {
                                "Unknown"
                            }
                        } else { "Unknown" }
                        FirstName = ""
                        LastName = ""
                        Department = ""
                        Title = ""
                        DistinguishedName = $Device.DistinguishedName
                    }
                }
            }
        }
        
        # Get phone numbers from meeting rooms (if any)
        Write-Host "  Collecting phone numbers from meeting rooms..." -ForegroundColor Gray
        try {
            $MeetingRooms = Get-CsMeetingRoom -ErrorAction SilentlyContinue
            if ($MeetingRooms) {
                foreach ($Room in $MeetingRooms) {
                    if (![string]::IsNullOrEmpty($Room.LineURI)) {
                        $PhoneNumber = $Room.LineURI -replace "tel:", "" -replace ";.*", ""
                        $Extension = if ($Room.LineURI -match ";ext=(\d+)") { $matches[1] } else { "" }
                        
                        $AllPhoneNumbers += [PSCustomObject]@{
                            PhoneNumber = $PhoneNumber
                            Extension = $Extension
                            FullLineURI = $Room.LineURI
                            AssignedTo = $Room.DisplayName
                            AssignedType = "Meeting Room"
                            SIPAddress = $Room.SipAddress
                            UserPrincipalName = $Room.UserPrincipalName
                            Pool = $Room.RegistrarPool
                            VoicePolicy = $Room.VoicePolicy
                            VoiceEnabled = $Room.EnterpriseVoiceEnabled
                            Enabled = $Room.Enabled
                            Site = if ($Room.RegistrarPool) {
                                if ($Room.RegistrarPool -match "^(.+?)-" -or $Room.RegistrarPool -match "^(.+?)\.") {
                                    $matches[1]
                                } else {
                                    "Unknown"
                                }
                            } else { "Unknown" }
                            FirstName = ""
                            LastName = ""
                            Department = ""
                            Title = ""
                            DistinguishedName = $Room.DistinguishedName
                        }
                    }
                }
            }
        } catch {
            Write-Host "    ⚠️  Meeting rooms not available in this version" -ForegroundColor Gray
        }
        
        if ($AllPhoneNumbers.Count -gt 0) {
            # Sort by phone number for easier analysis
            $SortedPhoneNumbers = $AllPhoneNumbers | Sort-Object PhoneNumber
            
            $SortedPhoneNumbers | Export-Csv -Path $PhoneNumberFile -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Exported $($SortedPhoneNumbers.Count) phone numbers to: $PhoneNumberFile" -ForegroundColor Green
            
            # Display summary statistics
            $NumberSummary = $SortedPhoneNumbers | Group-Object AssignedType
            Write-Host "  📊 Phone Number Summary:" -ForegroundColor Cyan
            foreach ($Type in $NumberSummary) {
                Write-Host "    $($Type.Name): $($Type.Count) numbers" -ForegroundColor White
            }
            
            # Check for duplicate numbers
            $Duplicates = $SortedPhoneNumbers | Group-Object PhoneNumber | Where-Object { $_.Count -gt 1 }
            if ($Duplicates) {
                Write-Host "  ⚠️  Found $($Duplicates.Count) duplicate phone numbers!" -ForegroundColor Yellow
                $DuplicateFile = "$OutputDirectory\Lync_DuplicatePhoneNumbers_$Timestamp.csv"
                $DuplicateNumbers = @()
                foreach ($Dup in $Duplicates) {
                    foreach ($Item in $Dup.Group) {
                        $DuplicateNumbers += $Item | Select-Object *, @{Name='DuplicateCount'; Expression={$Dup.Count}}
                    }
                }
                $DuplicateNumbers | Export-Csv -Path $DuplicateFile -NoTypeInformation -Encoding UTF8
                Write-Host "    📋 Duplicate numbers exported to: $DuplicateFile" -ForegroundColor Yellow
            } else {
                Write-Host "  ✅ No duplicate phone numbers found" -ForegroundColor Green
            }
            
            # Show number ranges/patterns
            $NumberPatterns = $SortedPhoneNumbers | Group-Object { ($_.PhoneNumber -replace '[^0-9]', '').Substring(0, [Math]::Min(7, ($_.PhoneNumber -replace '[^0-9]', '').Length)) } | 
                             Sort-Object Name | Select-Object -First 10
            if ($NumberPatterns) {
                Write-Host "  📞 Top Number Patterns:" -ForegroundColor Cyan
                foreach ($Pattern in $NumberPatterns) {
                    Write-Host "    $($Pattern.Name)xxxx: $($Pattern.Count) numbers" -ForegroundColor White
                }
            }
            
        } else {
            Write-Host "⚠️  No phone numbers found in the environment" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "❌ Error exporting phone numbers: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to export site policy mappings
function Export-SitePolicyMappings {
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    try {
        Write-Host "Exporting Site Policy Mappings..." -ForegroundColor Yellow
        
        # Get all Lync sites
        $Sites = Get-CsSite -ErrorAction SilentlyContinue
        if ($Sites) {
            $SitePolicyFile = "$OutputDirectory\Lync_SitePolicyMappings_$Timestamp.csv"
            
            $SiteMappings = @()
            foreach ($Site in $Sites) {
                try {
                    # Get site-specific policies
                    $VoicePolicy = Get-CsVoicePolicy -Identity "site:$($Site.Identity)" -ErrorAction SilentlyContinue
                    $ConferencingPolicy = Get-CsConferencingPolicy -Identity "site:$($Site.Identity)" -ErrorAction SilentlyContinue
                    $LocationPolicy = Get-CsLocationPolicy -Identity "site:$($Site.Identity)" -ErrorAction SilentlyContinue
                    $ClientPolicy = Get-CsClientPolicy -Identity "site:$($Site.Identity)" -ErrorAction SilentlyContinue
                    
                    $SiteMappings += [PSCustomObject]@{
                        SiteIdentity = $Site.Identity
                        SiteName = $Site.DisplayName
                        SiteDescription = $Site.Description
                        SiteId = $Site.SiteId
                        VoicePolicy = if ($VoicePolicy) { $VoicePolicy.Identity } else { "Not Set (Uses Global)" }
                        ConferencingPolicy = if ($ConferencingPolicy) { $ConferencingPolicy.Identity } else { "Not Set (Uses Global)" }
                        LocationPolicy = if ($LocationPolicy) { $LocationPolicy.Identity } else { "Not Set (Uses Global)" }
                        ClientPolicy = if ($ClientPolicy) { $ClientPolicy.Identity } else { "Not Set (Uses Global)" }
                        Pools = ($Site.Pools -join '; ')
                        CentralSite = $Site.CentralSite
                    }
                } catch {
                    Write-Warning "Error processing site $($Site.Identity): $($_.Exception.Message)"
                }
            }
            
            if ($SiteMappings.Count -gt 0) {
                $SiteMappings | Export-Csv -Path $SitePolicyFile -NoTypeInformation -Encoding UTF8
                Write-Host "✅ Exported $($SiteMappings.Count) site policy mappings to: $SitePolicyFile" -ForegroundColor Green
            } else {
                Write-Host "⚠️  No site policy mappings found" -ForegroundColor Yellow
            }
        } else {
            Write-Host "⚠️  No Lync sites found" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "❌ Error exporting site policy mappings: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to export policies
function Export-LyncPolicies {
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    try {
        # Voice Policies
        Write-Host "Exporting Voice Policies..." -ForegroundColor Yellow
        $VoicePolicies = Get-CsVoicePolicy -ErrorAction SilentlyContinue
        if ($VoicePolicies) {
            $VoicePolicyFile = "$OutputDirectory\Lync_VoicePolicies_$Timestamp.csv"
            
            # Enhanced voice policy export with site information
            $EnhancedVoicePolicies = $VoicePolicies | Select-Object @{
                Name = 'Identity'; Expression = { $_.Identity }
            }, @{
                Name = 'Name'; Expression = { $_.Name }
            }, @{
                Name = 'Description'; Expression = { $_.Description }
            }, @{
                Name = 'Site'; Expression = { 
                    # Extract site information from Identity
                    if ($_.Identity -match "site:(.+?)/" -or $_.Identity -match "site:(.+)$") {
                        $matches[1]
                    } elseif ($_.Identity -like "site:*") {
                        ($_.Identity -replace "site:", "") -replace "/.*", ""
                    } elseif ($_.Identity -eq "Global") {
                        "Global"
                    } else {
                        # Try to determine from policy name patterns
                        if ($_.Name -match "^(.+?)[-_](Voice|Policy)" -and $_.Identity -ne "Global") {
                            $matches[1]
                        } else {
                            "Unknown"
                        }
                    }
                }
            }, @{
                Name = 'Scope'; Expression = { 
                    if ($_.Identity -eq "Global") { "Global" }
                    elseif ($_.Identity -like "site:*") { "Site" }
                    elseif ($_.Identity -like "tag:*") { "User/Tag" }
                    else { "Other" }
                }
            }, @{
                Name = 'EnableDelegation'; Expression = { $_.EnableDelegation }
            }, @{
                Name = 'EnableTeamCall'; Expression = { $_.EnableTeamCall }
            }, @{
                Name = 'EnableCallTransfer'; Expression = { $_.EnableCallTransfer }
            }, @{
                Name = 'EnableCallPark'; Expression = { $_.EnableCallPark }
            }, @{
                Name = 'EnableMaliciousCallTracing'; Expression = { $_.EnableMaliciousCallTracing }
            }, @{
                Name = 'EnableBWPolicyOverride'; Expression = { $_.EnableBWPolicyOverride }
            }, @{
                Name = 'PstnUsages'; Expression = { $_.PstnUsages -join '; ' }
            }, @{
                Name = 'AssignedUsers'; Expression = { 
                    # Count users assigned to this policy
                    try {
                        $PolicyUsers = Get-CsUser -VoicePolicy $_.Identity -ErrorAction SilentlyContinue
                        if ($PolicyUsers) {
                            $PolicyUsers.Count
                        } else {
                            0
                        }
                    } catch {
                        "Error"
                    }
                }
            }
            
            $EnhancedVoicePolicies | Export-Csv -Path $VoicePolicyFile -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Exported voice policies to: $VoicePolicyFile" -ForegroundColor Green
            
            # Display summary by scope/site
            $PolicySummary = $EnhancedVoicePolicies | Group-Object Scope
            Write-Host "  📊 Voice Policy Summary:" -ForegroundColor Cyan
            foreach ($Scope in $PolicySummary) {
                Write-Host "    $($Scope.Name): $($Scope.Count) policies" -ForegroundColor White
            }
        }
        
        # Conferencing Policies
        Write-Host "Exporting Conferencing Policies..." -ForegroundColor Yellow
        $ConferencingPolicies = Get-CsConferencingPolicy -ErrorAction SilentlyContinue
        if ($ConferencingPolicies) {
            $ConferencingPolicyFile = "$OutputDirectory\Lync_ConferencingPolicies_$Timestamp.csv"
            $ConferencingPolicies | Select-Object Identity, MaxMeetingSize, AllowIPAudio, AllowIPVideo,
                AllowMultiView, MaxVideoConferenceResolution, AllowLargeMeetings, EnableDialInConferencing,
                AllowExternalUsersToSaveContent, AllowExternalUserControl | Export-Csv -Path $ConferencingPolicyFile -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Exported conferencing policies to: $ConferencingPolicyFile" -ForegroundColor Green
        }
        
    } catch {
        Write-Host "❌ Error exporting policies: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main menu
Clear-Host
Write-Host $Separator -ForegroundColor Cyan
Write-Host "$OrganizationName - LYNC CSV EXPORT TOOL" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Host ""
Write-Host "Output Directory: $OutputDirectory" -ForegroundColor Yellow
Write-Host ""

do {
    Write-Host "Select export option:" -ForegroundColor White
    Write-Host ""
    Write-Host "USER EXPORTS:" -ForegroundColor Cyan
    Write-Host "  1. Users Summary (Basic info)" -ForegroundColor White
    Write-Host "  2. Voice Users Only" -ForegroundColor White  
    Write-Host "  3. SBA Users (School sites)" -ForegroundColor White
    Write-Host "  4. Complete User Export (All fields)" -ForegroundColor White
    Write-Host ""
    Write-Host "PHONE INVENTORY:" -ForegroundColor Magenta
    Write-Host "  5. Common Area Phones" -ForegroundColor White
    Write-Host "  6. Analog Devices (Fax machines, etc.)" -ForegroundColor White
    Write-Host "  7. IP Phones (Desk phones, conference room)" -ForegroundColor White
    Write-Host "  8. USB Communication Devices (Headsets, etc.)" -ForegroundColor White
    Write-Host "  9. All Phone/Device Inventory" -ForegroundColor White
    Write-Host ""
    Write-Host "INFRASTRUCTURE EXPORTS:" -ForegroundColor Cyan
    Write-Host " 10. Pool Information" -ForegroundColor White
    Write-Host " 11. Policy Information" -ForegroundColor White
    Write-Host " 12. Site Policy Mappings" -ForegroundColor White
    Write-Host " 13. All Phone Numbers (LineURIs)" -ForegroundColor White
    Write-Host ""
    Write-Host "BULK EXPORTS:" -ForegroundColor Cyan
    Write-Host " 14. Export All (Users + Phones + Pools + Policies + Numbers)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Q. Quit" -ForegroundColor Red
    Write-Host ""
    
    $Choice = Read-Host "Enter your choice"
    Write-Host ""
    
    switch ($Choice.ToUpper()) {
        "1" { Export-LyncUsers -Type "Summary" }
        "2" { Export-LyncUsers -Type "Voice" }
        "3" { Export-LyncUsers -Type "SBA" }
        "4" { Export-LyncUsers -Type "Complete" }
        "5" { Export-PhoneInventory -Type "CommonArea" }
        "6" { Export-PhoneInventory -Type "Analog" }
        "7" { Export-PhoneInventory -Type "IPPhones" }
        "8" { Export-PhoneInventory -Type "USB" }
        "9" { Export-PhoneInventory -Type "All" }
        "10" { Export-LyncPools }
        "11" { Export-LyncPolicies }
        "12" { Export-SitePolicyMappings }
        "13" { Export-PhoneNumbers }
        "14" { 
            Write-Host "Exporting all data..." -ForegroundColor Cyan
            Export-LyncUsers -Type "Complete"
            Export-PhoneInventory -Type "All"
            Export-LyncPools
            Export-LyncPolicies
            Export-SitePolicyMappings
            Export-PhoneNumbers
            Write-Host "All exports completed!" -ForegroundColor Green
        }
        "Q" { 
            Write-Host "Exiting..." -ForegroundColor Yellow
            break 
        }
        default { 
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red 
        }
    }
    
    if ($Choice.ToUpper() -ne "Q") {
        Write-Host ""
        Write-Host "Press Enter to continue..." -ForegroundColor Gray
        Read-Host
        Write-Host ""
    }
    
} while ($Choice.ToUpper() -ne "Q")

Write-Host ""
Write-Host "CSV exports saved to: $OutputDirectory" -ForegroundColor Green
Write-Host "You can now import these files into Excel, Power BI, or other analysis tools." -ForegroundColor White