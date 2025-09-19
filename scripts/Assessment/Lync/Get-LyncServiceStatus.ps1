<#
.SYNOPSIS
    Generates a comprehensive Lync/Skype for Business service status report.

.DESCRIPTION
    This script analyzes and reports on the status of Lync/Skype for Business services across the environment.
    It checks Windows services, Lync Management Shell services, and related processes to provide
    administrators with a complete view of service health and operational status.

.PARAMETER OrganizationName
    The name of the organization for which the report is being generated. Used in the report header
    and provides organizational context throughout the service analysis.

.PARAMETER ReportPath
    The file path where the service status report will be saved. If not specified, defaults to
    "C:\Reports\Lync_Service_Status_[timestamp].txt".

.PARAMETER ServicePatterns
    An array of patterns used to identify Lync/Skype for Business related services. Defaults to
    @("*RTC*", "*Lync*", "*Skype*"). These patterns are used to filter and identify relevant services.

.PARAMETER SpecificServices
    An array of specific service names to check individually. Defaults to common Lync service names:
    @("RTCSRV", "RTCCLSAGT", "RTCATS", "RTCDSS", "RTCMCU", "RTCASMCU"). These services are checked
    for detailed status reporting regardless of whether they match the service patterns.

.EXAMPLE
    .\Get-LyncServiceStatus.ps1
    
    Generates a service status report using default organization name and service patterns.

.EXAMPLE
    .\Get-LyncServiceStatus.ps1 -OrganizationName "Contoso Corp" -ReportPath "D:\Reports\Contoso_Services.txt"
    
    Generates a service status report with custom organization name and output path.

.EXAMPLE
    .\Get-LyncServiceStatus.ps1 -ServicePatterns @("*Teams*", "*SfB*", "*RTC*") -SpecificServices @("TeamsService", "SfBService")
    
    Generates a report with custom service patterns and specific service names for Teams/SfB environment.

.NOTES
    Author: W. Ford
    Date: 2025-09-17
    Version: 2.0
    
    Requirements:
    - Administrative privileges to query services
    - PowerShell 3.0 or higher
    - Lync/Skype for Business Management Shell (optional, for enhanced reporting)
    
    The script analyzes the following service aspects:
    - Windows Services status and configuration
    - Lync Management Shell service information (if available)
    - Related process information including resource usage
    - Service dependency and startup configuration
    - Process performance metrics (CPU, Memory usage)

.LINK
    https://docs.microsoft.com/en-us/skypeforbusiness/manage/services/

#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Organization name for the report header")]
    [string]$OrganizationName = "Organization",
    
    [Parameter(Mandatory=$false, HelpMessage="Path where the service status report will be saved")]
    [string]$ReportPath = "C:\Reports\Lync_Service_Status_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",
    
    [Parameter(Mandatory=$false, HelpMessage="Patterns to identify Lync/SfB services")]
    [string[]]$ServicePatterns = @("*RTC*", "*Lync*", "*Skype*"),
    
    [Parameter(Mandatory=$false, HelpMessage="Specific service names to check individually")]
    [string[]]$SpecificServices = @("RTCSRV", "RTCCLSAGT", "RTCATS", "RTCDSS", "RTCMCU", "RTCASMCU")
)
$Separator = "=" * 80

# Create reports directory if it doesn't exist
$ReportsDir = Split-Path $ReportPath -Parent
if (!(Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force
}

# Start report
$Report = @()
$Report += "$OrganizationName - LYNC SERVICE STATUS REPORT"
$Report += "Generated: $(Get-Date)"
$Report += "Server: $($env:COMPUTERNAME)"
$Report += $Separator
$Report += ""

# Section 1: Windows Services Status
$Report += "WINDOWS SERVICES STATUS"
$Report += $Separator
try {
    # Get services matching any of the specified patterns
    $LyncServices = @()
    foreach ($Pattern in $ServicePatterns) {
        $LyncServices += Get-Service -Name $Pattern -ErrorAction SilentlyContinue
    }
    
    if ($LyncServices) {
        $Report += "Lync/Skype for Business Services:"
        $LyncServices | Sort-Object Name | ForEach-Object {
            $Report += "  Service: $($_.Name)"
            $Report += "  Display Name: $($_.DisplayName)"
            $Report += "  Status: $($_.Status)"
            $Report += "  Start Type: $($_.StartType)"
            $Report += ""
        }
    } else {
        $Report += "No matching services found on this server."
        $Report += ""
    }
    
    # Check specific Lync services
    $Report += "Specific Lync Services Check:"
    foreach ($ServiceName in $SpecificServices) {
        try {
            $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if ($Service) {
                $Report += "  $ServiceName - Status: $($Service.Status) | Start Type: $($Service.StartType)"
            } else {
                $Report += "  $ServiceName - Not Found/Not Installed"
            }
        } catch {
            $Report += "  $ServiceName - Error: $($_.Exception.Message)"
        }
    }
} catch {
    $Report += "Error retrieving Windows services: $($_.Exception.Message)"
}
$Report += ""

# Section 2: Lync Windows Services (if Lync cmdlets available)
$Report += "LYNC MANAGEMENT SHELL SERVICES"
$Report += $Separator
try {
    $LyncWindowsServices = Get-CsWindowsService -ErrorAction SilentlyContinue
    if ($LyncWindowsServices) {
        $LyncWindowsServices | ForEach-Object {
            $Report += "  Computer: $($_.Computer)"
            $Report += "  Service: $($_.Name)"
            $Report += "  Status: $($_.Status)"
            $Report += "  Activity: $($_.Activity)"
            $Report += ""
        }
    } else {
        $Report += "Lync Management Shell not available or no services found."
        $Report += ""
    }
} catch {
    $Report += "Lync Management Shell not available: $($_.Exception.Message)"
    $Report += ""
}

# Section 3: Process Information
$Report += "LYNC PROCESSES"
$Report += $Separator
try {
    # Build process filter dynamically from service patterns
    $ProcessFilter = {
        $ProcessName = $_.ProcessName
        $ServicePatterns | ForEach-Object { 
            if ($ProcessName -like $_.Replace("*", "*")) { return $true }
        }
        return $false
    }
    
    $LyncProcesses = Get-Process | Where-Object $ProcessFilter
    if ($LyncProcesses) {
        $LyncProcesses | ForEach-Object {
            $Report += "  Process: $($_.ProcessName)"
            $Report += "  ID: $($_.Id)"
            $Report += "  CPU: $($_.CPU)"
            $Report += "  Memory (MB): $([math]::Round($_.WorkingSet64/1MB, 2))"
            $Report += "  Start Time: $($_.StartTime)"
            $Report += ""
        }
    } else {
        $Report += "No Lync-related processes found."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving process information: $($_.Exception.Message)"
    $Report += ""
}

# Export report
$Report | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "Service Status Report exported to: $ReportPath" -ForegroundColor Green