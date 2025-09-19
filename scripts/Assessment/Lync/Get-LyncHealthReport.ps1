<#
.SYNOPSIS
    Generates a comprehensive Lync/Skype for Business health and diagnostics report.

.DESCRIPTION
    This script performs health monitoring and diagnostic analysis of a Lync/Skype for Business environment.
    It checks certificate status, database mirror states, health monitoring configuration, recent event logs,
    and system performance metrics to provide administrators with a complete health overview.

.PARAMETER PoolFQDN
    The fully qualified domain name of the Lync pool to analyze for database mirror state and health checks.
    This parameter is mandatory as it's required for pool-specific health analysis.

.PARAMETER ReportPath
    The file path where the health diagnostics report will be saved. If not specified, defaults to
    "C:\Reports\Lync_Health_Diagnostics_[timestamp].txt".

.PARAMETER EventLogHours
    The number of hours to look back when analyzing event log errors. Defaults to 24 hours.
    This controls the timeframe for recent error analysis.

.PARAMETER MaxEventLogErrors
    The maximum number of event log errors to retrieve and analyze. Defaults to 20 errors.
    This helps limit the report size while capturing the most recent issues.

.PARAMETER OrganizationName
    The name of the organization for which the report is being generated. Used in the report header
    and provides organizational context throughout the health analysis.

.EXAMPLE
    .\Get-LyncHealthReport.ps1 -PoolFQDN "lyncpool.contoso.com"
    
    Generates a health report for the specified Lync pool using default settings.

.EXAMPLE
    .\Get-LyncHealthReport.ps1 -PoolFQDN "lyncpool.contoso.com" -OrganizationName "Contoso Corp" -EventLogHours 48
    
    Generates a health report with custom organization name and extended event log analysis period.

.EXAMPLE
    .\Get-LyncHealthReport.ps1 -PoolFQDN "lyncpool.contoso.com" -ReportPath "D:\Reports\Health_Report.txt" -MaxEventLogErrors 50
    
    Generates a health report with custom output path and increased event log error limit.

.NOTES
    Author: W. Ford
    Date: 2025-09-17
    Version: 2.0
    
    Requirements:
    - Lync/Skype for Business Management Shell
    - Administrative privileges on Lync infrastructure
    - PowerShell 3.0 or higher
    - Access to Windows Event Logs
    
    The script performs the following health checks:
    - Certificate expiration and validity analysis
    - Database mirror state assessment
    - Health monitoring configuration validation
    - Recent event log error analysis
    - System performance metrics collection
    - Lync-specific performance counters (if available)

.LINK
    https://docs.microsoft.com/en-us/skypeforbusiness/manage/health-and-monitoring/

#>

param(
    [Parameter(Mandatory=$true, HelpMessage="The FQDN of the Lync pool to analyze")]
    [ValidateNotNullOrEmpty()]
    [string]$PoolFQDN,

    [Parameter(Mandatory=$false, HelpMessage="Path where the health report will be saved")]
    [string]$ReportPath = "C:\Reports\Lync_Health_Diagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",
    
    [Parameter(Mandatory=$false, HelpMessage="Hours to look back for event log analysis")]
    [ValidateRange(1, 168)] # 1 hour to 1 week
    [int]$EventLogHours = 24,
    
    [Parameter(Mandatory=$false, HelpMessage="Maximum number of event log errors to retrieve")]
    [ValidateRange(1, 100)]
    [int]$MaxEventLogErrors = 20,
    
    [Parameter(Mandatory=$false, HelpMessage="Organization name for the report header")]
    [string]$OrganizationName = "Organization"
)

# Use provided or default report path if not specified directly
$Separator = "=" * 80

# Create reports directory if it doesn't exist
$ReportsDir = Split-Path $ReportPath -Parent
if (!(Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force
}

# Start report
$Report = @()
$Report += "$OrganizationName - LYNC HEALTH AND DIAGNOSTICS REPORT"
$Report += "Generated: $(Get-Date)"
$Report += "Server: $($env:COMPUTERNAME)"
$Report += $Separator
$Report += ""

# Section 1: Certificate Information
$Report += "CERTIFICATE INFORMATION"
$Report += $Separator
try {
    $Certificates = Get-CsCertificate -ErrorAction SilentlyContinue
    if ($Certificates) {
        $Certificates | ForEach-Object {
            $Report += "  Use: $($_.Use)"
            $Report += "  Thumbprint: $($_.Thumbprint)"
            $Report += "  Subject: $($_.Subject)"
            $Report += "  Issuer: $($_.Issuer)"
            $Report += "  Not Before: $($_.NotBefore)"
            $Report += "  Not After: $($_.NotAfter)"
            $DaysToExpiry = ($_.NotAfter - (Get-Date)).Days
            $Report += "  Days to Expiry: $DaysToExpiry"
            if ($DaysToExpiry -lt 30) {
                $Report += "  WARNING: Certificate expires in less than 30 days!"
            }
            $Report += ""
        }
    } else {
        $Report += "No certificate information available."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving certificate information: $($_.Exception.Message)"
    $Report += ""
}

# Section 2: Database Mirror State
$Report += "DATABASE MIRROR STATE"
$Report += $Separator
try {
    $MirrorState = Get-CsDatabaseMirrorState -PoolFqdn $PoolFQDN -ErrorAction SilentlyContinue
    if ($MirrorState) {
        $Report += "DATABASE MIRROR ANALYSIS:"
        $Report += $Separator
        
        $MirrorState | ForEach-Object {
            $Report += "Database: $($_.DatabaseName)"
            
            # Handle empty/null values properly
            $Principal = if ([string]::IsNullOrEmpty($_.Principal)) { "Not configured" } else { $_.Principal }
            $Mirror = if ([string]::IsNullOrEmpty($_.Mirror)) { "Not configured" } else { $_.Mirror }
            $State = if ([string]::IsNullOrEmpty($_.State)) { "Not mirrored" } else { $_.State }
            $Witness = if ([string]::IsNullOrEmpty($_.Witness)) { "Not configured" } else { $_.Witness }
            
            $Report += "  Principal Server: $Principal"
            $Report += "  Mirror Server: $Mirror" 
            $Report += "  Mirror State: $State"
            $Report += "  Witness Server: $Witness"
            
            # Determine mirror status
            if ($State -eq "Not mirrored" -or [string]::IsNullOrEmpty($_.State)) {
                $Report += "  Status: ⚠️  Database not mirrored (Single server deployment)"
            } elseif ($State -like "*Synchronized*") {
                $Report += "  Status: ✅ Mirror synchronized"
            } elseif ($State -like "*Disconnected*") {
                $Report += "  Status: ❌ Mirror disconnected - requires attention"
            } else {
                $Report += "  Status: 🔄 $State"
            }
            $Report += ""
        }
        
        # Summary of mirror configuration
        $MirroredDatabases = $MirrorState | Where-Object { -not [string]::IsNullOrEmpty($_.State) -and $_.State -ne "Not mirrored" }
        $UnmirroredDatabases = $MirrorState | Where-Object { [string]::IsNullOrEmpty($_.State) -or $_.State -eq "Not mirrored" }
        
        $Report += "MIRROR CONFIGURATION SUMMARY:"
        $Report += "  Total Databases: $($MirrorState.Count)"
        $Report += "  Mirrored Databases: $($MirroredDatabases.Count)"
        $Report += "  Non-Mirrored Databases: $($UnmirroredDatabases.Count)"
        $Report += ""
        
        if ($UnmirroredDatabases.Count -gt 0) {
            $Report += "NON-MIRRORED DATABASES:"
            $UnmirroredDatabases | ForEach-Object {
                $Report += "  - $($_.DatabaseName)"
            }
            $Report += ""
            $Report += "NOTE: Non-mirrored databases are common in smaller deployments"
            $Report += "or development environments. Consider implementing database"
            $Report += "mirroring for high availability in production environments."
            $Report += ""
        }
        
    } else {
        $Report += "No database mirror information available."
        $Report += "This may indicate:"
        $Report += "- Single server deployment without mirroring"
        $Report += "- Insufficient permissions to query mirror state"
        $Report += "- Database mirroring not configured"
        $Report += ""
    }
} catch {
    $Report += "Error retrieving database mirror state: $($_.Exception.Message)"
    $Report += ""
    
    # Try alternative approach to get database information
    try {
        $Report += "ALTERNATIVE DATABASE CHECK:"
        # Try to get SQL Server information if available
        $SqlServices = Get-Service -Name "*SQL*" -ErrorAction SilentlyContinue
        if ($SqlServices) {
            $Report += "SQL Server services found:"
            $SqlServices | ForEach-Object {
                $Report += "  Service: $($_.Name) - Status: $($_.Status)"
            }
        } else {
            $Report += "No SQL Server services found on this server"
        }
        $Report += ""
    } catch {
        $Report += "Could not retrieve alternative database information"
        $Report += ""
    }
}

# Section 3: Health Monitoring Configuration
$Report += "HEALTH MONITORING CONFIGURATION"
$Report += $Separator
try {
    $HealthConfig = Get-CsHealthMonitoringConfiguration -ErrorAction SilentlyContinue
    if ($HealthConfig) {
        $HealthConfig | ForEach-Object {
            $Report += "  Identity: $($_.Identity)"
            $Report += "  Target FQDN: $($_.TargetFqdn)"
            $Report += "  First Test User SIP: $($_.FirstTestUserSipUri)"
            $Report += "  Second Test User SIP: $($_.SecondTestUserSipUri)"
            $Report += ""
        }
    } else {
        $Report += "No health monitoring configuration found."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving health monitoring configuration: $($_.Exception.Message)"
    $Report += ""
}

# Section 4: Event Log Errors (Last $EventLogHours hours)
$Report += "RECENT EVENT LOG ERRORS (Last $EventLogHours Hours)"
$Report += $Separator
try {
    $StartTime = (Get-Date).AddHours(-$EventLogHours)
    
    # Try to get Lync Server event logs
    $LyncErrors = Get-WinEvent -FilterHashtable @{
        LogName = "Lync Server"
        Level = 2
        StartTime = $StartTime
    } -MaxEvents $MaxEventLogErrors -ErrorAction SilentlyContinue
    
    if ($LyncErrors) {
        $Report += "  Lync Server Log Errors:"
        $LyncErrors | ForEach-Object {
            $Report += "    Time: $($_.TimeCreated)"
            $Report += "    Event ID: $($_.Id)"
            $Report += "    Level: $($_.LevelDisplayName)"
            $Report += "    Message: $($_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)))..."
            $Report += ""
        }
    } else {
        $Report += "  No Lync Server errors found in the last $EventLogHours hours."
    }
    
    # Try to get Application log errors related to Lync/RTC
    $AppErrors = Get-WinEvent -FilterHashtable @{
        LogName = "Application"
        Level = 2
        StartTime = $StartTime
    } -MaxEvents 50 -ErrorAction SilentlyContinue | Where-Object { 
        $_.ProviderName -like "*RTC*" -or $_.ProviderName -like "*Lync*" 
    } | Select-Object -First 10
    
    if ($AppErrors) {
        $Report += "  Application Log RTC/Lync Errors:"
        $AppErrors | ForEach-Object {
            $Report += "    Time: $($_.TimeCreated)"
            $Report += "    Source: $($_.ProviderName)"
            $Report += "    Event ID: $($_.Id)"
            $Report += "    Message: $($_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)))..."
            $Report += ""
        }
    } else {
        $Report += "  No RTC/Lync related application errors found."
    }
    
} catch {
    $Report += "Error retrieving event logs: $($_.Exception.Message)"
    $Report += ""
}

# Section 5: System Performance Counters
$Report += "SYSTEM PERFORMANCE METRICS"
$Report += $Separator
try {
    # Get basic system metrics
    $CPU = Get-Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 3 | 
           Select-Object -ExpandProperty CounterSamples | 
           Measure-Object -Property CookedValue -Average
    
    $Memory = Get-Counter "\Memory\Available MBytes"
    $Disk = Get-Counter "\PhysicalDisk(_Total)\% Disk Time"
    
    $Report += "  Average CPU Usage: $([Math]::Round($CPU.Average, 2))%"
    $Report += "  Available Memory: $([Math]::Round($Memory.CounterSamples[0].CookedValue, 0)) MB"
    $Report += "  Disk Usage: $([Math]::Round($Disk.CounterSamples[0].CookedValue, 2))%"
    $Report += ""
    
    # Try to get Lync-specific counters if available
    try {
        $LyncCounters = Get-Counter -ListSet "*LS:*" -ErrorAction SilentlyContinue
        if ($LyncCounters) {
            $Report += "  Lync Performance Counters Available: $($LyncCounters.Count)"
        } else {
            $Report += "  No Lync-specific performance counters found."
        }
    } catch {
        $Report += "  Lync performance counters not accessible."
    }
    
} catch {
    $Report += "Error retrieving performance metrics: $($_.Exception.Message)"
    $Report += ""
}

# Export report
$Report | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "Health and Diagnostics Report exported to: $ReportPath" -ForegroundColor Green