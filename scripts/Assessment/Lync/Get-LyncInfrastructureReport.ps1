<#
.SYNOPSIS
    Generates a comprehensive Lync/Skype for Business infrastructure and configuration report.

.DESCRIPTION
    This script analyzes and reports on the infrastructure components of a Lync/Skype for Business environment.
    It provides detailed information about pools, computers, services, topology, and conference directories,
    along with categorization by pool type and health status assessment.

.PARAMETER OrganizationName
    The name of the organization for which the report is being generated. Used in the report header
    and provides organizational context throughout the infrastructure analysis.

.PARAMETER ReportPath
    The file path where the infrastructure report will be saved. If not specified, defaults to
    "C:\Reports\Lync_Infrastructure_[timestamp].txt".

.PARAMETER SBAPattern
    The pattern used to identify Survivable Branch Appliances (SBA) pools. Defaults to "*MSSBA*".
    This pattern helps categorize and analyze branch office infrastructure deployments.

.PARAMETER IVRPattern
    The pattern used to identify Interactive Voice Response (IVR) pools. Defaults to "*ivr*".
    Used for categorizing voice infrastructure components.

.PARAMETER EdgePattern
    The pattern used to identify Edge server pools. Defaults to "*edge*".
    Important for external access and federation infrastructure analysis.

.PARAMETER LyncPattern
    The pattern used to identify standard Lync/Skype for Business pools. Defaults to "*lync*".
    Used to categorize core communication infrastructure.

.PARAMETER MaxComputersPerPool
    The maximum number of computers to display per pool in the detailed analysis. Defaults to 5.
    This helps keep the report manageable while showing key infrastructure components.

.PARAMETER MaxServicesPerRole
    The maximum number of services to display per role in the detailed analysis. Defaults to 5.
    Controls the verbosity of service configuration reporting.

.EXAMPLE
    .\Get-LyncInfrastructureReport.ps1
    
    Generates an infrastructure report using default organization name and all default patterns.

.EXAMPLE
    .\Get-LyncInfrastructureReport.ps1 -OrganizationName "Contoso Corp" -ReportPath "D:\Reports\Contoso_Infrastructure.txt"
    
    Generates an infrastructure report with custom organization name and output path.

.EXAMPLE
    .\Get-LyncInfrastructureReport.ps1 -SBAPattern "*Branch*" -LyncPattern "*teams*" -MaxComputersPerPool 10
    
    Generates a report with custom pool patterns and increased computer display limit.

.NOTES
    Author: W. Ford
    Date: 2025-09-17
    Version: 2.0
    
    Requirements:
    - Lync/Skype for Business Management Shell
    - Administrative privileges on Lync infrastructure
    - PowerShell 3.0 or higher
    
    The script analyzes the following infrastructure components:
    - Pool categorization and detailed configuration
    - Computer deployment across pools
    - Service configuration and role distribution
    - Topology structure and site information
    - Conference directory configuration
    - Infrastructure health summary with status indicators

.LINK
    https://docs.microsoft.com/en-us/skypeforbusiness/plan-your-deployment/

#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Organization name for the report header")]
    [string]$OrganizationName = "Organization",
    
    [Parameter(Mandatory=$false, HelpMessage="Path where the infrastructure report will be saved")]
    [string]$ReportPath = "C:\Reports\Lync_Infrastructure_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",
    
    [Parameter(Mandatory=$false, HelpMessage="Pattern to identify SBA pools")]
    [string]$SBAPattern = "*MSSBA*",
    
    [Parameter(Mandatory=$false, HelpMessage="Pattern to identify IVR pools")]
    [string]$IVRPattern = "*ivr*",
    
    [Parameter(Mandatory=$false, HelpMessage="Pattern to identify Edge pools")]
    [string]$EdgePattern = "*edge*",
    
    [Parameter(Mandatory=$false, HelpMessage="Pattern to identify Lync pools")]
    [string]$LyncPattern = "*lync*",
    
    [Parameter(Mandatory=$false, HelpMessage="Maximum computers to display per pool")]
    [ValidateRange(1, 50)]
    [int]$MaxComputersPerPool = 5,
    
    [Parameter(Mandatory=$false, HelpMessage="Maximum services to display per role")]
    [ValidateRange(1, 50)]
    [int]$MaxServicesPerRole = 5
)
$Separator = "=" * 80

# Create reports directory if it doesn't exist
$ReportsDir = Split-Path $ReportPath -Parent
if (!(Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force
}

# Start report
$Report = @()
$Report += "$OrganizationName - LYNC INFRASTRUCTURE REPORT"
$Report += "Generated: $(Get-Date)"
$Report += "Server: $($env:COMPUTERNAME)"
$Report += $Separator
$Report += ""

# Executive Summary
$Report += "EXECUTIVE SUMMARY"
$Report += $Separator
try {
    $AllPools = Get-CsPool -ErrorAction SilentlyContinue
    $AllServices = Get-CsService -ErrorAction SilentlyContinue
    $AllComputers = Get-CsComputer -ErrorAction SilentlyContinue
    
    $Report += "Infrastructure Overview:"
    $Report += "  Total Pools: $($AllPools.Count)"
    $Report += "  Total Services: $($AllServices.Count)"
    $Report += "  Total Computers: $($AllComputers.Count)"
    $Report += ""
    
    if ($AllPools) {
        # Categorize pools
        $SBAPools = $AllPools | Where-Object { $_.Fqdn -like $SBAPattern }
        $StandardPools = $AllPools | Where-Object { $_.Fqdn -like $LyncPattern }
        $IVRPools = $AllPools | Where-Object { $_.Fqdn -like $IVRPattern }
        $EdgePools = $AllPools | Where-Object { $_.Fqdn -like $EdgePattern }
        
        $Report += "Pool Categorization:"
        $Report += "  Standard Lync Pools: $($StandardPools.Count)"
        $Report += "  Survivable Branch Appliances: $($SBAPools.Count)"
        $Report += "  IVR Pools: $($IVRPools.Count)"
        $Report += "  Edge Pools: $($EdgePools.Count)"
        $Report += ""
    }
} catch {
    $Report += "Error generating executive summary: $($_.Exception.Message)"
    $Report += ""
}

# Section 1: Pool Information (Enhanced)
$Report += "POOL INFORMATION"
$Report += $Separator
try {
    $Pools = Get-CsPool -ErrorAction SilentlyContinue
    if ($Pools) {
        $Report += "Total Pools Found: $($Pools.Count)"
        $Report += ""
        
        # Group and display by pool type
        $PoolsByType = @{
            "Standard Lync Pools" = $Pools | Where-Object { $_.Fqdn -like $LyncPattern -and $_.Fqdn -notlike $SBAPattern }
            "Survivable Branch Appliances (SBA)" = $Pools | Where-Object { $_.Fqdn -like $SBAPattern }
            "IVR Pools" = $Pools | Where-Object { $_.Fqdn -like $IVRPattern }
            "Edge Pools" = $Pools | Where-Object { $_.Fqdn -like $EdgePattern }
            "SQL/Witness Pools" = $Pools | Where-Object { $_.Fqdn -like "*sql*" -or $_.Fqdn -like "*witness*" }
            "Office Web Apps" = $Pools | Where-Object { $_.Fqdn -like "*office*" }
            "IP Address Pools" = $Pools | Where-Object { $_.Fqdn -match '^\d+\.\d+\.\d+\.\d+' }
            "Other Pools" = $Pools | Where-Object { 
                $_.Fqdn -notlike $LyncPattern -and $_.Fqdn -notlike $SBAPattern -and 
                $_.Fqdn -notlike $IVRPattern -and $_.Fqdn -notlike $EdgePattern -and
                $_.Fqdn -notlike "*sql*" -and $_.Fqdn -notlike "*witness*" -and
                $_.Fqdn -notlike "*office*" -and $_.Fqdn -notmatch '^\d+\.\d+\.\d+\.\d+'
            }
        }
        
        foreach ($PoolType in $PoolsByType.Keys) {
            $TypePools = $PoolsByType[$PoolType]
            if ($TypePools.Count -gt 0) {
                $Report += "$PoolType ($($TypePools.Count) pools):"
                $Report += $("-" * 40)
                
                $TypePools | Select-Object -First 10 | ForEach-Object {
                    $Report += "  Pool: $($_.Fqdn)"
                    $Report += "    Identity: $($_.Identity)"
                    $Report += "    Site: $($_.Site)"
                    if ($_.Services) {
                        $Report += "    Services: $($_.Services -join ', ')"
                    }
                    if ($_.Computers) {
                        $Report += "    Computers: $($_.Computers -join ', ')"
                    }
                    if ($_.BackupPoolFqdn) {
                        $Report += "    Backup Pool: $($_.BackupPoolFqdn)"
                    }
                    $Report += ""
                }
                
                if ($TypePools.Count -gt 10) {
                    $Report += "  ... and $($TypePools.Count - 10) more $PoolType"
                    $Report += ""
                }
            }
        }
    } else {
        $Report += "No pools found or Lync Management Shell not available."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving pool information: $($_.Exception.Message)"
    $Report += ""
}

# Section 2: Computer Information (Enhanced)
$Report += "COMPUTER INFORMATION"
$Report += $Separator
try {
    $Computers = Get-CsComputer -ErrorAction SilentlyContinue
    if ($Computers) {
        $Report += "Total Computers Found: $($Computers.Count)"
        $Report += ""
        
        # Group by pool for better organization
        $ComputersByPool = $Computers | Group-Object Pool | Sort-Object Count -Descending
        
        $Report += "COMPUTERS BY POOL (Top 10):"
        $ComputersByPool | Select-Object -First 10 | ForEach-Object {
            $Report += "  Pool: $($_.Name)"
            $Report += "  Computer Count: $($_.Count)"
            
            # Show first few computers
            $_.Group | Select-Object -First $MaxComputersPerPool | ForEach-Object {
                $Report += "    • Computer: $($_.Fqdn)"
                if ($_.Site) { $Report += "      Site: $($_.Site)" }
            }
            
            if ($_.Count -gt $MaxComputersPerPool) {
                $Report += "    ... and $($_.Count - $MaxComputersPerPool) more computers"
            }
            $Report += ""
        }
        
        if ($ComputersByPool.Count -gt 10) {
            $Report += "... and $($ComputersByPool.Count - 10) more pools with computers"
            $Report += ""
        }
        
    } else {
        $Report += "No computer information available."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving computer information: $($_.Exception.Message)"
    $Report += ""
}

# Section 3: Service Configuration (Enhanced)
$Report += "SERVICE CONFIGURATION"
$Report += $Separator
try {
    $Services = Get-CsService -ErrorAction SilentlyContinue
    if ($Services) {
        $Report += "Total Services Found: $($Services.Count)"
        $Report += ""
        
        # Group by service role
        $ServicesByRole = $Services | Group-Object Role | Sort-Object Name
        
        $Report += "SERVICES BY ROLE:"
        $ServicesByRole | ForEach-Object {
            $Report += "  Role: $($_.Name) ($($_.Count) services)"
            
            $_.Group | Select-Object -First $MaxServicesPerRole | ForEach-Object {
                $Report += "    • Service: $($_.Identity)"
                if ($_.PoolFqdn) { $Report += "      Pool: $($_.PoolFqdn)" }
                if ($_.SiteId) { $Report += "      Site: $($_.SiteId)" }
                if ($_.DependsOn) { $Report += "      Dependencies: $($_.DependsOn -join ', ')" }
            }
            
            if ($_.Count -gt $MaxServicesPerRole) {
                $Report += "    ... and $($_.Count - $MaxServicesPerRole) more $($_.Name) services"
            }
            $Report += ""
        }
    } else {
        $Report += "No service configuration available."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving service configuration: $($_.Exception.Message)"
    $Report += ""
}

# Section 4: Topology Information (Enhanced)
$Report += "TOPOLOGY SUMMARY"
$Report += $Separator
try {
    $Topology = Get-CsTopology -ErrorAction SilentlyContinue
    if ($Topology) {
        $Report += "✅ Topology retrieved successfully"
        
        if ($Topology.Sites) {
            $Report += "Total Sites: $($Topology.Sites.Count)"
            $Report += ""
            
            $Report += "SITE DETAILS:"
            $Topology.Sites | ForEach-Object {
                $Report += "  Site: $($_.Name)"
                if ($_.SiteId) { $Report += "    Site ID: $($_.SiteId)" }
                if ($_.Kind) { $Report += "    Type: $($_.Kind)" }
                if ($_.Description) { $Report += "    Description: $($_.Description)" }
                
                # Try to get pools in this site
                if ($AllPools) {
                    $SitePools = $AllPools | Where-Object { $_.Site -like "*$($_.Name)*" }
                    if ($SitePools) {
                        $Report += "    Pools in Site: $($SitePools.Count)"
                        $SitePools | Select-Object -First 3 | ForEach-Object {
                            $Report += "      • $($_.Fqdn)"
                        }
                        if ($SitePools.Count -gt 3) {
                            $Report += "      ... and $($SitePools.Count - 3) more"
                        }
                    }
                }
                $Report += ""
            }
        } else {
            $Report += "No sites found in topology"
            $Report += ""
        }
    } else {
        $Report += "No topology information available."
        $Report += "This may indicate insufficient permissions or Lync topology not accessible."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving topology: $($_.Exception.Message)"
    $Report += ""
}

# Section 5: Conference Directory (Enhanced)
$Report += "CONFERENCE DIRECTORIES"
$Report += $Separator
try {
    $ConferenceDirs = Get-CsConferenceDirectory -ErrorAction SilentlyContinue
    if ($ConferenceDirs) {
        $Report += "Total Conference Directories: $($ConferenceDirs.Count)"
        $Report += ""
        
        $Report += "CONFERENCE DIRECTORY DETAILS:"
        $ConferenceDirs | ForEach-Object {
            $Report += "  Directory ID: $($_.Identity)"
            $Report += "    Service: $($_.ServiceId)"
            
            # Try to identify the pool/service
            if ($_.ServiceId -and $AllPools) {
                $RelatedPool = $AllPools | Where-Object { $_.Identity -eq $_.ServiceId -or $_.Fqdn -eq $_.ServiceId }
                if ($RelatedPool) {
                    $Report += "    Pool: $($RelatedPool.Fqdn)"
                }
            }
            $Report += ""
        }
    } else {
        $Report += "No conference directories found."
        $Report += "This may indicate no conferencing services are configured."
        $Report += ""
    }
} catch {
    $Report += "Error retrieving conference directories: $($_.Exception.Message)"
    $Report += ""
}

# Infrastructure Health Summary
$Report += "INFRASTRUCTURE HEALTH SUMMARY"
$Report += $Separator
try {
    $HealthChecks = @()
    
    # Pool health
    if ($AllPools -and $AllPools.Count -gt 0) {
        $HealthChecks += "✅ Pools: $($AllPools.Count) pools discovered"
        
        $SBACount = ($AllPools | Where-Object { $_.Fqdn -like $SBAPattern }).Count
        if ($SBACount -gt 0) {
            $HealthChecks += "✅ SBA Deployment: $SBACount branch appliances for site resilience"
        }
        
        $EdgeCount = ($AllPools | Where-Object { $_.Fqdn -like $EdgePattern }).Count
        if ($EdgeCount -gt 0) {
            $HealthChecks += "✅ External Access: $EdgeCount edge pools configured"
        }
    } else {
        $HealthChecks += "⚠️  Pools: No pools found or access issues"
    }
    
    # Service health
    if ($AllServices -and $AllServices.Count -gt 0) {
        $HealthChecks += "✅ Services: $($AllServices.Count) services configured"
        
        $ServiceRoles = ($AllServices | Group-Object Role).Count
        $HealthChecks += "✅ Service Roles: $ServiceRoles different role types deployed"
    } else {
        $HealthChecks += "⚠️  Services: No services found or access issues"
    }
    
    # Computer health
    if ($AllComputers -and $AllComputers.Count -gt 0) {
        $HealthChecks += "✅ Computers: $($AllComputers.Count) computers in topology"
    } else {
        $HealthChecks += "⚠️  Computers: No computer information available"
    }
    
    # Conference health
    if ($ConferenceDirs -and $ConferenceDirs.Count -gt 0) {
        $HealthChecks += "✅ Conferencing: $($ConferenceDirs.Count) conference directories"
    } else {
        $HealthChecks += "⚠️  Conferencing: No conference directories found"
    }
    
    $HealthChecks | ForEach-Object { $Report += "  $_" }
    $Report += ""
    
} catch {
    $Report += "Error generating health summary: $($_.Exception.Message)"
    $Report += ""
}

$Report += ""
$Report += $Separator
$Report += "Report completed: $(Get-Date)"

# Export report
$Report | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host ""
Write-Host "ENHANCED LYNC INFRASTRUCTURE REPORT COMPLETE" -ForegroundColor Cyan
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow
Write-Host ""

# Display quick summary
if ($AllPools) {
    Write-Host "Infrastructure Summary:" -ForegroundColor White
    Write-Host "  Total Pools: $($AllPools.Count)" -ForegroundColor Green
    
    $SBACount = ($AllPools | Where-Object { $_.Fqdn -like $SBAPattern }).Count
    if ($SBACount -gt 0) {
        Write-Host "  SBA Sites: $SBACount" -ForegroundColor Green
    }
    
    if ($AllServices) {
        Write-Host "  Total Services: $($AllServices.Count)" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Key infrastructure components have been analyzed and documented." -ForegroundColor White
}