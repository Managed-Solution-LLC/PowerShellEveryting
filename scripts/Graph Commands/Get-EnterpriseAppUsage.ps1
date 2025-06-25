# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Report usage and ownership of all Azure AD Enterprise Applications (App Registrations).

.DESCRIPTION
    This script retrieves all enterprise applications in a tenant, gathers sign-in usage statistics for a specified timeframe, and lists application owners. Results can be exported to JSON. Handles throttling and permission errors gracefully.

.PARAMETER TenantId
    The Azure AD tenant ID to authenticate against.

.PARAMETER ClientId
    The application (client) ID. Optional if using interactive authentication.

.PARAMETER ClientSecret
    The client secret for app authentication. Optional if using interactive authentication.

.PARAMETER TimeFrame
    Number of days to look back for sign-in activity. Defaults to 30.

.PARAMETER export
    Switch to export results to EnterpriseAppUsage.json.

.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.0.0
    This script is validated and safe for public release.
    Requires: Microsoft.Graph, Get-GraphToken, Get-GraphHeaders, Get-AzureResourcePaging
    Permissions: AuditLog.Read.All, Directory.Read.All

.EXAMPLE
    Get-EnterpriseAppUsage -TenantId "<tenant-guid>" -ClientId "<app-id>" -ClientSecret "<secret>" -TimeFrame 60 -export
    Retrieves usage for the last 60 days and exports results to JSON.
#>
function Get-EnterpriseAppUsage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $false)]
        [number]$TimeFrame = 30, # Default to 30 days 
        [Parameter(Mandatory = $false)]
        [switch]$export
    )
    # Import the required module
    if ($null -eq (Get-Module -Name Microsoft.Graph -ErrorAction SilentlyContinue)) {
        Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } else {
        Write-Host "Microsoft.Graph module found." -ForegroundColor Green
    }

    try {
        $Scope = @("AuditLog.Read.All")
        if ($ClientSecret) {
            $MgToken = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Scope $Scope
        } else {
            # Get Graph Token
            $MgToken = Get-GraphToken -TenantId $TenantId -Scope $Scope
        }
    } catch {
        Write-Host "Error getting Graph Token: $_" -ForegroundColor Red
        return
    }
    
    # Get Headers for Graph API
    try {
        $headers = Get-GraphHeaders -AccessToken $MgToken
    } catch {
        Write-Host "Error getting Graph Headers: $_" -ForegroundColor Red
        return
    }

    # Build the API URL
    $BaseURL = "https://graph.microsoft.com/v1.0"

    #Build a Dateformat for the Filter
    $TimeFrameDate = Get-Date -format yyyy-MM-dd  ((Get-Date).AddDays(-$TimeFrame))
    
    #Build Array to store PSCustomObject
    $Array = @()

    # Get all Enterprise Applications
    # List Get all Apps from Azure
    $URLGetApps = "https://graph.microsoft.com/v1.0/applications"
    $AllApps = Get-AzureResourcePaging -URL $URLGetApps -AuthHeader $headers
    
    # Loop through each application and get usage details
    foreach ($App in $AllApps.value) {
        #Get Sign In/Usage
        try {
            $URLGetSignIns = "$BaseURL/auditLogs/signIns?`$filter=appid eq '$($App.appId)' and createdDateTime gt $TimeFrameDate"
            $maxRetries = 5
            $retryCount = 0
            $success = $false
            
            while (-not $success -and $retryCount -lt $maxRetries) {
            try {
                $SignIns = Invoke-RestMethod -Method GET -Uri $URLGetSignIns -Headers $headers -ErrorAction Stop
                $success = $true
            } catch {
                if ($_.Exception.Response.StatusCode.value__ -eq 429) {
                $retryAfter = 60 # Default retry after 60 seconds
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers.Contains("Retry-After")) {
                    $retryAfter = [int]$_.Exception.Response.Headers.GetValues("Retry-After")[0]
                }
                Write-Host "Request throttled. Retrying after $retryAfter seconds. Retry attempt $($retryCount + 1) of $maxRetries..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryAfter
                $retryCount++
                } else {
                throw $_
                }  
            }
            }
            
            if (-not $success) {
            Write-Host "Maximum retry attempts reached for $($App.displayName). Skipping..." -ForegroundColor Red
            $SignIns = @{value = @()}
            }
        } catch {
            if ($_.Exception.Response.StatusCode.value__ -eq 403) {
                Write-Host "Permission error getting sign-ins. Required permissions: AuditLog.Read.All and Directory.Read.All" -ForegroundColor Yellow
                $SignIns = @{value = @()}
            } else {
                Write-Host "Error getting sign-ins for $($App.displayName): $_" -ForegroundColor Red
                $SignIns = @{value = @()}
            }
        }
        
        Start-Sleep -Seconds 1
        
        # Get Owners
        $URLGetOwner = "$BaseURL/applications/$($App.id)/owners"
        try {
            $Owner = Invoke-RestMethod -Method GET -Uri $URLGetOwner -Headers $headers
            # Check if owner collection is empty
            if ($null -eq $Owner.value -or $Owner.value.Count -eq 0) {
            Write-Verbose "No owners found for application $($App.displayName)"
            $Owner = $null
            }
        } catch {
            Write-Host "Error getting owners: $_" -ForegroundColor Red
            return
        }

        if ($Owner) {
            foreach ($o in $Owner.value) {
                $Array += [PSCustomObject]@{
                    "App ID"           = $App.id
                    "App AppID"        = $App.appId
                    "App Name"         = $App.displayName
                    "Owner UPN"        = $o.userprincipalname
                    "Owner Name"       = $o.displayName
                    "Owner ID"         = $o.id
                    "Usage Count"      = ($SignIns.value ).count
                }
            }
        } else {
            $Array += [PSCustomObject]@{
                "App ID"           = $App.id
                "App AppID"        = $App.appId
                "App Name"         = $App.displayName
                "Owner UPN"        = "NONE"
                "Owner Name"       = "NONE"
                "Owner ID"         = "NONE"
                "Usage Count"      = ($SignIns.value ).count
            }
        }
    }
    if ($Array.Count -ne 0) {
        $Array | Select-Object -Property "App Name", "Owner UPN", "Usage Count" | Sort-Object -Property "Usage Count" -Descending
        if ($export) {
            $Array | ConvertTo-Json -depth 10 | Out-File -FilePath "EnterpriseAppUsage.json" -Force
            Write-Host "Exported to EnterpriseAppUsage.csv" -ForegroundColor Green
        }
        } else {
            Write-Host "No applications found." -ForegroundColor Yellow
        }
}