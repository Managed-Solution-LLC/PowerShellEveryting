# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Retrieve and export Exchange Online user provisioning errors using Microsoft Graph.

.DESCRIPTION
    This script queries Microsoft Graph for users with provisioning errors in Exchange Online, formats the errors for readability, and optionally exports the results to JSON. Handles paging and ensures all results are collected.

.PARAMETER TenantId
    The Azure AD tenant ID to authenticate against.

.PARAMETER ClientId
    The application (client) ID. Optional if using interactive authentication.

.PARAMETER ClientSecret
    The client secret for app authentication. Optional if using interactive authentication.

.PARAMETER export
    Switch to export results to C:\temp\ExchangeErrors.json.

.NOTES
    Author: Managed Solution
    Date: 2025-06-25
    Version: 1.0.0
    This script is validated and safe for public release.
    Requires: Microsoft.Graph, Get-GraphToken, Get-GraphHeaders
    Permissions: Directory.Read.All

.EXAMPLE
    Get-ExchangeErrorsGraph -TenantId "<tenant-guid>" -ClientId "<app-id>" -ClientSecret "<secret>" -export
    Retrieves users with provisioning errors and exports results to JSON.
#>
Function Get-ExchangeErrorsGraph {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret,
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
        if ($ClientSecret) {
            $MgToken = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        } else {
            # Get Graph Token
            $MgToken = Get-GraphToken -TenantId $TenantId 
        }
    } catch {
        Write-Host "Error getting Graph Token: $_" -ForegroundColor Red
        return
    }
    
    # Get Headers for Graph API
    try {
        $headers = Get-GraphHeaders -AccessToken $MgToken.AccessToken
    } catch {
        Write-Host "Error getting Graph Headers: $_" -ForegroundColor Red
        return
    }

    try {
        # Define the request URI
        $requestUri = "https://graph.microsoft.com/beta/users?$select=id,userPrincipalName,serviceProvisioningErrors"

        $response = Invoke-WebRequest -Uri $requestUri -Headers $headers -Method Get
        # Parse the response
        $result = $response.Content | ConvertFrom-Json
        $allUsers=@()
        if (!($result.'@odata.nextLink')){
            $allUsers += $result.value | Where-Object {$_.serviceProvisioningErrors -notlike $null}
        } else {
            do {
                write-host "Loading Page" -ForegroundColor Green
                $response = Invoke-WebRequest -Uri $result.'@odata.nextLink' -Headers $headers -Method Get
                $result = $response.Content | ConvertFrom-Json
                $allUsers += $result.value | Where-Object {$_.serviceProvisioningErrors -notlike $null}
            
                # Update the request URI to the next link, if it exists
                $uri = $result.'@odata.nextLink'
                Write-host "Getting to the next page.. Please"  -ForegroundColor Green
            } while ($null -ne $uri)
        }
    } catch {
        Write-Host "Error making API request: $_" -ForegroundColor Red
        return
    }

    if ($allUsers) {
        Write-Host "Found $($allUsers.Count) users with provisioning errors." -ForegroundColor Green
        # Check if any errors were found
        # Format the errors in a more readable way
        $formattedUsers = $allUsers | Select-Object -Property UserPrincipalName, 
            @{
            Name = "Errors"; 
            Expression = {
                $errorDetails = ([xml]$_.serviceProvisioningErrors.errorDetail).ServiceInstance.ObjectErrors.ErrorRecord.ErrorDescription
                if ($errorDetails -is [array]) {
                $errorDetails -join "`n"
                } else {
                $errorDetails
                }
            }
            }
        
        # Display the formatted results
        $formattedUsers | Format-Table -AutoSize -Wrap
        if ($export) {
            # Check if Path Exists
            if (!(Test-Path -Path "C:\temp")) {
                New-Item -Path "C:\" -Name "temp" -ItemType Directory -Force
            }
            # Check if File Exists
            if (Test-Path -Path "C:\temp\ExchangeErrors.json") {
                Remove-Item -Path "C:\temp\ExchangeErrors.json" -Force
            }
            Write-Host "Exporting results to C:\temp\ExchangeErrors.json" -ForegroundColor Green
            # Export to JSON
            $allUsers | Select-Object userPrincipalName, @{n="Errors";e={([xml]$_.serviceProvisioningErrors.errorDetail).ServiceInstance.ObjectErrors.ErrorRecord.ErrorDescription } } | ConvertTo-Json -Depth 10 | Out-File -FilePath "C:\temp\ExchangeErrors.json" -Force
        }
    } else {
        Write-Host "No users found with provisioning errors." -ForegroundColor Yellow
        if ($export) {
            Write-Host "Nothing to Export" -ForegroundColor Yellow
        }
    }
}