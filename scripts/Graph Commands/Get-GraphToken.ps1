# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Acquire an OAuth2 access token for Microsoft Graph using MSAL.PS.

.DESCRIPTION
    This function retrieves an access token for Microsoft Graph using either client secret or interactive authentication.
    It installs the MSAL.PS module if not already present. The default client ID is the Microsoft Graph PowerShell client.

.PARAMETER TenantId
    The Azure AD tenant ID to authenticate against.

.PARAMETER ClientId
    The application (client) ID. Defaults to the Microsoft Graph PowerShell client ID.

.PARAMETER ClientSecret
    The client secret for app authentication. If not provided, interactive authentication is used.

.PARAMETER Scope
    The scopes to request. Defaults to 'https://graph.microsoft.com/.default'.

.EXAMPLE
    $token = Get-GraphToken -TenantId "<tenant-guid>" -ClientSecret "<secret>"
    Gets a token using client credentials.

.EXAMPLE
    $token = Get-GraphToken -TenantId "<tenant-guid>"
    Prompts for interactive authentication and returns a token.
.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.0
    Required Modules: MSAL.PS
    Output: Access token string for Microsoft Graph API
#>
function Get-GraphToken {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e", # Microsoft Graph PowerShell client ID
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $false)]
        [array]$Scope = @("https://graph.microsoft.com/.default")
    )

    # Import the Microsoft Authentication Library (MSAL) module if not already installed
    $module = Get-Module -Name MSAL.PS -ListAvailable
    if (-not $module) {
        Install-Module -Name MSAL.PS -Force
    } else {
        Write-Host "MSAL.PS module is already installed."
    }

    if ($ClientSecret) {
        $authParams = @{
            ClientId     = $ClientId
            TenantId     = $TenantId
            ClientSecret = $ClientSecret
            Scopes       = $Scope
        }
    } else {
        $authParams = @{
            ClientId = $ClientId
            TenantId = $TenantId
            Interactive = $true
            Scopes = $Scope
        }
    }
    
    # Get the access token
    $authResult = Get-MsalToken @authParams
    if ($authResult) {
        Write-Host "Access token retrieved successfully."
        # Format the token for use with the Graph API
        $accessToken = $authResult.AccessToken
    } else {
        Write-Host "Failed to retrieve access token." -ForegroundColor Red
        $accessToken = $null
    }
    return $accessToken
}