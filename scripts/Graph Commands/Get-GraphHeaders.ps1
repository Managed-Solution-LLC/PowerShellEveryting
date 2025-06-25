# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Build the authorization headers for Microsoft Graph API requests.

.DESCRIPTION
    This function returns a hashtable with the required Authorization and Content-Type headers for Microsoft Graph REST API calls.

.PARAMETER accessToken
    The OAuth2 access token to use in the Authorization header.

.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.0.0
    This script is validated and safe for public release.

.EXAMPLE
    $headers = Get-GraphHeaders -accessToken $token
#>
Function Get-GraphHeaders {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$accessToken
    )
    # Define the authorization header
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type"  = "application/json"
    }
    # Return the headers
    return $headers
}