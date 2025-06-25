# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Retrieve all paged resources from an Azure REST API endpoint using @odata.nextLink.

.DESCRIPTION
    This function retrieves all resources from a paged Azure REST API endpoint by following the @odata.nextLink property until all pages are collected.
    It returns a combined array of all resources.

.PARAMETER URL
    The initial API endpoint URL to query.

.PARAMETER AuthHeader
    The authentication header (e.g., Bearer token) to use for the API requests.

.EXAMPLE
    $header = @{ Authorization = "Bearer $token" }
    $allApps = Get-AzureResourcePaging -URL "https://graph.microsoft.com/v1.0/applications" -AuthHeader $header

.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.0.0
    This script is validated and safe for public release.
#>
function Get-AzureResourcePaging {
    param (
        $URL,
        $AuthHeader
    )

    # List Get all Apps from Azure

    $Response = Invoke-RestMethod -Method GET -Uri $URL -Headers $AuthHeader
    $Resources = $Response.value

    $ResponseNextLink = $Response."@odata.nextLink"
    while ($ResponseNextLink -ne $null) {

        $Response = (Invoke-RestMethod -Uri $ResponseNextLink -Headers $AuthHeader -Method Get)
        $ResponseNextLink = $Response."@odata.nextLink"
        $Resources += $Response.value
    }
    return $Resources
}