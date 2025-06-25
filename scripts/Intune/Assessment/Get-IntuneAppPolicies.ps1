# PowerShell function to get all Intune App Protection Policies and export them to a CSV file

function Get-IntuneAppPolicies {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    # Connect to Intune
    $IntuneConnection = Get-IntuneConnection

    # Get all Intune App Protection Policies
    $AppPolicies = Get-IntuneAppProtectionPolicy -Connection $IntuneConnection

    # Export the App Policies to a CSV file
    $AppPolicies | Export-Csv -Path $ExportPath -NoTypeInformation
}
