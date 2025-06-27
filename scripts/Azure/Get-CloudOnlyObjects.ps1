# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
<#
.SYNOPSIS
    Export all cloud-only users, groups, and distribution groups from Microsoft 365/Azure AD.

.DESCRIPTION
    This script exports lists of cloud-only users, groups, and distribution groups (not synced from on-premises AD) using Microsoft Graph and Exchange Online.
    Results are saved as CSV and JSON files in a specified output directory. You can exclude users by UPN pattern.

.PARAMETER OutputPath
    The directory where export files will be saved. Defaults to .\CloudOnlyExports

.PARAMETER ExcludePattern
    Array of patterns to exclude users by UserPrincipalName (UPN).

.NOTES
    Required Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement
    Output: CSV and JSON files for users, groups, and distribution groups

.EXAMPLE
    .\Get-CloudOnlyObjects.ps1
    Exports all cloud-only users, groups, and distribution groups to the default output directory.

.EXAMPLE
    .\Get-CloudOnlyObjects.ps1 -OutputPath "C:\Exports" -ExcludePattern "*@testdomain.com"
    Exports to C:\Exports and excludes users with UPNs matching *@testdomain.com
.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.0
    Required Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement
    Output: CSV and JSON files in the specified output directory
#>

# Export a list of all cloud users, groups, distros

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement

# Simple script to export cloud-only objects
param(
    [string]$OutputPath = ".\CloudOnlyExports",
    [string[]]$ExcludePattern = @()
)

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

Write-Host "Exporting cloud-only objects..." -ForegroundColor Cyan

# Export cloud-only users
Write-Host "Getting cloud-only users..." -ForegroundColor Yellow
$allUsers = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled,OnPremisesImmutableId,AccountEnabled,UserType,Mail,JobTitle,Department"

$cloudOnlyUsers = $allUsers | Where-Object {
    # Exclude synced users
    ($_.OnPremisesSyncEnabled -ne $true -and [string]::IsNullOrEmpty($_.OnPremisesImmutableId)) -and
    # Exclude pattern matches
    (-not ($ExcludePattern | Where-Object { $_.UserPrincipalName -like $_ }))
} | Select-Object Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType, JobTitle, Department

$cloudOnlyUsers | Export-Csv -Path "$OutputPath\CloudOnly_Users_$timestamp.csv" -NoTypeInformation
$cloudOnlyUsers | ConvertTo-Json -Depth 3 | Out-File "$OutputPath\CloudOnly_Users_$timestamp.json"
Write-Host "✓ Exported $($cloudOnlyUsers.Count) cloud-only users" -ForegroundColor Green

# Export cloud-only groups
Write-Host "Getting cloud-only groups..." -ForegroundColor Yellow
$allGroups = Get-MgGroup -All -Property "Id,DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,OnPremisesSyncEnabled,Description,Visibility"

$cloudOnlyGroups = $allGroups | Where-Object {
    $_.OnPremisesSyncEnabled -ne $true
} | Select-Object Id, DisplayName, Mail, MailEnabled, SecurityEnabled, 
    @{Name="GroupType"; Expression={
        if ($_.GroupTypes -contains "Unified") { "Microsoft 365" }
        elseif ($_.MailEnabled -and -not $_.SecurityEnabled) { "Distribution" }
        elseif ($_.MailEnabled -and $_.SecurityEnabled) { "Mail-enabled Security" }
        else { "Security" }
    }}, Description, Visibility

$cloudOnlyGroups | Export-Csv -Path "$OutputPath\CloudOnly_Groups_$timestamp.csv" -NoTypeInformation
$cloudOnlyGroups | ConvertTo-Json -Depth 3 | Out-File "$OutputPath\CloudOnly_Groups_$timestamp.json"
Write-Host "✓ Exported $($cloudOnlyGroups.Count) cloud-only groups" -ForegroundColor Green

# Export cloud-only distribution groups
Write-Host "Getting cloud-only distribution groups..." -ForegroundColor Yellow
$allDistGroups = Get-DistributionGroup -ResultSize Unlimited

$cloudOnlyDistGroups = $allDistGroups | Where-Object {
    $_.IsDirSynced -ne $true
} | ForEach-Object {
    $members = @(Get-DistributionGroupMember -Identity $_.Identity -ErrorAction SilentlyContinue)
    [PSCustomObject]@{
        Identity = $_.Identity
        DisplayName = $_.DisplayName
        PrimarySmtpAddress = $_.PrimarySmtpAddress
        RecipientType = $_.RecipientType
        MemberCount = $members.Count
        RequireSenderAuthenticationEnabled = $_.RequireSenderAuthenticationEnabled
        HiddenFromAddressListsEnabled = $_.HiddenFromAddressListsEnabled
        WhenCreated = $_.WhenCreated
    }
}

$cloudOnlyDistGroups | Export-Csv -Path "$OutputPath\CloudOnly_DistributionGroups_$timestamp.csv" -NoTypeInformation
$cloudOnlyDistGroups | ConvertTo-Json -Depth 3 | Out-File "$OutputPath\CloudOnly_DistributionGroups_$timestamp.json"
Write-Host "✓ Exported $($cloudOnlyDistGroups.Count) cloud-only distribution groups" -ForegroundColor Green

Write-Host "`nExport Summary:" -ForegroundColor Cyan
Write-Host "  Users: $($cloudOnlyUsers.Count)" -ForegroundColor White
Write-Host "  Groups: $($cloudOnlyGroups.Count)" -ForegroundColor White
Write-Host "  Distribution Groups: $($cloudOnlyDistGroups.Count)" -ForegroundColor White
Write-Host "  Files saved to: $OutputPath" -ForegroundColor White
