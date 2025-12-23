<#
.SYNOPSIS
    Comprehensive mailbox permissions and delegation audit for Exchange Online
.DESCRIPTION
    Collects detailed mailbox delegation permissions including:
    - Full Access (Read and Manage) permissions
    - Send As permissions
    - Send on Behalf permissions
    - Inbox folder permissions
    - Calendar folder permissions
    - Email forwarding configuration
    
    Results are exported to CSV with optional ZIP archive for easy distribution.
    Essential for security audits, compliance reviews, and migration planning.
.PARAMETER MailboxFilter
    Filter mailboxes to audit. Options:
    - "All" (default): All user and shared mailboxes
    - "UserMailboxes": Only user mailboxes
    - "SharedMailboxes": Only shared mailboxes
.PARAMETER UserPrincipalName
    Audit specific users by UPN. Accepts comma-separated list.
    Example: "user1@contoso.com","user2@contoso.com"
.PARAMETER CsvFilePath
    Path to CSV file containing list of users to audit.
    CSV format: UserPrincipalName,DisplayName
.PARAMETER IncludeFolderPermissions
    Include inbox and calendar folder permissions (user-delegated).
    Default: $true. Set to $false to speed up collection.
.PARAMETER ResolveDisplayNames
    Resolve user identities to display names for better readability.
    Default: $true. Set to $false to improve performance on large tenants.
.PARAMETER InboxFolderName
    Name of inbox folder in your environment. Default: "Inbox"
.PARAMETER CalendarFolderName
    Name of calendar folder in your environment. Default: "Calendar"
.PARAMETER OutputDirectory
    Custom output directory path. Default: Documents\MailboxPermissions_<timestamp>
.PARAMETER CreateZip
    Create ZIP archive of the report. Default: $true
.EXAMPLE
    .\Get-MailboxPermissionsReport.ps1
    
    Audit all mailboxes with full permissions detail
.EXAMPLE
    .\Get-MailboxPermissionsReport.ps1 -MailboxFilter SharedMailboxes
    
    Audit only shared mailboxes
.EXAMPLE
    .\Get-MailboxPermissionsReport.ps1 -UserPrincipalName "user1@contoso.com","user2@contoso.com"
    
    Audit specific users
.EXAMPLE
    .\Get-MailboxPermissionsReport.ps1 -CsvFilePath "C:\Users.csv"
    
    Audit users from CSV file
.EXAMPLE
    .\Get-MailboxPermissionsReport.ps1 -IncludeFolderPermissions:$false
    
    Quick audit without folder-level permissions
.NOTES
    Author: W. Ford (Managed Solution LLC)
    Date: 2025-12-23
    Version: 2.0
    
    Requirements:
    - ExchangeOnlineManagement module v2.0 or later
    - Exchange Administrator or Global Reader role
    - PowerShell 5.1 or later
    
    Based on original work by R. Mens - LazyAdmin.nl
    Enhanced with project standards and modern patterns
.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailboxpermission
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Filter mailboxes: All, UserMailboxes, or SharedMailboxes")]
    [ValidateSet("All", "UserMailboxes", "SharedMailboxes")]
    [string]$MailboxFilter = "All",
    
    [Parameter(Mandatory=$false, HelpMessage="Comma-separated list of UserPrincipalNames to audit")]
    [string[]]$UserPrincipalName,
    
    [Parameter(Mandatory=$false, HelpMessage="Path to CSV file with list of users")]
    [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory=$false, HelpMessage="Include inbox and calendar folder permissions")]
    [bool]$IncludeFolderPermissions = $true,
    
    [Parameter(Mandatory=$false, HelpMessage="Resolve identities to display names")]
    [bool]$ResolveDisplayNames = $true,
    
    [Parameter(Mandatory=$false, HelpMessage="Inbox folder name")]
    [string]$InboxFolderName = "Inbox",
    
    [Parameter(Mandatory=$false, HelpMessage="Calendar folder name")]
    [string]$CalendarFolderName = "Calendar",
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory path")]
    [string]$OutputDirectory,
    
    [Parameter(Mandatory=$false, HelpMessage="Create ZIP archive")]
    [bool]$CreateZip = $true
)

$ErrorActionPreference = 'Stop'
$StartTime = Get-Date
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if ([string]::IsNullOrEmpty($OutputDirectory)) {
    $OutputDirectory = "$env:USERPROFILE\Documents\MailboxPermissions_$Timestamp"
}

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  MAILBOX PERMISSIONS AUDIT" -ForegroundColor Cyan
Write-Host "  Exchange Online Delegation Report" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# Create output directory
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
Write-Host "✅ Output directory: $OutputDirectory" -ForegroundColor Green

# Check and install Exchange Online Management module
Write-Host "`n📦 Checking Exchange Online Management module..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "   Installing ExchangeOnlineManagement..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop | Out-Null
        Write-Host "   ✅ ExchangeOnlineManagement installed" -ForegroundColor Green
    }
    catch {
        Write-Host "   ❌ Failed to install ExchangeOnlineManagement" -ForegroundColor Red
        Write-Host "   You may need to run as Administrator" -ForegroundColor Yellow
        exit 1
    }
}
else {
    Write-Host "   ✅ ExchangeOnlineManagement ready" -ForegroundColor Green
}

# Connect to Exchange Online
Write-Host "`n🔌 Connecting to Exchange Online..." -ForegroundColor Cyan
Write-Host "   (Sign-in prompt will appear)" -ForegroundColor Yellow
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    
    # Verify connection
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "✅ Connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "❌ Failed to connect to Exchange Online" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Yellow
    exit 1
}

# Helper Functions
function Get-DisplayName {
    <#
    .SYNOPSIS
        Resolve identity to display name or return UPN
    #>
    param([Parameter(Mandatory=$true)]$Identity)
    
    if ($ResolveDisplayNames) {
        try {
            return (Get-EXOMailbox -Identity $Identity -ErrorAction Stop).DisplayName
        }
        catch {
            return $Identity
        }
    }
    else {
        return $Identity.ToString().Split("@")[0]
    }
}

function Get-TargetMailboxes {
    <#
    .SYNOPSIS
        Get mailboxes to audit based on parameters
    #>
    Write-Host "`n📊 Collecting mailboxes to audit..." -ForegroundColor Cyan
    $mailboxes = @()
    
    if ($UserPrincipalName) {
        # Audit specific users
        Write-Host "   Mode: Specific users" -ForegroundColor Gray
        foreach ($upn in $UserPrincipalName) {
            Write-Host "   - Getting mailbox: $upn" -ForegroundColor Gray
            try {
                $mb = Get-EXOMailbox -Identity $upn -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress -ErrorAction Stop
                $mailboxes += $mb | Select-Object UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress
            }
            catch {
                Write-Host "   ⚠️  Could not find mailbox: $upn" -ForegroundColor Yellow
            }
        }
    }
    elseif ($CsvFilePath) {
        # Audit users from CSV
        Write-Host "   Mode: CSV file import" -ForegroundColor Gray
        Write-Host "   File: $CsvFilePath" -ForegroundColor Gray
        
        Import-Csv $CsvFilePath | ForEach-Object {
            Write-Host "   - Getting mailbox: $($_.UserPrincipalName)" -ForegroundColor Gray
            try {
                $mb = Get-EXOMailbox -Identity $_.UserPrincipalName -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress -ErrorAction Stop
                $mailboxes += $mb | Select-Object UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress
            }
            catch {
                Write-Host "   ⚠️  Could not find mailbox: $($_.UserPrincipalName)" -ForegroundColor Yellow
            }
        }
    }
    else {
        # Audit all mailboxes based on filter
        Write-Host "   Mode: $MailboxFilter" -ForegroundColor Gray
        
        switch ($MailboxFilter) {
            "UserMailboxes" { $recipientTypes = "UserMailbox" }
            "SharedMailboxes" { $recipientTypes = "SharedMailbox" }
            "All" { $recipientTypes = "UserMailbox,SharedMailbox" }
        }
        
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress -ErrorAction Stop |
            Select-Object UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress
    }
    
    Write-Host "✅ Found $($mailboxes.Count) mailboxes to audit" -ForegroundColor Green
    return $mailboxes
}

function Get-SendOnBehalfUsers {
    <#
    .SYNOPSIS
        Get users with Send on Behalf permissions
    #>
    param([Parameter(Mandatory=$true)]$Mailbox)
    
    $users = @()
    if ($Mailbox.GrantSendOnBehalfTo -ne $null) {
        $Mailbox.GrantSendOnBehalfTo | ForEach-Object {
            $users += Get-DisplayName -Identity $_
        }
    }
    return $users
}

function Get-SendAsUsers {
    <#
    .SYNOPSIS
        Get users with Send As permissions
    #>
    param([Parameter(Mandatory=$true)]$Identity)
    
    $permissions = Get-EXORecipientPermission -Identity $Identity -ErrorAction SilentlyContinue | 
        Where-Object { -not ($_.Trustee -match "NT AUTHORITY") -and ($_.IsInherited -eq $false) }
    
    $users = @()
    $permissions | ForEach-Object {
        $users += Get-DisplayName -Identity $_.Trustee
    }
    return $users
}

function Get-FullAccessUsers {
    <#
    .SYNOPSIS
        Get users with Full Access (Read and Manage) permissions
    #>
    param([Parameter(Mandatory=$true)]$Identity)
    
    $permissions = Get-EXOMailboxPermission -Identity $Identity -ErrorAction SilentlyContinue | 
        Where-Object { -not ($_.User -match "NT AUTHORITY") -and ($_.IsInherited -eq $false) }
    
    $users = @()
    $permissions | ForEach-Object {
        $users += Get-DisplayName -Identity $_.User
    }
    return $users
}

function Get-FolderPermissions {
    <#
    .SYNOPSIS
        Get folder-level permissions (Inbox or Calendar)
    #>
    param(
        [Parameter(Mandatory=$true)]$Identity,
        [Parameter(Mandatory=$true)]$FolderName
    )
    
    $result = @{
        Users = @()
        AccessRights = @()
        Delegated = @()
    }
    
    try {
        $ErrorActionPreference = "Stop"
        $permissions = Get-EXOMailboxFolderPermission -Identity "$($Identity):\$($FolderName)" -ErrorAction Stop | 
            Where-Object { -not ($_.User -match "Default") -and -not ($_.AccessRights -match "None") }
        
        $permissions | ForEach-Object {
            $result.Users += Get-DisplayName -Identity $_.User
            $result.AccessRights += $_.AccessRights -join ","
            $result.Delegated += $_.SharingPermissionFlags
        }
    }
    catch {
        # Folder doesn't exist or access denied - return empty arrays
    }
    finally {
        $ErrorActionPreference = "Continue"
    }
    
    return $result
}

function Get-MaxCount {
    <#
    .SYNOPSIS
        Find the maximum count among permission arrays
    #>
    param($Counts)
    
    $max = 0
    $Counts | ForEach-Object {
        if ($_ -gt $max) { $max = $_ }
    }
    return $max
}

# Collect permissions
Write-Host "`n🔍 Collecting permissions..." -ForegroundColor Cyan
$mailboxes = Get-TargetMailboxes

if ($mailboxes.Count -eq 0) {
    Write-Host "❌ No mailboxes found to audit" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$results = @()
$i = 0
$errorCount = 0

foreach ($mailbox in $mailboxes) {
    $i++
    Write-Progress -Activity "Auditing Mailbox Permissions" -Status "$($mailbox.DisplayName)" -PercentComplete (($i/$mailboxes.Count)*100)
    
    try {
        # Collect all permission types
        $sendOnBehalf = Get-SendOnBehalfUsers -Mailbox $mailbox
        $sendAs = Get-SendAsUsers -Identity $mailbox.UserPrincipalName
        $fullAccess = Get-FullAccessUsers -Identity $mailbox.UserPrincipalName
        
        # Get folder permissions if requested
        if ($IncludeFolderPermissions) {
            $inbox = Get-FolderPermissions -Identity $mailbox.UserPrincipalName -FolderName $InboxFolderName
            $calendar = Get-FolderPermissions -Identity $mailbox.UserPrincipalName -FolderName $CalendarFolderName
        }
        else {
            $inbox = @{ Users = @(); AccessRights = @(); Delegated = @() }
            $calendar = @{ Users = @(); AccessRights = @(); Delegated = @() }
        }
        
        # Find max records to handle multiple delegations
        $maxCount = Get-MaxCount -Counts @(
            $fullAccess.Count,
            $sendAs.Count,
            $sendOnBehalf.Count,
            $inbox.Users.Count,
            $calendar.Users.Count
        )
        
        if ($maxCount -gt 0) {
            # Create rows for each delegation
            for ($x = 0; $x -lt $maxCount; $x++) {
                $results += [PSCustomObject]@{
                    "Display Name" = if ($x -eq 0) { $mailbox.DisplayName } else { "" }
                    "Email Address" = if ($x -eq 0) { $mailbox.PrimarySMTPAddress } else { "" }
                    "Mailbox Type" = if ($x -eq 0) { $mailbox.RecipientTypeDetails } else { "" }
                    "Forwarding Address" = if ($x -eq 0) { $mailbox.ForwardingSMTPAddress } else { "" }
                    "Full Access (Read & Manage)" = if ($x -lt $fullAccess.Count) { $fullAccess[$x] } else { "" }
                    "Send As" = if ($x -lt $sendAs.Count) { $sendAs[$x] } else { "" }
                    "Send on Behalf" = if ($x -lt $sendOnBehalf.Count) { $sendOnBehalf[$x] } else { "" }
                    "Inbox User" = if ($x -lt $inbox.Users.Count) { $inbox.Users[$x] } else { "" }
                    "Inbox Permission" = if ($x -lt $inbox.AccessRights.Count) { $inbox.AccessRights[$x] } else { "" }
                    "Inbox Delegated" = if ($x -lt $inbox.Delegated.Count) { $inbox.Delegated[$x] } else { "" }
                    "Calendar User" = if ($x -lt $calendar.Users.Count) { $calendar.Users[$x] } else { "" }
                    "Calendar Permission" = if ($x -lt $calendar.AccessRights.Count) { $calendar.AccessRights[$x] } else { "" }
                    "Calendar Delegated" = if ($x -lt $calendar.Delegated.Count) { $calendar.Delegated[$x] } else { "" }
                }
            }
        }
        else {
            # No permissions found, still include mailbox in report
            $results += [PSCustomObject]@{
                "Display Name" = $mailbox.DisplayName
                "Email Address" = $mailbox.PrimarySMTPAddress
                "Mailbox Type" = $mailbox.RecipientTypeDetails
                "Forwarding Address" = $mailbox.ForwardingSMTPAddress
                "Full Access (Read & Manage)" = ""
                "Send As" = ""
                "Send on Behalf" = ""
                "Inbox User" = ""
                "Inbox Permission" = ""
                "Inbox Delegated" = ""
                "Calendar User" = ""
                "Calendar Permission" = ""
                "Calendar Delegated" = ""
            }
        }
    }
    catch {
        Write-Host "   ⚠️  Error processing $($mailbox.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
        $errorCount++
    }
}
Write-Progress -Activity "Auditing Mailbox Permissions" -Completed

# Export results
Write-Host "`n📝 Exporting results..." -ForegroundColor Cyan
$csvFile = Join-Path $OutputDirectory "MailboxPermissions.csv"
$results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "✅ Exported $($results.Count) permission records" -ForegroundColor Green

if ($errorCount -gt 0) {
    Write-Host "   ⚠️  $errorCount mailboxes had errors" -ForegroundColor Yellow
}

# Generate summary
Write-Host "`n📊 Generating summary..." -ForegroundColor Cyan
$summaryFile = Join-Path $OutputDirectory "Summary.txt"

# Calculate statistics
$mailboxesWithPermissions = ($results | Where-Object { $_."Full Access (Read & Manage)" -or $_."Send As" -or $_."Send on Behalf" }).Count
$fullAccessCount = ($results | Where-Object { $_."Full Access (Read & Manage)" }).Count
$sendAsCount = ($results | Where-Object { $_."Send As" }).Count
$sendOnBehalfCount = ($results | Where-Object { $_."Send on Behalf" }).Count
$forwardingCount = ($results | Where-Object { $_."Forwarding Address" }).Count
$inboxSharingCount = ($results | Where-Object { $_."Inbox User" }).Count
$calendarSharingCount = ($results | Where-Object { $_."Calendar User" }).Count

$summary = @"
============================================
MAILBOX PERMISSIONS AUDIT SUMMARY
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
============================================

AUDIT SCOPE
----------------------------------------
Total Mailboxes Audited: $($mailboxes.Count)
Mailbox Filter: $MailboxFilter
Folder Permissions: $(if($IncludeFolderPermissions){'Included'}else{'Excluded'})

DELEGATION SUMMARY
----------------------------------------
Mailboxes with Delegations: $mailboxesWithPermissions
Full Access Permissions: $fullAccessCount
Send As Permissions: $sendAsCount
Send on Behalf Permissions: $sendOnBehalfCount
Forwarding Configured: $forwardingCount

FOLDER SHARING
----------------------------------------
Inbox Sharing: $inboxSharingCount
Calendar Sharing: $calendarSharingCount

TOP DELEGATED MAILBOXES
----------------------------------------
$(($results | Where-Object { $_."Display Name" -ne "" } | 
    Group-Object "Email Address" | 
    Select-Object @{N="Mailbox";E={$_.Name}}, @{N="Delegations";E={$_.Count-1}} |
    Where-Object { $_.Delegations -gt 0 } |
    Sort-Object Delegations -Descending | 
    Select-Object -First 10 | 
    ForEach-Object { "  • $($_.Mailbox) - $($_.Delegations) delegations" }) -join "`n")

EXECUTION
----------------------------------------
Duration: $(((Get-Date) - $StartTime).ToString('mm\:ss'))
Errors: $errorCount
============================================
"@

$summary | Out-File -FilePath $summaryFile -Encoding UTF8
Write-Host "✅ Summary saved" -ForegroundColor Green

# Create ZIP archive
if ($CreateZip) {
    Write-Host "`n📦 Creating ZIP archive..." -ForegroundColor Cyan
    $zipFile = "$($OutputDirectory).zip"
    $filesToZip = Get-ChildItem -Path $OutputDirectory -File
    
    if ($filesToZip.Count -gt 0) {
        Compress-Archive -Path "$OutputDirectory\*" -DestinationPath $zipFile -Force
        Write-Host "✅ Archive created: $zipFile" -ForegroundColor Green
    }
}

# Cleanup
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

# Display results
Write-Host "`n============================================" -ForegroundColor Green
Write-Host "  ✅ PERMISSIONS AUDIT COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

if ($CreateZip -and $zipFile -and (Test-Path $zipFile)) {
    Write-Host "`nReport archive:" -ForegroundColor Cyan
    Write-Host "  $zipFile" -ForegroundColor Yellow
    
    # Open location in Explorer
    Write-Host "`nOpening report location..." -ForegroundColor Cyan
    Start-Process "explorer.exe" -ArgumentList "/select,`"$zipFile`""
}
else {
    Write-Host "`nReport directory:" -ForegroundColor Cyan
    Write-Host "  $OutputDirectory" -ForegroundColor Yellow
}

Write-Host "`nExecution time: $(((Get-Date) - $StartTime).ToString('mm\:ss'))" -ForegroundColor Cyan
Write-Host "Mailboxes audited: $($mailboxes.Count)" -ForegroundColor Cyan
Write-Host "Delegations found: $mailboxesWithPermissions" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Green

# Display summary
Write-Host $summary
