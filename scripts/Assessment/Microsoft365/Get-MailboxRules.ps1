<#
.SYNOPSIS
    Get all rules from user mailboxes or a specific user's mailbox

.DESCRIPTION
    This script retrieves mailbox rules from Exchange Online mailboxes and exports them to CSV format.
    
    Features:
    - Query all users or a specific user
    - Validates and creates output directory if needed
    - Exports rules with key properties (redirects, forwards, move actions)
    - Includes enabled/disabled status
    - Timestamps output file for tracking
   
.PARAMETER OutputDirectory
    Directory where the CSV file will be saved. Default: C:\Temp\MailboxRules

.PARAMETER UserPrincipalName
    Specific user's email address to check. If not provided, checks all mailboxes in the tenant.

.EXAMPLE
    .\Get-MailboxRules.ps1
    
    Exports rules from all mailboxes to default directory.

.EXAMPLE
    .\Get-MailboxRules.ps1 -OutputDirectory "C:\Reports\MailboxRules"
    
    Exports rules from all mailboxes to specified directory.

.EXAMPLE
    .\Get-MailboxRules.ps1 -UserPrincipalName "john.doe@contoso.com"
    
    Exports rules only for the specified user.

.EXAMPLE
    .\Get-MailboxRules.ps1 -UserPrincipalName "john.doe@contoso.com" -OutputDirectory "D:\Reports"
    
    Exports rules for specific user to custom directory.
   
.NOTES
    Name: Get-MailboxRules
    Author: W. Ford
    Version: 2.1
    DateCreated: 2022-11
    DateUpdated: 2025-12-23
    
    Version History:
    - 1.0: Initial release
    - 1.1: Added parameter, added if statement for file creation
    - 2.0: Added UserPrincipalName parameter, improved directory validation, added timestamped filenames,
           enhanced error handling, updated to modern standards
    - 2.1: Fixed Windows path handling, corrected default directory

    Requirements:
    - ExchangeOnlineManagement module
    - Exchange Online administrator permissions
    
.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-inboxrule
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Directory path for CSV export")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = "C:\Temp\MailboxRules",
    
    [Parameter(Mandatory = $false, HelpMessage = "Specific user email address to check")]
    [string]$UserPrincipalName
)

# Check if ExchangeOnlineManagement module is available
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "❌ ExchangeOnlineManagement module not installed" -ForegroundColor Red
    Write-Host "   Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
    Write-Host "✅ Connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "❌ Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Validate and create output directory
if (-not (Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
        Write-Host "✅ Created output directory: $OutputDirectory" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to create directory: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "✅ Output directory exists: $OutputDirectory" -ForegroundColor Green
}

# Test write permissions
$testFile = Join-Path $OutputDirectory "test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
try {
    "test" | Out-File -FilePath $testFile -ErrorAction Stop
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host "❌ No write permission to directory: $OutputDirectory" -ForegroundColor Red
    exit 1
}

# Determine which users to check
if ($UserPrincipalName) {
    Write-Host "ℹ️  Checking mailbox rules for: $UserPrincipalName" -ForegroundColor Cyan
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        $users = @($mailbox.UserPrincipalName)
        Write-Host "✅ Found mailbox: $($mailbox.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to find mailbox: $UserPrincipalName" -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }
}
else {
    Write-Host "ℹ️  Retrieving all mailboxes in tenant..." -ForegroundColor Cyan
    try {
        $users = (Get-Mailbox -ResultSize Unlimited -ErrorAction Stop).UserPrincipalName
        Write-Host "✅ Found $($users.Count) mailboxes to check" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to retrieve mailboxes: $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }
}

# Generate timestamped output filename
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if ($UserPrincipalName) {
    $outputFile = Join-Path $OutputDirectory "MailboxRules_$($UserPrincipalName.Split('@')[0])_$timestamp.csv"
}
else {
    $outputFile = Join-Path $OutputDirectory "MailboxRules_AllUsers_$timestamp.csv"
}

Write-Host "`nChecking mailbox rules..." -ForegroundColor Yellow

# Process each user and collect rules
$rulesFound = 0
$usersWithRules = 0
$processedCount = 0
$allRules = @()

foreach ($user in $users) {
    $processedCount++
    Write-Progress -Activity "Checking mailbox rules" -Status "Processing $user ($processedCount of $($users.Count))" -PercentComplete (($processedCount / $users.Count) * 100)
    
    try {
        $rules = Get-InboxRule -Mailbox $user -ErrorAction Stop
        
        if ($rules) {
            $usersWithRules++
            $rulesFound += $rules.Count
            
            $rulesData = $rules | Select-Object MailboxOwnerID, Name, Description, Enabled, RedirectTo, MoveToFolder, ForwardTo, ForwardAsAttachmentTo, DeleteMessage, MarkAsRead, StopProcessingRules
            $allRules += $rulesData
            
            Write-Host "  ✅ $user : $($rules.Count) rule(s) found" -ForegroundColor Green
        }
        else {
            Write-Host "  ℹ️  $user : No rules found" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  ⚠️  $user : Error checking rules - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Export all rules to CSV
if ($allRules.Count -gt 0) {
    try {
        $allRules | Export-Csv -Path $outputFile -NoTypeInformation -Force
        Write-Host "`n✅ Exported $($allRules.Count) rule(s) to: $outputFile" -ForegroundColor Green
    }
    catch {
        Write-Host "`n❌ Failed to export CSV: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    # Create empty file with headers
    try {
        $emptyRule = [PSCustomObject]@{
            MailboxOwnerID = ""
            Name = ""
            Description = ""
            Enabled = ""
            RedirectTo = ""
            MoveToFolder = ""
            ForwardTo = ""
            ForwardAsAttachmentTo = ""
            DeleteMessage = ""
            MarkAsRead = ""
            StopProcessingRules = ""
        }
        $emptyRule | Export-Csv -Path $outputFile -NoTypeInformation -Force
        Write-Host "`n⚠️  No rules found. Created empty CSV file: $outputFile" -ForegroundColor Yellow
    }
    catch {
        Write-Host "`n❌ Failed to create CSV file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Checking mailbox rules" -Completed

# Disconnect from Exchange Online
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "`n✅ Disconnected from Exchange Online" -ForegroundColor Green
}
catch {
    # Ignore disconnect errors
}

# Summary
Write-Host "`n" -NoNewline
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Export Complete" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Mailboxes Checked:     $($users.Count)" -ForegroundColor White
Write-Host "  Users with Rules:      $usersWithRules" -ForegroundColor White
Write-Host "  Total Rules Found:     $rulesFound" -ForegroundColor White
Write-Host "  Output File:           $outputFile" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan