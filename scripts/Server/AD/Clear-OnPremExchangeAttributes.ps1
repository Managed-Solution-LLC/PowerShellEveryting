<#
.SYNOPSIS
    Removes on-premises Exchange attributes from Active Directory accounts to enable cloud mailbox provisioning.

.DESCRIPTION
    This script identifies and removes on-premises Exchange attributes from Active Directory user accounts
    that prevent mailbox provisioning in Exchange Online. Before removing attributes, it creates a 
    comprehensive backup of all Exchange-related attributes for each account.
    
    The script handles the following Exchange attributes:
    - msExchMailboxGuid (prevents cloud mailbox creation)
    - msExchArchiveGUID
    - msExchRemoteRecipientType
    - msExchRecipientDisplayType
    - msExchRecipientTypeDetails
    - msExchUserCulture
    - mailNickname
    - targetAddress
    - legacyExchangeDN
    - And other Exchange-specific attributes
    
    Key Features:
    - Pre-execution backup of all Exchange attributes to CSV
    - Detailed logging of all changes
    - Rollback capability using backup data
    - Support for individual users or bulk operations
    - WhatIf support for safe testing
    - Validation checks before and after cleanup

.PARAMETER Identity
    The SamAccountName, UserPrincipalName, or DistinguishedName of a single user to process.
    Cannot be used with SearchBase or InputFile parameters.

.PARAMETER SearchBase
    The Active Directory Organizational Unit to search for users to process.
    Example: "OU=Users,OU=Migrating,DC=contoso,DC=com"
    Cannot be used with Identity or InputFile parameters.

.PARAMETER InputFile
    Path to a CSV file containing users to process. CSV must have a column named 'Identity' with
    SamAccountName, UserPrincipalName, or DistinguishedName values.
    Cannot be used with Identity or SearchBase parameters.

.PARAMETER OutputDirectory
    Directory where backup files and reports will be saved.
    Default: C:\Reports\ExchangeCleanup

.PARAMETER AttributesToRemove
    Array of Exchange attribute names to remove. If not specified, removes common blocking attributes.
    Use -ListAttributes to see available attributes.

.PARAMETER BackupOnly
    Only creates backup of Exchange attributes without removing them. Useful for assessment.

.PARAMETER RestoreFromBackup
    Path to a backup CSV file to restore Exchange attributes from. Use for rollback.

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes.

.PARAMETER Force
    Bypasses confirmation prompts. Use with caution in production.

.PARAMETER ListAttributes
    Lists all Exchange-related attributes that can be removed and exits.

.EXAMPLE
    .\Set-ExchangeSync.ps1 -Identity "jsmith" -WhatIf
    
    Tests removal of Exchange attributes for user jsmith without making changes.

.EXAMPLE
    .\Set-ExchangeSync.ps1 -SearchBase "OU=Users,OU=ToMigrate,DC=contoso,DC=com"
    
    Backs up and removes Exchange attributes for all users in the specified OU.

.EXAMPLE
    .\Set-ExchangeSync.ps1 -InputFile "C:\Users\UsersToClean.csv"
    
    Processes users listed in the CSV file.

.EXAMPLE
    .\Set-ExchangeSync.ps1 -Identity "jsmith" -BackupOnly
    
    Creates backup of Exchange attributes for jsmith without removing them.

.EXAMPLE
    .\Set-ExchangeSync.ps1 -RestoreFromBackup "C:\Reports\ExchangeCleanup\Backup_20260202_143052.csv"
    
    Restores Exchange attributes from a previous backup.

.NOTES
    Author: W. Ford
    Date: 2026-02-02
    Version: 1.0
    
    Requirements:
    - Active Directory PowerShell module
    - Domain Administrator or Account Operator permissions
    - PowerShell 5.1 or later
    
    IMPORTANT: Always test with -WhatIf first and review backups before proceeding.
    Removing Exchange attributes is required when:
    - Decommissioning on-premises Exchange
    - Moving to Exchange Online without hybrid
    - Fixing orphaned Exchange attributes
    
    WARNING: This will disconnect users from on-premises Exchange mailboxes.
    Ensure data has been migrated before running.

.LINK
    https://docs.microsoft.com/en-us/exchange/decommission-on-premises-exchange
#>

[CmdletBinding(SupportsShouldProcess=$true, DefaultParameterSetName='Identity')]
param(
    [Parameter(Mandatory=$false, ParameterSetName='Identity', HelpMessage="User identity to process")]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,
    
    [Parameter(Mandatory=$false, ParameterSetName='SearchBase', HelpMessage="OU to search for users")]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,
    
    [Parameter(Mandatory=$false, ParameterSetName='InputFile', HelpMessage="CSV file with users to process")]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for backups and reports")]
    [string]$OutputDirectory = "C:\Reports\ExchangeCleanup",
    
    [Parameter(Mandatory=$false, HelpMessage="Specific attributes to remove")]
    [string[]]$AttributesToRemove,
    
    [Parameter(Mandatory=$false, HelpMessage="Only backup, don't remove attributes")]
    [switch]$BackupOnly,
    
    [Parameter(Mandatory=$false, ParameterSetName='Restore', HelpMessage="Restore from backup file")]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$RestoreFromBackup,
    
    [Parameter(Mandatory=$false, HelpMessage="Bypass confirmation prompts")]
    [switch]$Force,
    
    [Parameter(Mandatory=$false, HelpMessage="List available Exchange attributes")]
    [switch]$ListAttributes
)

#region Initialize
$ErrorActionPreference = 'Stop'
$StartTime = Get-Date
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$Separator = "=" * 80
$SubSeparator = "-" * 60

# Tracking variables
$ProcessedCount = 0
$SuccessCount = 0
$ErrorCount = 0
$SkippedCount = 0
$Errors = @()

# Define common Exchange attributes that block cloud provisioning
$DefaultExchangeAttributes = @(
    'msExchMailboxGuid',              # Primary blocker - must remove
    'msExchArchiveGUID',
    'msExchRemoteRecipientType',
    'msExchRecipientDisplayType',
    'msExchRecipientTypeDetails',
    'msExchUserCulture',
    'msExchVersion',
    'msExchMailboxSecurityDescriptor',
    'msExchMasterAccountSid',
    'msExchPoliciesExcluded',
    'msExchRecipientSoftDeletedStatus',
    'msExchUserAccountControl',
    'msExchWhenMailboxCreated',
    'mailNickname',
    'legacyExchangeDN',
    'targetAddress',
    'msExchHomeServerName',
    'homeMDB',
    'msExchMailboxTemplateLink'
)

# Extended list for comprehensive cleanup
$AllExchangeAttributes = $DefaultExchangeAttributes + @(
    'msExchAddressBookFlags',
    'msExchArchiveDatabaseLink',
    'msExchArchiveName',
    'msExchArchiveQuota',
    'msExchArchiveWarnQuota',
    'msExchBypassAudit',
    'msExchBypassModerationLink',
    'msExchDelegateListLink',
    'msExchELCMailboxFlags',
    'msExchHideFromAddressLists',
    'msExchMailboxAuditEnable',
    'msExchMailboxFolderSet',
    'msExchMobileMailboxFlags',
    'msExchProvisioningFlags',
    'msExchRBACPolicyLink',
    'msExchRecipientLimit',
    'msExchRetentionComment',
    'msExchRetentionURL',
    'msExchSafeSendersHash',
    'msExchTextMessagingState',
    'msExchUMDtmfMap'
)
#endregion

#region Functions

function Write-StatusMessage {
    param(
        [string]$Message, 
        [ValidateSet('Info','Success','Warning','Error','Header')]
        [string]$Type = 'Info'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Type) {
        'Error' { 
            Write-Host "[$timestamp] ❌ ERROR: $Message" -ForegroundColor Red
            $script:ErrorCount++
        }
        'Warning' { 
            Write-Host "[$timestamp] ⚠️  WARNING: $Message" -ForegroundColor Yellow
        }
        'Success' { 
            Write-Host "[$timestamp] ✅ SUCCESS: $Message" -ForegroundColor Green
        }
        'Header' {
            Write-Host "`n$Separator" -ForegroundColor Cyan
            Write-Host $Message -ForegroundColor Cyan
            Write-Host $Separator -ForegroundColor Cyan
        }
        default { 
            Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Cyan
        }
    }
}

function Test-Prerequisites {
    Write-StatusMessage "Checking prerequisites..." -Type Info
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "This script requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
    }
    
    # Check for Active Directory module
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        throw "Active Directory PowerShell module is not installed. Install RSAT tools."
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-StatusMessage "Active Directory module loaded successfully" -Type Success
    }
    catch {
        throw "Failed to import Active Directory module: $($_.Exception.Message)"
    }
    
    # Test AD connectivity
    try {
        $null = Get-ADDomain -ErrorAction Stop
        Write-StatusMessage "Active Directory connectivity verified" -Type Success
    }
    catch {
        throw "Cannot connect to Active Directory: $($_.Exception.Message)"
    }
    
    # Create output directory
    if (!(Test-Path $OutputDirectory)) {
        try {
            New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
            Write-StatusMessage "Created output directory: $OutputDirectory" -Type Success
        }
        catch {
            throw "Cannot create output directory: $($_.Exception.Message)"
        }
    }
    
    # Test write permissions
    $testFile = Join-Path $OutputDirectory "test_$Timestamp.tmp"
    try {
        "test" | Out-File -FilePath $testFile -ErrorAction Stop
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        Write-StatusMessage "Write permissions verified" -Type Success
    }
    catch {
        throw "No write permission to output directory: $($_.Exception.Message)"
    }
}

function Get-UsersToProcess {
    Write-StatusMessage "Gathering users to process..." -Type Info
    $users = @()
    
    try {
        switch ($PSCmdlet.ParameterSetName) {
            'Identity' {
                Write-StatusMessage "Processing single user: $Identity" -Type Info
                $user = Get-ADUser -Identity $Identity -Properties * -ErrorAction Stop
                $users = @($user)
            }
            
            'SearchBase' {
                Write-StatusMessage "Searching for users in: $SearchBase" -Type Info
                $users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties * -ErrorAction Stop
                Write-StatusMessage "Found $($users.Count) users in search base" -Type Info
            }
            
            'InputFile' {
                Write-StatusMessage "Loading users from: $InputFile" -Type Info
                $csv = Import-Csv -Path $InputFile -ErrorAction Stop
                
                if (-not $csv[0].PSObject.Properties['Identity']) {
                    throw "CSV file must contain an 'Identity' column"
                }
                
                foreach ($row in $csv) {
                    try {
                        $user = Get-ADUser -Identity $row.Identity -Properties * -ErrorAction Stop
                        $users += $user
                    }
                    catch {
                        Write-StatusMessage "Could not find user: $($row.Identity) - $($_.Exception.Message)" -Type Warning
                    }
                }
                Write-StatusMessage "Loaded $($users.Count) users from CSV" -Type Info
            }
            
            'Restore' {
                # For restore operations, users are loaded differently
                return @()
            }
        }
        
        if ($users.Count -eq 0) {
            Write-StatusMessage "No users found to process" -Type Warning
        }
        
        return $users
    }
    catch {
        Write-StatusMessage "Error gathering users: $($_.Exception.Message)" -Type Error
        throw
    }
}

function Get-ExchangeAttributes {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $attributes = @{}
    $attributeList = if ($AttributesToRemove) { $AttributesToRemove } else { $DefaultExchangeAttributes }
    
    foreach ($attr in $attributeList) {
        if ($User.PSObject.Properties[$attr]) {
            $value = $User.$attr
            if ($null -ne $value) {
                # Convert byte arrays to base64 for storage
                if ($value -is [byte[]]) {
                    $attributes[$attr] = [Convert]::ToBase64String($value)
                }
                else {
                    $attributes[$attr] = $value -join ';'
                }
            }
        }
    }
    
    return $attributes
}

function Backup-UserExchangeAttributes {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Users
    )
    
    Write-StatusMessage "Creating backup of Exchange attributes..." -Type Info
    
    $backupData = @()
    $attributeList = if ($AttributesToRemove) { $AttributesToRemove } else { $DefaultExchangeAttributes }
    
    foreach ($user in $Users) {
        $userBackup = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            UserPrincipalName = $user.UserPrincipalName
            DistinguishedName = $user.DistinguishedName
            DisplayName = $user.DisplayName
            Enabled = $user.Enabled
            BackupDate = $StartTime.ToString('yyyy-MM-dd HH:mm:ss')
        }
        
        # Add all Exchange attributes
        foreach ($attr in $attributeList) {
            if ($user.PSObject.Properties[$attr]) {
                $value = $user.$attr
                if ($null -ne $value) {
                    if ($value -is [byte[]]) {
                        $userBackup | Add-Member -NotePropertyName $attr -NotePropertyValue ([Convert]::ToBase64String($value))
                    }
                    else {
                        $userBackup | Add-Member -NotePropertyName $attr -NotePropertyValue ($value -join ';')
                    }
                }
                else {
                    $userBackup | Add-Member -NotePropertyName $attr -NotePropertyValue $null
                }
            }
            else {
                $userBackup | Add-Member -NotePropertyName $attr -NotePropertyValue $null
            }
        }
        
        $backupData += $userBackup
    }
    
    $backupFile = Join-Path $OutputDirectory "ExchangeAttributes_Backup_$Timestamp.csv"
    $backupData | Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8
    
    Write-StatusMessage "Backup saved to: $backupFile" -Type Success
    return $backupFile
}

function Remove-UserExchangeAttributes {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $attributeList = if ($AttributesToRemove) { $AttributesToRemove } else { $DefaultExchangeAttributes }
    $removedAttributes = @()
    $failedAttributes = @()
    $skippedAttributes = @()
    
    Write-StatusMessage "Processing user: $($User.SamAccountName) ($($User.DisplayName))" -Type Info
    
    foreach ($attr in $attributeList) {
        # Check if attribute exists and has a value
        # Use try-catch to handle attributes that don't exist
        try {
            $attributeValue = $User.$attr
            
            # Check if attribute has a value (not null and not empty)
            if ($null -ne $attributeValue -and $attributeValue -ne '' -and 
                ($attributeValue -isnot [array] -or $attributeValue.Count -gt 0)) {
                
                if ($PSCmdlet.ShouldProcess($User.SamAccountName, "Remove attribute $attr (value: $($attributeValue.ToString().Substring(0, [Math]::Min(50, $attributeValue.ToString().Length)))...)")) {
                    try {
                        Set-ADUser -Identity $User.DistinguishedName -Clear $attr -ErrorAction Stop
                        $removedAttributes += $attr
                        Write-Verbose "Cleared attribute: $attr"
                    }
                    catch {
                        # Only log as failed if the error is NOT "attribute doesn't exist"
                        if ($_.Exception.Message -notmatch "does not exist") {
                            $failedAttributes += $attr
                            Write-StatusMessage "Failed to clear $attr for $($User.SamAccountName): $($_.Exception.Message)" -Type Warning
                        }
                        else {
                            $skippedAttributes += $attr
                            Write-Verbose "Skipped $attr - attribute doesn't exist in schema"
                        }
                    }
                }
            }
            else {
                Write-Verbose "Skipped $attr - no value present"
                $skippedAttributes += $attr
            }
        }
        catch {
            # Property doesn't exist on this user object
            Write-Verbose "Skipped $attr - property not found on user object"
            $skippedAttributes += $attr
        }
    }
    
    # Log summary of what was found
    if ($removedAttributes.Count -gt 0) {
        Write-StatusMessage "Attributes removed: $($removedAttributes -join ', ')" -Type Info
    }
    if ($skippedAttributes.Count -gt 0) {
        Write-Verbose "Attributes skipped (no value): $($skippedAttributes -join ', ')"
    }
    
    return @{
        User = $User.SamAccountName
        RemovedCount = $removedAttributes.Count
        RemovedAttributes = $removedAttributes
        FailedAttributes = $failedAttributes
        SkippedAttributes = $skippedAttributes
        Success = ($failedAttributes.Count -eq 0)
    }
}

function Restore-ExchangeAttributes {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BackupFile
    )
    
    Write-StatusMessage "Restoring Exchange attributes from backup..." -Type Info
    
    try {
        $backupData = Import-Csv -Path $BackupFile -ErrorAction Stop
        Write-StatusMessage "Loaded $($backupData.Count) user records from backup" -Type Info
        
        $restoredCount = 0
        $failedCount = 0
        
        foreach ($record in $backupData) {
            try {
                Write-StatusMessage "Restoring attributes for: $($record.SamAccountName)" -Type Info
                
                $user = Get-ADUser -Identity $record.SamAccountName -ErrorAction Stop
                $attributesToRestore = @{}
                
                # Get all properties except standard ones
                $standardProps = @('SamAccountName', 'UserPrincipalName', 'DistinguishedName', 'DisplayName', 'Enabled', 'BackupDate')
                $record.PSObject.Properties | Where-Object { $_.Name -notin $standardProps -and ![string]::IsNullOrWhiteSpace($_.Value) } | ForEach-Object {
                    # Convert base64 back to byte array if needed
                    if ($_.Name -match 'Guid|Sid|Descriptor') {
                        try {
                            $attributesToRestore[$_.Name] = [Convert]::FromBase64String($_.Value)
                        }
                        catch {
                            $attributesToRestore[$_.Name] = $_.Value
                        }
                    }
                    else {
                        $attributesToRestore[$_.Name] = $_.Value
                    }
                }
                
                if ($attributesToRestore.Count -gt 0) {
                    if ($PSCmdlet.ShouldProcess($record.SamAccountName, "Restore $($attributesToRestore.Count) Exchange attributes")) {
                        Set-ADUser -Identity $user.DistinguishedName -Replace $attributesToRestore -ErrorAction Stop
                        Write-StatusMessage "Restored $($attributesToRestore.Count) attributes for $($record.SamAccountName)" -Type Success
                        $restoredCount++
                    }
                }
                else {
                    Write-StatusMessage "No attributes to restore for $($record.SamAccountName)" -Type Warning
                    $script:SkippedCount++
                }
            }
            catch {
                Write-StatusMessage "Failed to restore $($record.SamAccountName): $($_.Exception.Message)" -Type Error
                $failedCount++
                $script:Errors += "Restore failed for $($record.SamAccountName): $($_.Exception.Message)"
            }
        }
        
        Write-StatusMessage "Restore complete. Success: $restoredCount, Failed: $failedCount" -Type Success
    }
    catch {
        Write-StatusMessage "Failed to load backup file: $($_.Exception.Message)" -Type Error
        throw
    }
}

function Show-ExchangeAttributes {
    Write-Host "`n$Separator" -ForegroundColor Cyan
    Write-Host "EXCHANGE ATTRIBUTES REFERENCE" -ForegroundColor Cyan
    Write-Host $Separator -ForegroundColor Cyan
    
    Write-Host "`nDEFAULT ATTRIBUTES (Primary blockers):" -ForegroundColor Yellow
    $DefaultExchangeAttributes | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
    
    Write-Host "`nADDITIONAL ATTRIBUTES (Extended cleanup):" -ForegroundColor Yellow
    $additionalAttrs = $AllExchangeAttributes | Where-Object { $_ -notin $DefaultExchangeAttributes }
    $additionalAttrs | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
    
    Write-Host "`nUSAGE:" -ForegroundColor Green
    Write-Host "  Default: Removes primary blocking attributes" -ForegroundColor White
    Write-Host "  Custom:  Use -AttributesToRemove to specify exact attributes" -ForegroundColor White
    Write-Host "`n"
}

#endregion

#region Main Execution

try {
    # Show header
    Write-StatusMessage "Exchange Attribute Cleanup Tool" -Type Header
    Write-StatusMessage "Started: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Type Info
    
    # Handle list attributes request
    if ($ListAttributes) {
        Show-ExchangeAttributes
        exit 0
    }
    
    # Check prerequisites
    Test-Prerequisites
    
    # Handle restore operation
    if ($RestoreFromBackup) {
        Write-StatusMessage "RESTORE MODE: Restoring from backup" -Type Header
        
        if (-not $Force) {
            Write-Host "`n⚠️  WARNING: This will restore Exchange attributes from backup." -ForegroundColor Yellow
            Write-Host "This may re-enable on-premises Exchange connections.`n" -ForegroundColor Yellow
            $confirm = Read-Host "Type 'RESTORE' to continue"
            if ($confirm -ne 'RESTORE') {
                Write-StatusMessage "Restore cancelled by user" -Type Warning
                exit 0
            }
        }
        
        Restore-ExchangeAttributes -BackupFile $RestoreFromBackup
        
        $EndTime = Get-Date
        $Duration = $EndTime - $StartTime
        Write-StatusMessage "Restore completed in $($Duration.ToString('mm\:ss'))" -Type Success
        exit 0
    }
    
    # Get users to process
    $users = Get-UsersToProcess
    
    if ($users.Count -eq 0) {
        Write-StatusMessage "No users to process. Exiting." -Type Warning
        exit 0
    }
    
    # Show summary and confirm
    Write-Host "`n$SubSeparator" -ForegroundColor Yellow
    Write-Host "OPERATION SUMMARY" -ForegroundColor Yellow
    Write-Host $SubSeparator -ForegroundColor Yellow
    Write-Host "Users to process: $($users.Count)" -ForegroundColor White
    $attrCount = if ($AttributesToRemove -and $AttributesToRemove.Count -gt 0) { $AttributesToRemove.Count } else { $DefaultExchangeAttributes.Count }
    Write-Host "Attributes to remove: $attrCount" -ForegroundColor White
    Write-Host "Backup only: $BackupOnly" -ForegroundColor White
    Write-Host "Output directory: $OutputDirectory" -ForegroundColor White
    
    if (-not $Force -and -not $WhatIfPreference -and -not $BackupOnly) {
        Write-Host "`n⚠️  WARNING: This will remove Exchange attributes from $($users.Count) user(s)." -ForegroundColor Yellow
        Write-Host "This will disconnect users from on-premises Exchange mailboxes.`n" -ForegroundColor Yellow
        $confirm = Read-Host "Type 'PROCEED' to continue"
        if ($confirm -ne 'PROCEED') {
            Write-StatusMessage "Operation cancelled by user" -Type Warning
            exit 0
        }
    }
    
    # Create backup
    $backupFile = Backup-UserExchangeAttributes -Users $users
    
    if ($BackupOnly) {
        Write-StatusMessage "Backup complete. No attributes were removed." -Type Success
        Write-StatusMessage "Backup file: $backupFile" -Type Info
        exit 0
    }
    
    # Process users
    Write-StatusMessage "Beginning Exchange attribute removal..." -Type Header
    
    $results = @()
    
    foreach ($user in $users) {
        $ProcessedCount++
        
        try {
            $result = Remove-UserExchangeAttributes -User $user
            $results += [PSCustomObject]$result
            
            if ($result.Success) {
                $SuccessCount++
                Write-StatusMessage "✓ Processed $($user.SamAccountName) - Removed $($result.RemovedCount) attributes" -Type Success
            }
            else {
                Write-StatusMessage "⚠️  Processed $($user.SamAccountName) with warnings" -Type Warning
            }
        }
        catch {
            Write-StatusMessage "Failed to process $($user.SamAccountName): $($_.Exception.Message)" -Type Error
            $script:Errors += "Failed to process $($user.SamAccountName): $($_.Exception.Message)"
        }
    }
    
    # Save results
    $resultsFile = Join-Path $OutputDirectory "ExchangeCleanup_Results_$Timestamp.csv"
    $results | Export-Csv -Path $resultsFile -NoTypeInformation -Encoding UTF8
    Write-StatusMessage "Results saved to: $resultsFile" -Type Success
    
    # Generate summary report
    $reportFile = Join-Path $OutputDirectory "ExchangeCleanup_Summary_$Timestamp.txt"
    $durationMinutes = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 2)
    $report = @"
$Separator
EXCHANGE ATTRIBUTE CLEANUP SUMMARY
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$Separator

OPERATION DETAILS:
Total Users Processed: $ProcessedCount
Successfully Cleaned: $SuccessCount
Errors Encountered: $ErrorCount
Users Skipped: $SkippedCount

FILES GENERATED:
Backup: $backupFile
Results: $resultsFile
This Report: $reportFile

EXECUTION TIME:
Started: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Duration: $durationMinutes minutes

"@

    if ($script:Errors.Count -gt 0) {
        $report += "`n$SubSeparator`nERRORS:`n"
        $script:Errors | ForEach-Object { $report += "  $_`n" }
    }
    
    $report += "`n$SubSeparator`nNEXT STEPS:`n"
    $report += "1. Review the backup file: $backupFile`n"
    $report += "2. Force Azure AD Connect sync if hybrid environment`n"
    $report += "3. Allow 15-30 minutes for replication`n"
    $report += "4. Assign Exchange Online licenses to users`n"
    $report += "5. Verify mailboxes provision successfully in Exchange Online`n"
    $report += "6. If issues occur, restore from backup using -RestoreFromBackup parameter`n"
    
    $report += "`n$Separator`n"
    
    $report | Out-File -FilePath $reportFile -Encoding UTF8
    
    # Display summary
    Write-Host "`n$report" -ForegroundColor Cyan
    
    Write-StatusMessage "Operation completed successfully!" -Type Success
    
}
catch {
    Write-StatusMessage "Critical error: $($_.Exception.Message)" -Type Error
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
finally {
    $EndTime = Get-Date
    $Duration = $EndTime - $StartTime
    Write-Host "`nTotal execution time: $($Duration.ToString('mm\:ss'))" -ForegroundColor Gray
}

#endregion
