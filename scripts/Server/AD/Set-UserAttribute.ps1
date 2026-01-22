<#
.SYNOPSIS
    Updates Active Directory user attributes based on UPN or username submission.

.DESCRIPTION
    This script provides a comprehensive solution for updating Active Directory user attributes.
    It supports finding users by either UserPrincipalName (UPN) or sAMAccountName (username)
    and can update any AD attribute, including standard and extended attributes.
    
    The script includes:
    - Flexible user identification (UPN or username)
    - Support for all standard AD attributes (DisplayName, Email, Department, etc.)
    - Support for extended attributes (extensionAttribute1-15)
    - Bulk update capability via CSV import
    - Interactive mode for single user updates
    - Validation and error handling
    - Detailed logging and reporting
    - Dry-run mode to preview changes before applying

.PARAMETER Identity
    The user identifier(s). Can be either UserPrincipalName (UPN) or sAMAccountName (username).
    Accepts a single user or an array of users.
    Examples: "john.doe@contoso.com" or "johndoe" or @("user1@contoso.com", "user2")

.PARAMETER AttributeName
    The name of the AD attribute(s) to update. Supports standard attributes (DisplayName, 
    Title, Department, etc.) and extended attributes (extensionAttribute1-15).
    Accepts a single attribute or an array of attributes.

.PARAMETER AttributeValue
    The new value(s) to set for the specified attribute. Use empty string "" to clear an attribute.
    When using arrays for AttributeName, provide matching array of values.

.PARAMETER CsvPath
    Path to a CSV file for bulk updates. CSV should have columns: Identity, AttributeName, AttributeValue.
    The Identity column can contain either UPNs or sAMAccountNames.

.PARAMETER UpdateData
    Array of objects for bulk updates. Each object should have properties: Identity, AttributeName, AttributeValue.
    This allows programmatic bulk updates without creating a CSV file.
    Accepts hashtables or PSCustomObjects.

.PARAMETER WhatIf
    Performs a dry-run showing what would be changed without making actual modifications.

.PARAMETER OutputDirectory
    Directory where reports and logs will be saved. Defaults to C:\Reports\AD_Exports

.PARAMETER SearchByUPN
    Forces the script to search only by UserPrincipalName. By default, the script auto-detects.

.PARAMETER SearchByUsername
    Forces the script to search only by sAMAccountName. By default, the script auto-detects.

.EXAMPLE
    .\Set-UserAttribute.ps1 -Identity "john.doe@contoso.com" -AttributeName "Department" -AttributeValue "IT"
    
    Updates the Department attribute for user john.doe@contoso.com to "IT".

.EXAMPLE
    .\Set-UserAttribute.ps1 -Identity "johndoe" -AttributeName "extensionAttribute1" -AttributeValue "Building-A"
    
    Updates extensionAttribute1 for user johndoe to "Building-A".

.EXAMPLE
    .\Set-UserAttribute.ps1 -Identity @("user1@contoso.com", "user2@contoso.com") -AttributeName "Department" -AttributeValue "IT"
    
    Updates the Department attribute to "IT" for multiple users.

.EXAMPLE
    .\Set-UserAttribute.ps1 -Identity "john.doe@contoso.com" -AttributeName @("Department", "Title") -AttributeValue @("IT", "Manager")
    
    Updates multiple attributes for a single user.

.EXAMPLE
    .\Set-UserAttribute.ps1 -Identity "john.doe@contoso.com" -AttributeName "Title" -AttributeValue "Senior Engineer" -WhatIf
    
    Shows what would be changed without making actual modifications.

.EXAMPLE
    .\Set-UserAttribute.ps1 -CsvPath "C:\Updates\UserAttributes.csv"
    
    Performs bulk updates from a CSV file with columns: Identity, AttributeName, AttributeValue.

.EXAMPLE
    $updates = @(
        @{Identity="user1@contoso.com"; AttributeName="Department"; AttributeValue="IT"},
        @{Identity="user2@contoso.com"; AttributeName="Title"; AttributeValue="Manager"},
        @{Identity="user3"; AttributeName="extensionAttribute1"; AttributeValue="Location-A"}
    )
    .\Set-UserAttribute.ps1 -UpdateData $updates
    
    Performs bulk updates using an array of hashtables.

.EXAMPLE
    $updates = Import-Csv "source.csv" | Where-Object {$_.Department -eq "IT"}
    .\Set-UserAttribute.ps1 -UpdateData $updates
    
    Filters CSV data and passes the filtered objects directly to the script.

.EXAMPLE
    .\Set-UserAttribute.ps1
    
    Runs in interactive mode, prompting for user identity and attribute updates.

.NOTES
    Author: W. Ford (Managed Solution LLC)
    Date: 2026-01-22
    Version: 1.0
    
    Requirements:
    - PowerShell 5.1 or later
    - Active Directory PowerShell module (RSAT-AD-PowerShell)
    - Appropriate AD permissions to modify user objects
    - Domain connectivity
    
    Common AD Attributes:
    Standard: DisplayName, GivenName, Surname, Title, Department, Company, Office, 
              StreetAddress, City, State, PostalCode, Country, EmailAddress, 
              OfficePhone, MobilePhone, Description, Manager
    
    Extended: extensionAttribute1 through extensionAttribute15
    
    Output:
    - Console progress messages with color coding
    - CSV export of changes made
    - Detailed log file with all operations

.LINK
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/set-aduser
.LINK
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-aduser
#>

[CmdletBinding(SupportsShouldProcess=$true, DefaultParameterSetName='Interactive')]
param(
    [Parameter(Mandatory=$false, ParameterSetName='Single', HelpMessage="User identifier(s) (UPN or username)")]
    [ValidateNotNullOrEmpty()]
    [string[]]$Identity,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single', HelpMessage="Attribute name(s) to update")]
    [ValidateNotNullOrEmpty()]
    [string[]]$AttributeName,
    
    [Parameter(Mandatory=$false, ParameterSetName='Single', HelpMessage="New value(s) for the attribute(s)")]
    [AllowEmptyString()]
    [string[]]$AttributeValue,
    
    [Parameter(Mandatory=$false, ParameterSetName='Bulk', HelpMessage="Path to CSV file for bulk updates")]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { throw "CSV file not found: $_" }
    })]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false, ParameterSetName='Array', HelpMessage="Array of objects for bulk updates")]
    [ValidateNotNullOrEmpty()]
    [array]$UpdateData,
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for reports and logs")]
    [string]$OutputDirectory = "C:\Reports\AD_Exports",
    
    [Parameter(Mandatory=$false, HelpMessage="Search only by UserPrincipalName")]
    [switch]$SearchByUPN,
    
    [Parameter(Mandatory=$false, HelpMessage="Search only by sAMAccountName")]
    [switch]$SearchByUsername
)

#Requires -Modules ActiveDirectory

# Initialize script variables
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$Script:ErrorCount = 0
$Script:WarningCount = 0
$Script:SuccessCount = 0
$Script:Changes = @()
$Separator = "=" * 80
$SubSeparator = "-" * 60

# Ensure output directory exists
if (!(Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
        Write-Host "✅ Created output directory: $OutputDirectory" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Cannot create output directory: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Initialize log file
$LogFile = Join-Path $OutputDirectory "AD_AttributeUpdate_Log_$Timestamp.txt"
$ChangesFile = Join-Path $OutputDirectory "AD_AttributeChanges_$Timestamp.csv"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logMessage
    
    # Write to console with color
    switch ($Level) {
        'Success' { 
            Write-Host "✅ $Message" -ForegroundColor Green
            $Script:SuccessCount++
        }
        'Warning' { 
            Write-Host "⚠️  $Message" -ForegroundColor Yellow
            $Script:WarningCount++
        }
        'Error' { 
            Write-Host "❌ $Message" -ForegroundColor Red
            $Script:ErrorCount++
        }
        'Info' { 
            Write-Host "ℹ️  $Message" -ForegroundColor Cyan
        }
    }
}

function Get-ADUserByIdentity {
    <#
    .SYNOPSIS
        Finds an AD user by UPN or username with flexible search logic.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        
        [switch]$ForceUPN,
        [switch]$ForceUsername
    )
    
    try {
        $user = $null
        
        # Determine search strategy
        if ($ForceUPN) {
            # Search by UPN only
            Write-Verbose "Searching by UPN: $Identity"
            $user = Get-ADUser -Filter "UserPrincipalName -eq '$Identity'" -Properties * -ErrorAction Stop
        }
        elseif ($ForceUsername) {
            # Search by sAMAccountName only
            Write-Verbose "Searching by sAMAccountName: $Identity"
            $user = Get-ADUser -Filter "sAMAccountName -eq '$Identity'" -Properties * -ErrorAction Stop
        }
        else {
            # Auto-detect: if contains @, likely UPN; otherwise username
            if ($Identity -like "*@*") {
                Write-Verbose "Detected UPN format, searching by UserPrincipalName: $Identity"
                $user = Get-ADUser -Filter "UserPrincipalName -eq '$Identity'" -Properties * -ErrorAction Stop
                
                # Fallback to username if UPN search fails
                if (-not $user) {
                    Write-Verbose "UPN search failed, trying as sAMAccountName: $Identity"
                    $user = Get-ADUser -Filter "sAMAccountName -eq '$Identity'" -Properties * -ErrorAction Stop
                }
            }
            else {
                Write-Verbose "Detected username format, searching by sAMAccountName: $Identity"
                $user = Get-ADUser -Filter "sAMAccountName -eq '$Identity'" -Properties * -ErrorAction Stop
                
                # Fallback to UPN if username search fails
                if (-not $user) {
                    Write-Verbose "sAMAccountName search failed, trying as UPN: $Identity"
                    $user = Get-ADUser -Filter "UserPrincipalName -eq '$Identity'" -Properties * -ErrorAction Stop
                }
            }
        }
        
        if ($user) {
            Write-Verbose "Found user: $($user.DisplayName) ($($user.sAMAccountName))"
            return $user
        }
        else {
            Write-Log "User not found: $Identity" -Level Warning
            return $null
        }
    }
    catch {
        Write-Log "Error searching for user '$Identity': $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Set-ADUserAttributeSafe {
    <#
    .SYNOPSIS
        Safely updates an AD user attribute with validation and logging.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        
        [Parameter(Mandatory=$true)]
        [string]$AttributeName,
        
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$AttributeValue,
        
        [switch]$WhatIfMode
    )
    
    try {
        # Get current value
        $currentValue = $User.$AttributeName
        if ($null -eq $currentValue) { $currentValue = "" }
        
        # Check if change is needed
        if ($currentValue -eq $AttributeValue) {
            Write-Log "No change needed for $($User.sAMAccountName) - $AttributeName already set to '$AttributeValue'" -Level Info
            return $false
        }
        
        # Record the change
        $changeRecord = [PSCustomObject]@{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Identity = $User.sAMAccountName
            UPN = $User.UserPrincipalName
            DisplayName = $User.DisplayName
            AttributeName = $AttributeName
            OldValue = $currentValue
            NewValue = $AttributeValue
            Status = "Pending"
            Error = ""
        }
        
        if ($WhatIfMode) {
            $changeRecord.Status = "WhatIf"
            Write-Log "WHATIF: Would update $($User.sAMAccountName) - $AttributeName from '$currentValue' to '$AttributeValue'" -Level Info
            $Script:Changes += $changeRecord
            return $true
        }
        
        # Map common attribute names to Set-ADUser parameters
        $parameterMap = @{
            'DisplayName' = 'DisplayName'
            'GivenName' = 'GivenName'
            'Surname' = 'Surname'
            'Title' = 'Title'
            'Department' = 'Department'
            'Company' = 'Company'
            'Office' = 'Office'
            'StreetAddress' = 'StreetAddress'
            'City' = 'City'
            'State' = 'State'
            'PostalCode' = 'PostalCode'
            'Country' = 'Country'
            'EmailAddress' = 'EmailAddress'
            'OfficePhone' = 'OfficePhone'
            'MobilePhone' = 'MobilePhone'
            'Description' = 'Description'
            'Manager' = 'Manager'
        }
        
        # Perform the update
        $setParams = @{
            Identity = $User.DistinguishedName
            ErrorAction = 'Stop'
        }
        
        if ($parameterMap.ContainsKey($AttributeName)) {
            # Use mapped parameter
            if ([string]::IsNullOrEmpty($AttributeValue)) {
                $setParams[$parameterMap[$AttributeName]] = $null
                $setParams['Clear'] = $parameterMap[$AttributeName]
            }
            else {
                $setParams[$parameterMap[$AttributeName]] = $AttributeValue
            }
        }
        else {
            # Use Replace/Clear for extended or other attributes
            if ([string]::IsNullOrEmpty($AttributeValue)) {
                $setParams['Clear'] = $AttributeName
            }
            else {
                $setParams['Replace'] = @{ $AttributeName = $AttributeValue }
            }
        }
        
        Set-ADUser @setParams
        
        $changeRecord.Status = "Success"
        Write-Log "Updated $($User.sAMAccountName) - ${AttributeName}: '$currentValue' → '$AttributeValue'" -Level Success
        $Script:Changes += $changeRecord
        return $true
    }
    catch {
        $changeRecord.Status = "Failed"
        $changeRecord.Error = $_.Exception.Message
        Write-Log "Failed to update $($User.sAMAccountName) - ${AttributeName}: $($_.Exception.Message)" -Level Error
        $Script:Changes += $changeRecord
        return $false
    }
}

function Show-InteractiveMenu {
    Write-Host "`n$Separator" -ForegroundColor Cyan
    Write-Host "AD USER ATTRIBUTE UPDATER - INTERACTIVE MODE" -ForegroundColor Cyan
    Write-Host $Separator -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This tool allows you to update AD user attributes by UPN or username." -ForegroundColor White
    Write-Host ""
    Write-Host "Common Attributes:" -ForegroundColor Yellow
    Write-Host "  Standard: DisplayName, GivenName, Surname, Title, Department, Company" -ForegroundColor Gray
    Write-Host "           Office, City, State, PostalCode, EmailAddress, OfficePhone" -ForegroundColor Gray
    Write-Host "  Extended: extensionAttribute1 through extensionAttribute15" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Enter 'Q' at any prompt to quit" -ForegroundColor Gray
    Write-Host "$SubSeparator`n" -ForegroundColor Cyan
}

function Start-InteractiveMode {
    Show-InteractiveMenu
    
    do {
        # Get user identity
        $userIdentity = Read-Host "`nEnter user identifier (UPN or username, Q to quit)"
        if ($userIdentity -eq 'Q') { return }
        
        # Find user
        $user = Get-ADUserByIdentity -Identity $userIdentity -ForceUPN:$SearchByUPN -ForceUsername:$SearchByUsername
        
        if (-not $user) {
            Write-Host "User not found. Please try again." -ForegroundColor Yellow
            continue
        }
        
        Write-Host "`n✅ Found user: $($user.DisplayName) ($($user.sAMAccountName))" -ForegroundColor Green
        Write-Host "   UPN: $($user.UserPrincipalName)" -ForegroundColor Gray
        Write-Host "   DN: $($user.DistinguishedName)" -ForegroundColor Gray
        
        # Get attribute to update
        $attrName = Read-Host "`nEnter attribute name to update (Q to quit)"
        if ($attrName -eq 'Q') { return }
        
        # Show current value if it exists
        $currentValue = $user.$attrName
        if ($null -ne $currentValue -and $currentValue -ne "") {
            Write-Host "   Current value: $currentValue" -ForegroundColor Gray
        }
        else {
            Write-Host "   Current value: (not set)" -ForegroundColor Gray
        }
        
        # Get new value
        $attrValue = Read-Host "Enter new value (empty to clear, Q to quit)"
        if ($attrValue -eq 'Q') { return }
        
        # Confirm
        Write-Host "`n⚠️  Confirm update:" -ForegroundColor Yellow
        Write-Host "   User: $($user.DisplayName) ($($user.sAMAccountName))" -ForegroundColor White
        Write-Host "   Attribute: $attrName" -ForegroundColor White
        Write-Host "   New Value: $attrValue" -ForegroundColor White
        
        $confirm = Read-Host "`nProceed with update? (Y/N)"
        
        if ($confirm -eq 'Y') {
            $null = Set-ADUserAttributeSafe -User $user -AttributeName $attrName -AttributeValue $attrValue -WhatIfMode:$WhatIfPreference
        }
        else {
            Write-Host "Update cancelled." -ForegroundColor Yellow
        }
        
        $continue = Read-Host "`nUpdate another attribute? (Y/N)"
        
    } while ($continue -eq 'Y')
}

function Start-ArrayUpdate {
    param([array]$UpdateData)
    
    Write-Log "Starting bulk update from array of $($UpdateData.Count) objects" -Level Info
    
    try {
        if (-not $UpdateData -or $UpdateData.Count -eq 0) {
            Write-Log "Update array is empty or invalid" -Level Error
            return
        }
        
        # Validate object structure
        $firstItem = $UpdateData[0]
        $requiredProperties = @('Identity', 'AttributeName', 'AttributeValue')
        
        # Check if it's a hashtable or PSCustomObject
        if ($firstItem -is [hashtable]) {
            $missingProperties = $requiredProperties | Where-Object { -not $firstItem.ContainsKey($_) }
        }
        else {
            $properties = $firstItem.PSObject.Properties.Name
            $missingProperties = $requiredProperties | Where-Object { $_ -notin $properties }
        }
        
        if ($missingProperties) {
            Write-Log "Update objects missing required properties: $($missingProperties -join ', ')" -Level Error
            Write-Log "Required properties: $($requiredProperties -join ', ')" -Level Info
            return
        }
        
        Write-Log "Processing $($UpdateData.Count) updates from array" -Level Info
        
        $progressCount = 0
        foreach ($update in $UpdateData) {
            $progressCount++
            Write-Progress -Activity "Processing bulk updates from array" -Status "Update $progressCount of $($UpdateData.Count)" -PercentComplete (($progressCount / $UpdateData.Count) * 100)
            
            # Handle both hashtable and object access
            $identity = if ($update -is [hashtable]) { $update['Identity'] } else { $update.Identity }
            $attrName = if ($update -is [hashtable]) { $update['AttributeName'] } else { $update.AttributeName }
            $attrValue = if ($update -is [hashtable]) { $update['AttributeValue'] } else { $update.AttributeValue }
            
            if ([string]::IsNullOrWhiteSpace($identity)) {
                Write-Log "Skipping item $progressCount - empty Identity" -Level Warning
                continue
            }
            
            $user = Get-ADUserByIdentity -Identity $identity -ForceUPN:$SearchByUPN -ForceUsername:$SearchByUsername
            
            if ($user) {
                $null = Set-ADUserAttributeSafe -User $user -AttributeName $attrName -AttributeValue $attrValue -WhatIfMode:$WhatIfPreference
            }
        }
        
        Write-Progress -Activity "Processing bulk updates from array" -Completed
        
    }
    catch {
        Write-Log "Error processing array: $($_.Exception.Message)" -Level Error
    }
}

function Start-BulkUpdate {
    param([string]$CsvPath)
    
    Write-Log "Starting bulk update from CSV: $CsvPath" -Level Info
    
    try {
        $updates = Import-Csv -Path $CsvPath -ErrorAction Stop
        
        if (-not $updates) {
            Write-Log "CSV file is empty or invalid" -Level Error
            return
        }
        
        # Validate CSV structure
        $requiredColumns = @('Identity', 'AttributeName', 'AttributeValue')
        $csvColumns = $updates[0].PSObject.Properties.Name
        
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
        if ($missingColumns) {
            Write-Log "CSV missing required columns: $($missingColumns -join ', ')" -Level Error
            Write-Log "Required columns: $($requiredColumns -join ', ')" -Level Info
            return
        }
        
        Write-Log "Processing $($updates.Count) updates from CSV" -Level Info
        
        $progressCount = 0
        foreach ($update in $updates) {
            $progressCount++
            Write-Progress -Activity "Processing bulk updates" -Status "Update $progressCount of $($updates.Count)" -PercentComplete (($progressCount / $updates.Count) * 100)
            
            if ([string]::IsNullOrWhiteSpace($update.Identity)) {
                Write-Log "Skipping row $progressCount - empty Identity" -Level Warning
                continue
            }
            
            $user = Get-ADUserByIdentity -Identity $update.Identity -ForceUPN:$SearchByUPN -ForceUsername:$SearchByUsername
            
            if ($user) {
                $null = Set-ADUserAttributeSafe -User $user -AttributeName $update.AttributeName -AttributeValue $update.AttributeValue -WhatIfMode:$WhatIfPreference
            }
        }
        
        Write-Progress -Activity "Processing bulk updates" -Completed
        
    }
    catch {
        Write-Log "Error processing CSV: $($_.Exception.Message)" -Level Error
    }
}

#region Main Script Execution

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "AD USER ATTRIBUTE UPDATER" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

# Check Active Directory module
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    Write-Host "❌ Active Directory module not found" -ForegroundColor Red
    Write-Host "   Install with: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
    exit 1
}

# Import AD module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "Active Directory module loaded" -Level Info
}
catch {
    Write-Host "❌ Failed to import Active Directory module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Verify domain connectivity
try {
    $domain = Get-ADDomain -ErrorAction Stop
    Write-Log "Connected to domain: $($domain.DNSRoot)" -Level Info
}
catch {
    Write-Log "Cannot connect to Active Directory domain: $($_.Exception.Message)" -Level Error
    exit 1
}

# Execute based on parameter set
switch ($PSCmdlet.ParameterSetName) {
    'Single' {
        if ($Identity -and $AttributeName) {
            # Validate array lengths if multiple attributes provided
            if ($AttributeName.Count -gt 1) {
                if ($AttributeValue.Count -ne $AttributeName.Count) {
                    Write-Log "When providing multiple attributes, you must provide matching number of values" -Level Error
                    Write-Log "AttributeNames: $($AttributeName.Count), AttributeValues: $($AttributeValue.Count)" -Level Error
                    exit 1
                }
            }
            
            # Process array of users
            $userCount = 0
            foreach ($userIdentity in $Identity) {
                $userCount++
                
                if ($Identity.Count -gt 1) {
                    Write-Host "\n[$userCount of $($Identity.Count)] Processing user: $userIdentity" -ForegroundColor Cyan
                }
                
                $user = Get-ADUserByIdentity -Identity $userIdentity -ForceUPN:$SearchByUPN -ForceUsername:$SearchByUsername
                
                if ($user) {
                    # Process single or multiple attributes for this user
                    for ($i = 0; $i -lt $AttributeName.Count; $i++) {
                        $attrName = $AttributeName[$i]
                        $attrValue = if ($AttributeValue.Count -gt 1) { $AttributeValue[$i] } else { $AttributeValue[0] }
                        
                        $null = Set-ADUserAttributeSafe -User $user -AttributeName $attrName -AttributeValue $attrValue -WhatIfMode:$WhatIfPreference
                    }
                }
            }
        }
        else {
            # Missing required parameters, switch to interactive
            Start-InteractiveMode
        }
    }
    'Bulk' {
        # Bulk update from CSV
        Start-BulkUpdate -CsvPath $CsvPath
    }
    'Array' {
        # Bulk update from array of objects
        Start-ArrayUpdate -UpdateData $UpdateData
    }
    'Interactive' {
        # Interactive mode
        Start-InteractiveMode
    }
}

# Export changes to CSV
if ($Script:Changes.Count -gt 0) {
    try {
        $Script:Changes | Export-Csv -Path $ChangesFile -NoTypeInformation -Encoding UTF8
        Write-Log "Exported $($Script:Changes.Count) change records to: $ChangesFile" -Level Info
    }
    catch {
        Write-Log "Failed to export changes: $($_.Exception.Message)" -Level Error
    }
}

# Summary
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Successful Updates: $Script:SuccessCount" -ForegroundColor Green
Write-Host "Warnings: $Script:WarningCount" -ForegroundColor Yellow
Write-Host "Errors: $Script:ErrorCount" -ForegroundColor Red
Write-Host "`nLog file: $LogFile" -ForegroundColor Gray
if ($Script:Changes.Count -gt 0) {
    Write-Host "Changes file: $ChangesFile" -ForegroundColor Gray
}
Write-Host "$Separator`n" -ForegroundColor Cyan

#endregion
