<#
.SYNOPSIS
    Initialize OneDrive for Business accounts for a list of users via Microsoft Graph API.

.DESCRIPTION
    This script provisions OneDrive for Business storage for Microsoft 365 users by accessing
    their drive endpoint through Microsoft Graph API. OneDrive provisioning typically happens
    automatically when a user first accesses OneDrive, but this script forces provisioning
    ahead of time for bulk user migrations or new account setups.
    
    The script supports running in Azure Cloud Shell and local PowerShell environments, with
    automatic detection and appropriate output path selection.
    
    Key Features:
    - Bulk OneDrive initialization from an array of email addresses
    - Automatic Cloud Shell detection with persistent storage paths
    - Detailed progress tracking with colored console output
    - CSV export with timestamps for audit and verification
    - Error tracking with success/failure counts
    - Retry logic for transient failures

.PARAMETER EmailAddresses
    Array of user email addresses (UPNs) for OneDrive initialization.
    Example: @('user1@contoso.com', 'user2@contoso.com')

.PARAMETER OutputDirectory
    Directory path for result CSV exports. 
    Defaults to ~/clouddrive/Reports in Cloud Shell or C:\Reports\OneDrive_Exports locally.

.PARAMETER RetryCount
    Number of retry attempts for failed provisioning operations. Default: 2

.PARAMETER DelaySeconds
    Delay in seconds between provisioning operations to avoid throttling. Default: 2

.EXAMPLE
    .\Start-OneDriveProvisioning.ps1 -EmailAddresses @('user1@contoso.com', 'user2@contoso.com')
    
    Initializes OneDrive for two users using default settings.

.EXAMPLE
    $users = @('asmith@contoso.com', 'bjones@contoso.com', 'cwilliams@contoso.com')
    .\Start-OneDriveProvisioning.ps1 -EmailAddresses $users -DelaySeconds 3
    
    Initializes OneDrive for three users with a 3-second delay between operations.

.EXAMPLE
    # Running in Azure Cloud Shell
    $users = Get-Content ~/clouddrive/users.txt
    .\Start-OneDriveProvisioning.ps1 -EmailAddresses $users -OutputDirectory ~/clouddrive/Reports

.EXAMPLE
    # Direct invocation from GitHub in Cloud Shell
    $users = @('user1@contoso.com', 'user2@contoso.com')
    $scriptContent = Invoke-RestMethod 'https://raw.githubusercontent.com/Managed-Solution-LLC/PowerShellEveryting/main/scripts/Office365/Start-OneDriveProvisioning.ps1'
    $scriptBlock = [ScriptBlock]::Create($scriptContent)
    & $scriptBlock -EmailAddresses $users

.NOTES
    Author: W. Ford
    Date: 2026-01-29
    Version: 1.0
    
    Requirements:
    - Microsoft.Graph.Users module (2.x or later)
    - Microsoft.Graph.Files module
    - Microsoft.Graph.Sites module
    - Microsoft Graph permissions: User.Read.All, Files.ReadWrite.All, Sites.ReadWrite.All
    - PowerShell 5.1 or later (PowerShell 7+ recommended for Cloud Shell)
    
    The script connects to Microsoft Graph and accesses each user's drive endpoint. This
    action triggers OneDrive provisioning if not already initialized. The process can take
    several seconds per user and may fail for users without licenses or with OneDrive disabled.
    
    Output CSV includes: EmailAddress, OneDriveProvisioned, DriveId, Status, ErrorMessage, Timestamp

.LINK
    https://docs.microsoft.com/en-us/graph/api/drive-get
.LINK
    https://docs.microsoft.com/en-us/onedrive/pre-provision-accounts
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Array of user email addresses (UPNs) to provision OneDrive for")]
    [ValidateNotNullOrEmpty()]
    [string[]]$EmailAddresses,
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for results CSV")]
    [string]$OutputDirectory,
    
    [Parameter(Mandatory=$false, HelpMessage="Number of retry attempts for failed operations")]
    [ValidateRange(0, 5)]
    [int]$RetryCount = 2,
    
    [Parameter(Mandatory=$false, HelpMessage="Delay in seconds between provisioning operations")]
    [ValidateRange(0, 30)]
    [int]$DelaySeconds = 2
)

#region Functions

function Test-CloudShell {
    <#
    .SYNOPSIS
        Detects if the script is running in Azure Cloud Shell.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    # Check for Cloud Shell environment variables
    $cloudShellEnvVars = @(
        'ACC_CLOUD',
        'ACC_LOCATION',
        'AZUREPS_HOST_ENVIRONMENT'
    )
    
    $hasCloudShellVars = $cloudShellEnvVars | Where-Object { 
        $null -ne [System.Environment]::GetEnvironmentVariable($_) 
    }
    
    # Check for clouddrive mount
    $hasCloudDrive = Test-Path "$HOME/clouddrive" -PathType Container
    
    return ($hasCloudShellVars.Count -gt 0 -or $hasCloudDrive)
}

function Get-DefaultOutputDirectory {
    <#
    .SYNOPSIS
        Returns the appropriate default output directory based on environment.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    if (Test-CloudShell) {
        return "$HOME/clouddrive/Reports"
    }
    else {
        return "C:\Reports\OneDrive_Exports"
    }
}

function Write-StatusMessage {
    <#
    .SYNOPSIS
        Writes formatted status messages with timestamps and color coding.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Type) {
        'Success' { 
            Write-Host "[$timestamp] ✅ $Message" -ForegroundColor Green
        }
        'Warning' { 
            Write-Host "[$timestamp] ⚠️  $Message" -ForegroundColor Yellow
            $script:WarningCount++
        }
        'Error' { 
            Write-Host "[$timestamp] ❌ $Message" -ForegroundColor Red
            $script:ErrorCount++
        }
        default { 
            Write-Host "[$timestamp] ℹ️  $Message" -ForegroundColor Cyan
        }
    }
}

function Test-RequiredModules {
    <#
    .SYNOPSIS
        Verifies all required Microsoft Graph modules are installed.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    $RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Files',
        'Microsoft.Graph.Sites'
    )
    
    $allInstalled = $true
    
    foreach ($moduleName in $RequiredModules) {
        if (-not (Get-Module -Name $moduleName -ListAvailable)) {
            Write-StatusMessage "Required module '$moduleName' is not installed" -Type Error
            Write-Host "   Install with: Install-Module $moduleName -Scope CurrentUser" -ForegroundColor Yellow
            $allInstalled = $false
        }
        else {
            Write-Verbose "Module $moduleName is available"
        }
    }
    
    return $allInstalled
}

function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with required permissions.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    try {
        Write-StatusMessage "Connecting to Microsoft Graph..." -Type Info
        
        # Import required modules
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Import-Module Microsoft.Graph.Files -ErrorAction Stop
        Import-Module Microsoft.Graph.Sites -ErrorAction Stop
        
        # Connect with required scopes
        $scopes = @(
            'User.Read.All',
            'Files.ReadWrite.All',
            'Sites.ReadWrite.All'
        )
        
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if ($context) {
            Write-StatusMessage "Connected to Microsoft Graph as $($context.Account)" -Type Success
            Write-StatusMessage "Tenant: $($context.TenantId)" -Type Info
            return $true
        }
        else {
            Write-StatusMessage "Failed to verify Microsoft Graph connection" -Type Error
            return $false
        }
    }
    catch {
        Write-StatusMessage "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Initialize-OneDriveForUser {
    <#
    .SYNOPSIS
        Initializes OneDrive for a specific user by accessing their drive endpoint.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory=$false)]
        [int]$RetryAttempts = 2
    )
    
    $result = [PSCustomObject]@{
        EmailAddress = $EmailAddress
        OneDriveProvisioned = $false
        DriveId = $null
        Status = 'Failed'
        ErrorMessage = $null
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
    
    # Get user ID
    try {
        Write-Verbose "Looking up user: $EmailAddress"
        $user = Get-MgUser -UserId $EmailAddress -ErrorAction Stop
        
        if (-not $user) {
            $result.ErrorMessage = "User not found"
            Write-StatusMessage "User not found: $EmailAddress" -Type Warning
            return $result
        }
    }
    catch {
        $result.ErrorMessage = "User lookup failed: $($_.Exception.Message)"
        Write-StatusMessage "Failed to lookup user $EmailAddress : $($_.Exception.Message)" -Type Error
        return $result
    }
    
    # Attempt to access/provision OneDrive with retry logic
    $attempt = 0
    $success = $false
    
    while ($attempt -le $RetryAttempts -and -not $success) {
        $attempt++
        
        try {
            Write-Verbose "Provisioning attempt $attempt for $EmailAddress"
            
            # Access the user's drive to trigger provisioning
            $drive = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop
            
            if ($drive) {
                $result.OneDriveProvisioned = $true
                $result.DriveId = $drive.Id
                $result.Status = 'Success'
                $result.ErrorMessage = $null
                $success = $true
                
                Write-StatusMessage "OneDrive provisioned for $EmailAddress (Drive ID: $($drive.Id))" -Type Success
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            if ($attempt -le $RetryAttempts) {
                Write-StatusMessage "Retry $attempt failed for $EmailAddress : $errorMsg" -Type Warning
                Start-Sleep -Seconds 2
            }
            else {
                $result.ErrorMessage = $errorMsg
                Write-StatusMessage "Failed to provision OneDrive for $EmailAddress after $RetryAttempts retries: $errorMsg" -Type Error
            }
        }
    }
    
    return $result
}

#endregion

#region Main Script

# Initialize tracking variables
$StartTime = Get-Date
$ErrorCount = 0
$WarningCount = 0
$Results = @()

# Display banner
$Separator = "=" * 80
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "OneDrive Provisioning Script" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

# Detect environment
$isCloudShell = Test-CloudShell
if ($isCloudShell) {
    Write-StatusMessage "Running in Azure Cloud Shell" -Type Info
}
else {
    Write-StatusMessage "Running in local PowerShell environment" -Type Info
}

# Set output directory
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Get-DefaultOutputDirectory
    Write-StatusMessage "Using default output directory: $OutputDirectory" -Type Info
}

# Create output directory if it doesn't exist
if (!(Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
        Write-StatusMessage "Created output directory: $OutputDirectory" -Type Success
    }
    catch {
        Write-StatusMessage "Failed to create output directory: $($_.Exception.Message)" -Type Error
        exit 1
    }
}

# Verify write permissions
$testFile = Join-Path $OutputDirectory "test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
try {
    "test" | Out-File -FilePath $testFile -ErrorAction Stop
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Write-StatusMessage "Write permissions verified" -Type Success
}
catch {
    Write-StatusMessage "No write permission to output directory" -Type Error
    exit 1
}

# Verify required modules
Write-Host "`nVerifying required modules..." -ForegroundColor Yellow
if (-not (Test-RequiredModules)) {
    Write-StatusMessage "Missing required modules. Please install them and try again." -Type Error
    exit 1
}
Write-StatusMessage "All required modules are available" -Type Success

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
if (-not (Connect-ToMicrosoftGraph)) {
    Write-StatusMessage "Failed to connect to Microsoft Graph. Exiting." -Type Error
    exit 1
}

# Process users
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "Processing $($EmailAddresses.Count) user(s)" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

$counter = 0
foreach ($email in $EmailAddresses) {
    $counter++
    Write-Host "`n[$counter/$($EmailAddresses.Count)] Processing: $email" -ForegroundColor Yellow
    
    $result = Initialize-OneDriveForUser -EmailAddress $email -RetryAttempts $RetryCount
    $Results += $result
    
    # Delay between operations to avoid throttling
    if ($counter -lt $EmailAddresses.Count -and $DelaySeconds -gt 0) {
        Write-Verbose "Waiting $DelaySeconds seconds before next operation..."
        Start-Sleep -Seconds $DelaySeconds
    }
}

# Export results
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "Exporting Results" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputFile = Join-Path $OutputDirectory "OneDrive_Provisioning_Results_$Timestamp.csv"

try {
    $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-StatusMessage "Results exported to: $OutputFile" -Type Success
    
    # Display download command for Cloud Shell
    if ($isCloudShell) {
        Write-Host "`nTo download results file from Cloud Shell:" -ForegroundColor Yellow
        Write-Host "download `"$OutputFile`"" -ForegroundColor Green
    }
}
catch {
    Write-StatusMessage "Failed to export results: $($_.Exception.Message)" -Type Error
}

# Display summary
$SuccessCount = ($Results | Where-Object { $_.OneDriveProvisioned -eq $true }).Count
$FailureCount = ($Results | Where-Object { $_.OneDriveProvisioned -eq $false }).Count

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Total Users Processed: $($EmailAddresses.Count)" -ForegroundColor White
Write-Host "Successful: $SuccessCount" -ForegroundColor Green
Write-Host "Failed: $FailureCount" -ForegroundColor $(if($FailureCount -gt 0){'Red'}else{'Green'})
Write-Host "Warnings: $WarningCount" -ForegroundColor Yellow
Write-Host "Errors: $ErrorCount" -ForegroundColor $(if($ErrorCount -gt 0){'Red'}else{'Green'})

$EndTime = Get-Date
$Duration = $EndTime - $StartTime
Write-Host "`nExecution Time: $($Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

# Cleanup
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-StatusMessage "Disconnected from Microsoft Graph" -Type Success
}
catch {
    Write-Verbose "Disconnect error (non-critical): $($_.Exception.Message)"
}

Write-Host "`nScript completed successfully!" -ForegroundColor Green

#endregion
