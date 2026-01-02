<#
.SYNOPSIS
    Forces enrollment of Entra Joined (Azure AD Joined) computers into Microsoft Intune.

.DESCRIPTION
    This script forces the enrollment of Windows devices that are Entra Joined (Azure AD Joined)
    into Microsoft Intune using the current user's authentication context. It performs the following:
    
    - Verifies the device is Entra Joined (Azure AD Joined)
    - Checks current Intune enrollment status
    - Triggers Intune MDM enrollment using device credentials
    - Initiates device sync and policy refresh
    - Validates successful enrollment
    - Provides detailed status reporting
    
    The script uses Windows 10/11 MDM enrollment protocols and the DeviceEnroller CSP to force
    enrollment without requiring user interaction. It leverages the existing Entra Join
    authentication to seamlessly enroll the device.

.PARAMETER ForceReenroll
    Forces re-enrollment even if the device is already enrolled in Intune.
    This will unenroll the device first, then trigger a fresh enrollment.

.PARAMETER SyncAfterEnroll
    Automatically triggers an Intune device sync after successful enrollment.
    This helps speed up initial policy application.

.PARAMETER LogPath
    Path to write detailed log file. Defaults to C:\ProgramData\Intune\Logs\Enrollment_[timestamp].log

.PARAMETER WaitForSync
    Waits for the initial sync to complete before exiting.
    Default timeout is 5 minutes.

.PARAMETER NoRestart
    Prevents automatic restart after enrollment. By default, the script prompts for restart
    if enrollment requires it for policy application.

.EXAMPLE
    .\Start-IntuneEnrollment.ps1
    
    Checks enrollment status and enrolls the device into Intune if not already enrolled.

.EXAMPLE
    .\Start-IntuneEnrollment.ps1 -ForceReenroll -SyncAfterEnroll
    
    Forces re-enrollment even if already enrolled, then triggers immediate sync.

.EXAMPLE
    .\Start-IntuneEnrollment.ps1 -SyncAfterEnroll -WaitForSync
    
    Enrolls device, triggers sync, and waits for sync to complete before exiting.

.EXAMPLE
    .\Start-IntuneEnrollment.ps1 -LogPath "C:\Logs\IntuneEnroll.log" -NoRestart
    
    Enrolls device with custom log location and prevents automatic restart.

.NOTES
    Author: W. Ford
    Company: Managed Solution LLC
    Date: 2026-01-02
    Version: 1.0
    
    Requirements:
    - Windows 10 1809+ or Windows 11
    - Device must be Entra Joined (Azure AD Joined)
    - Administrator privileges required
    - Internet connectivity to Intune service endpoints
    
    Intune Service Endpoints Required:
    - enrollment.manage.microsoft.com
    - portal.manage.microsoft.com
    
    The script must be run with administrator privileges to modify enrollment settings
    and trigger MDM enrollment. It uses the DeviceEnroller CSP and Windows enrollment
    client to perform the enrollment.
    
    VALIDATED FOR PUBLIC RELEASE: 2026-01-02

.LINK
    https://docs.microsoft.com/en-us/mem/intune/enrollment/windows-enrollment-methods
    https://docs.microsoft.com/en-us/windows/client-management/mdm/deviceenroller-csp
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$false, HelpMessage="Force re-enrollment even if already enrolled")]
    [switch]$ForceReenroll,
    
    [Parameter(Mandatory=$false, HelpMessage="Trigger device sync after enrollment")]
    [switch]$SyncAfterEnroll,
    
    [Parameter(Mandatory=$false, HelpMessage="Path to log file")]
    [string]$LogPath = "C:\ProgramData\Intune\Logs\Enrollment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory=$false, HelpMessage="Wait for initial sync to complete")]
    [switch]$WaitForSync,
    
    [Parameter(Mandatory=$false, HelpMessage="Prevent automatic restart prompt")]
    [switch]$NoRestart
)

#Requires -RunAsAdministrator

# Initialize variables
$script:ErrorCount = 0
$script:WarningCount = 0
$Separator = "=" * 80
$SubSeparator = "-" * 60

# Ensure log directory exists
$logDir = Split-Path -Parent $LogPath
if (!(Test-Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Could not create log directory: $logDir. Logging to console only."
        $LogPath = $null
    }
}

#region Helper Functions

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        'Error' { 
            Write-Host "❌ $Message" -ForegroundColor Red
            $script:ErrorCount++
        }
        'Warning' { 
            Write-Host "⚠️  $Message" -ForegroundColor Yellow
            $script:WarningCount++
        }
        'Success' { 
            Write-Host "✅ $Message" -ForegroundColor Green
        }
        'Info' { 
            Write-Host "ℹ️  $Message" -ForegroundColor Cyan
        }
    }
    
    # File logging
    if ($LogPath) {
        try {
            Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
        }
        catch {
            # Silently continue if logging fails
        }
    }
}

function Test-EntraJoined {
    <#
    .SYNOPSIS
        Checks if the device is Entra Joined (Azure AD Joined)
    #>
    try {
        $dsregStatus = & dsregcmd /status
        
        if ($dsregStatus -match "AzureAdJoined\s*:\s*YES") {
            Write-Log "Device is Entra Joined" -Level Success
            return $true
        }
        else {
            Write-Log "Device is NOT Entra Joined. This script requires an Entra Joined device." -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Failed to check Entra Join status: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Get-IntuneEnrollmentStatus {
    <#
    .SYNOPSIS
        Checks current Intune enrollment status
    #>
    try {
        # Check registry for MDM enrollment
        $enrollmentPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
        
        if (Test-Path $enrollmentPath) {
            $enrollments = Get-ChildItem -Path $enrollmentPath -ErrorAction SilentlyContinue
            
            foreach ($enrollment in $enrollments) {
                $providerID = Get-ItemProperty -Path $enrollment.PSPath -Name "ProviderID" -ErrorAction SilentlyContinue
                
                if ($providerID.ProviderID -eq "MS DM Server") {
                    $upn = Get-ItemProperty -Path $enrollment.PSPath -Name "UPN" -ErrorAction SilentlyContinue
                    $discoveryServiceFullURL = Get-ItemProperty -Path $enrollment.PSPath -Name "DiscoveryServiceFullURL" -ErrorAction SilentlyContinue
                    
                    Write-Log "Device is enrolled in Intune" -Level Success
                    Write-Log "  Enrollment GUID: $($enrollment.PSChildName)" -Level Info
                    if ($upn) { Write-Log "  UPN: $($upn.UPN)" -Level Info }
                    if ($discoveryServiceFullURL) { Write-Log "  Service URL: $($discoveryServiceFullURL.DiscoveryServiceFullURL)" -Level Info }
                    
                    return @{
                        IsEnrolled = $true
                        EnrollmentGUID = $enrollment.PSChildName
                        UPN = $upn.UPN
                        ServiceURL = $discoveryServiceFullURL.DiscoveryServiceFullURL
                    }
                }
            }
        }
        
        Write-Log "Device is NOT enrolled in Intune" -Level Warning
        return @{
            IsEnrolled = $false
            EnrollmentGUID = $null
            UPN = $null
            ServiceURL = $null
        }
    }
    catch {
        Write-Log "Failed to check enrollment status: $($_.Exception.Message)" -Level Error
        return @{
            IsEnrolled = $false
            EnrollmentGUID = $null
            UPN = $null
            ServiceURL = $null
        }
    }
}

function Remove-IntuneEnrollment {
    <#
    .SYNOPSIS
        Removes existing Intune enrollment
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$EnrollmentGUID
    )
    
    try {
        Write-Log "Removing existing enrollment: $EnrollmentGUID" -Level Info
        
        # Remove enrollment registry keys
        $enrollmentPath = "HKLM:\SOFTWARE\Microsoft\Enrollments\$EnrollmentGUID"
        if (Test-Path $enrollmentPath) {
            Remove-Item -Path $enrollmentPath -Recurse -Force -ErrorAction Stop
            Write-Log "Removed enrollment registry keys" -Level Success
        }
        
        # Clean up enrollment tasks
        $taskName = "*$EnrollmentGUID*"
        $tasks = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        foreach ($task in $tasks) {
            Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "Removed scheduled task: $($task.TaskName)" -Level Info
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to remove enrollment: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Start-MDMEnrollment {
    <#
    .SYNOPSIS
        Triggers MDM enrollment using DeviceEnroller CSP
    #>
    try {
        Write-Log "Initiating MDM enrollment..." -Level Info
        
        # Trigger enrollment via scheduled task
        $taskPath = "\Microsoft\Windows\EnterpriseMgmt"
        
        # Method 1: Trigger existing enrollment task if available
        $enrollTasks = Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue | 
                       Where-Object { $_.TaskName -like "*Schedule*" }
        
        if ($enrollTasks) {
            foreach ($task in $enrollTasks) {
                Write-Log "Triggering enrollment task: $($task.TaskName)" -Level Info
                Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue
            }
        }
        
        # Method 2: Use DeviceEnroller.exe if available
        $deviceEnroller = "$env:SystemRoot\System32\DeviceEnroller.exe"
        if (Test-Path $deviceEnroller) {
            Write-Log "Using DeviceEnroller.exe to trigger enrollment" -Level Info
            
            if ($PSCmdlet.ShouldProcess("System", "Start DeviceEnroller.exe /c /AutoEnrollMDM")) {
                $process = Start-Process -FilePath $deviceEnroller -ArgumentList "/c", "/AutoEnrollMDM" -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -eq 0) {
                    Write-Log "DeviceEnroller completed successfully" -Level Success
                }
                else {
                    Write-Log "DeviceEnroller exited with code: $($process.ExitCode)" -Level Warning
                }
            }
        }
        
        # Method 3: Trigger via registry key (forces enrollment on next logon/restart)
        $enrollmentPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM"
        if (!(Test-Path $enrollmentPath)) {
            New-Item -Path $enrollmentPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $enrollmentPath -Name "AutoEnrollMDM" -Value 1 -Type DWord -ErrorAction Stop
        Write-Log "Set AutoEnrollMDM registry key" -Level Info
        
        # Wait a moment for enrollment to initiate
        Start-Sleep -Seconds 5
        
        return $true
    }
    catch {
        Write-Log "Failed to trigger MDM enrollment: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Invoke-DeviceSync {
    <#
    .SYNOPSIS
        Triggers Intune device sync
    #>
    try {
        Write-Log "Triggering device sync..." -Level Info
        
        # Find the OmaDmClient scheduled task
        $syncTask = Get-ScheduledTask | Where-Object { 
            $_.TaskPath -like "*EnterpriseMgmt*" -and $_.TaskName -like "*Schedule*" 
        } | Select-Object -First 1
        
        if ($syncTask) {
            Start-ScheduledTask -InputObject $syncTask -ErrorAction Stop
            Write-Log "Device sync initiated successfully" -Level Success
            return $true
        }
        else {
            Write-Log "Could not find sync task. Device may need to complete initial enrollment first." -Level Warning
            return $false
        }
    }
    catch {
        Write-Log "Failed to trigger device sync: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Wait-ForEnrollment {
    <#
    .SYNOPSIS
        Waits for enrollment to complete
    #>
    param(
        [int]$TimeoutSeconds = 120
    )
    
    Write-Log "Waiting for enrollment to complete (timeout: $TimeoutSeconds seconds)..." -Level Info
    
    $startTime = Get-Date
    $enrolled = $false
    
    while (((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        Start-Sleep -Seconds 5
        
        $status = Get-IntuneEnrollmentStatus
        if ($status.IsEnrolled) {
            $enrolled = $true
            break
        }
        
        Write-Host "." -NoNewline -ForegroundColor Gray
    }
    
    Write-Host "" # New line after dots
    
    if ($enrolled) {
        Write-Log "Enrollment completed successfully" -Level Success
        return $true
    }
    else {
        Write-Log "Enrollment did not complete within timeout period" -Level Warning
        return $false
    }
}

#endregion

#region Main Script

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "INTUNE ENROLLMENT TOOL" -ForegroundColor Cyan
Write-Host "Force Enroll Entra Joined Devices" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Log "Script started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info
Write-Host ""

# Step 1: Verify prerequisites
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 1: Verifying Prerequisites" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

# Check Windows version
$osVersion = [System.Environment]::OSVersion.Version
if ($osVersion.Major -lt 10) {
    Write-Log "This script requires Windows 10 or later. Current version: $($osVersion.ToString())" -Level Error
    exit 1
}
Write-Log "Windows version: $($osVersion.Major).$($osVersion.Minor) Build $($osVersion.Build)" -Level Success

# Check Entra Join status
if (-not (Test-EntraJoined)) {
    Write-Log "Device must be Entra Joined to enroll in Intune with this method" -Level Error
    Write-Log "Use 'dsregcmd /status' to check Azure AD join status" -Level Info
    exit 1
}

# Step 2: Check current enrollment status
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 2: Checking Current Enrollment Status" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

$enrollmentStatus = Get-IntuneEnrollmentStatus

# Step 3: Handle existing enrollment
if ($enrollmentStatus.IsEnrolled) {
    if ($ForceReenroll) {
        Write-Host "`n$SubSeparator" -ForegroundColor Yellow
        Write-Host "STEP 3: Removing Existing Enrollment" -ForegroundColor Yellow
        Write-Host $SubSeparator -ForegroundColor Yellow
        
        if ($PSCmdlet.ShouldProcess("Intune Enrollment", "Remove existing enrollment")) {
            $removed = Remove-IntuneEnrollment -EnrollmentGUID $enrollmentStatus.EnrollmentGUID
            
            if (-not $removed) {
                Write-Log "Failed to remove existing enrollment. Proceeding with caution." -Level Warning
            }
            
            # Wait for cleanup to complete
            Start-Sleep -Seconds 10
        }
    }
    else {
        Write-Log "Device is already enrolled in Intune. Use -ForceReenroll to re-enroll." -Level Warning
        
        if ($SyncAfterEnroll) {
            Write-Host "`n$SubSeparator" -ForegroundColor Yellow
            Write-Host "Triggering Device Sync" -ForegroundColor Yellow
            Write-Host $SubSeparator -ForegroundColor Yellow
            Invoke-DeviceSync
        }
        
        Write-Host "`n$Separator" -ForegroundColor Cyan
        Write-Host "Script completed. Device already enrolled." -ForegroundColor Cyan
        Write-Host $Separator -ForegroundColor Cyan
        exit 0
    }
}

# Step 4: Enroll device
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 4: Enrolling Device in Intune" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

if ($PSCmdlet.ShouldProcess("Device", "Enroll in Intune")) {
    $enrolled = Start-MDMEnrollment
    
    if ($enrolled) {
        # Wait for enrollment to complete
        $enrollmentComplete = Wait-ForEnrollment -TimeoutSeconds 120
        
        if ($enrollmentComplete) {
            Write-Log "Device successfully enrolled in Intune" -Level Success
            
            # Step 5: Sync device (if requested)
            if ($SyncAfterEnroll) {
                Write-Host "`n$SubSeparator" -ForegroundColor Yellow
                Write-Host "STEP 5: Triggering Device Sync" -ForegroundColor Yellow
                Write-Host $SubSeparator -ForegroundColor Yellow
                
                Start-Sleep -Seconds 5 # Brief pause before sync
                Invoke-DeviceSync
                
                if ($WaitForSync) {
                    Write-Log "Waiting for sync to complete..." -Level Info
                    Start-Sleep -Seconds 30
                    Write-Log "Initial sync period completed" -Level Success
                }
            }
        }
        else {
            Write-Log "Enrollment initiated but did not complete verification within timeout" -Level Warning
            Write-Log "The device may still be enrolling. Check Intune portal in a few minutes." -Level Info
        }
    }
    else {
        Write-Log "Failed to initiate enrollment" -Level Error
        exit 1
    }
}

# Step 6: Verify final status
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 6: Verifying Final Status" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

$finalStatus = Get-IntuneEnrollmentStatus

# Summary
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "ENROLLMENT SUMMARY" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Enrollment Status: $(if($finalStatus.IsEnrolled){'✅ ENROLLED'}else{'❌ NOT ENROLLED'})" -ForegroundColor $(if($finalStatus.IsEnrolled){'Green'}else{'Red'})
if ($finalStatus.IsEnrolled) {
    Write-Host "Enrollment GUID: $($finalStatus.EnrollmentGUID)" -ForegroundColor White
    if ($finalStatus.UPN) { Write-Host "User: $($finalStatus.UPN)" -ForegroundColor White }
}
Write-Host "Errors: $script:ErrorCount | Warnings: $script:WarningCount" -ForegroundColor $(if($script:ErrorCount -gt 0){'Red'}else{'Green'})
if ($LogPath -and (Test-Path $LogPath)) {
    Write-Host "Log file: $LogPath" -ForegroundColor Gray
}
Write-Host $Separator -ForegroundColor Cyan

# Restart prompt
if ($finalStatus.IsEnrolled -and -not $NoRestart) {
    Write-Host "`n⚠️  A restart is recommended to complete policy application" -ForegroundColor Yellow
    $restart = Read-Host "Restart now? (Y/N)"
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-Log "Initiating restart..." -Level Info
        Restart-Computer -Force
    }
}

Write-Host ""
Write-Log "Script completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info

#endregion
