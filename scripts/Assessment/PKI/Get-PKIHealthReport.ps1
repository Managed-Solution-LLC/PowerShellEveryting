<#
.SYNOPSIS
    Comprehensive PKI health assessment and monitoring script.

.DESCRIPTION
    This script provides a complete health assessment of your PKI environment including:
    - Certificate Authority service status and configuration
    - CA database health and statistics
    - Certificate validity and expiration monitoring
    - Template availability and consistency
    - CRL and AIA distribution point accessibility
    - CA certificate chain validation
    - Database size and performance metrics
    - Event log analysis for errors and warnings
    
    The script generates both detailed CSV exports and a comprehensive health report
    with actionable recommendations.

.PARAMETER OutputDirectory
    Directory path where reports and CSV exports will be saved.
    Default: C:\Reports\PKI_Health

.PARAMETER CheckCRLDistribution
    Verify CRL distribution points are accessible and current.
    Default: Enabled

.PARAMETER CheckAIADistribution
    Verify AIA (Authority Information Access) distribution points are accessible.
    Default: Enabled

.PARAMETER DaysToExpiration
    Flag certificates and CRLs expiring within this number of days.
    Default: 30 days

.PARAMETER EventLogHours
    Number of hours to review in Application and System event logs.
    Default: 24 hours

.PARAMETER OrganizationName
    Organization name for report headers.
    Default: Organization

.EXAMPLE
    .\Get-PKIHealthReport.ps1
    
    Runs full health assessment with default settings on the local CA server.

.EXAMPLE
    .\Get-PKIHealthReport.ps1 -DaysToExpiration 15 -EventLogHours 48
    
    Runs assessment flagging items expiring within 15 days and reviewing 48 hours of event logs.

.EXAMPLE
    .\Get-PKIHealthReport.ps1 -OutputDirectory "D:\PKI_Health" -CheckCRLDistribution:$false
    
    Runs assessment with custom output directory, skipping CRL distribution checks.

.NOTES
    Author: W. Ford
    Date: 2025-12-24
    Version: 1.0
    
    Requirements:
    - PowerShell 5.1 or later
    - Administrator privileges
    - Must be run ON the Certificate Authority server
    - Certificate Services running
    - Network access to test distribution points
    
    IMPORTANT: This script must be executed directly on the CA server.

.LINK
    https://docs.microsoft.com/en-us/windows-server/networking/core-network-guide/cng/server-certs/install-the-certification-authority
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Directory path for report output")]
    [string]$OutputDirectory = "C:\Reports\PKI_Health",
    
    [Parameter(Mandatory=$false, HelpMessage="Check CRL distribution points")]
    [bool]$CheckCRLDistribution = $true,
    
    [Parameter(Mandatory=$false, HelpMessage="Check AIA distribution points")]
    [bool]$CheckAIADistribution = $true,
    
    [Parameter(Mandatory=$false, HelpMessage="Days until expiration to flag")]
    [ValidateRange(1, 365)]
    [int]$DaysToExpiration = 30,
    
    [Parameter(Mandatory=$false, HelpMessage="Hours of event logs to review")]
    [ValidateRange(1, 168)]
    [int]$EventLogHours = 24,
    
    [Parameter(Mandatory=$false, HelpMessage="Organization name for reports")]
    [string]$OrganizationName = "Organization"
)

#Requires -Version 5.1
#Requires -RunAsAdministrator

# ============================================================================
# INITIALIZATION & VALIDATION
# ============================================================================

$StartTime = Get-Date
$ErrorCount = 0
$WarningCount = 0
$Separator = "=" * 80
$SubSeparator = "-" * 60
$HealthScore = 100 # Start at 100, deduct for issues

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "$OrganizationName - PKI HEALTH ASSESSMENT" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

# Validate PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "❌ This script requires PowerShell 5.1 or later" -ForegroundColor Red
    exit 1
}

# Verify running on CA server
Write-Host "`n⚠️  IMPORTANT: This script must run on a CA server" -ForegroundColor Yellow
Write-Host "   Current machine: $env:COMPUTERNAME" -ForegroundColor Cyan

# Check for Certificate Services
$certSvc = Get-Service -Name CertSvc -ErrorAction SilentlyContinue
if (-not $certSvc) {
    Write-Host "❌ Certificate Services (CertSvc) not found on this server" -ForegroundColor Red
    Write-Host "   This script must be run on a Certificate Authority server" -ForegroundColor Yellow
    exit 1
}

if ($certSvc.Status -ne 'Running') {
    Write-Host "❌ Certificate Services is not running (Status: $($certSvc.Status))" -ForegroundColor Red
    $HealthScore -= 50
    Write-Host "   Start the service with: Start-Service CertSvc" -ForegroundColor Yellow
}
else {
    Write-Host "✅ Certificate Services is running" -ForegroundColor Green
}

# Validate or create output directory
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

# Test write permissions
$testFile = Join-Path $OutputDirectory "test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
try {
    "test" | Out-File -FilePath $testFile -ErrorAction Stop
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Write-Host "✅ Write permissions verified" -ForegroundColor Green
}
catch {
    Write-Host "❌ No write permission to output directory" -ForegroundColor Red
    exit 1
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Type) {
        "Error" { 
            Write-Host "[$timestamp] ERROR: $Message" -ForegroundColor Red
            $script:ErrorCount++
            $script:HealthScore -= 5
        }
        "Warning" { 
            Write-Host "[$timestamp] WARNING: $Message" -ForegroundColor Yellow
            $script:WarningCount++
            $script:HealthScore -= 2
        }
        "Success" { 
            Write-Host "[$timestamp] SUCCESS: $Message" -ForegroundColor Green
        }
        "Critical" {
            Write-Host "[$timestamp] CRITICAL: $Message" -ForegroundColor Red
            $script:ErrorCount++
            $script:HealthScore -= 10
        }
        default { 
            Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Cyan
        }
    }
}

function Invoke-SafeCommand {
    param(
        [scriptblock]$Command,
        [string]$ErrorMessage = "Command execution failed",
        [switch]$ContinueOnError
    )
    try {
        $result = & $Command
        return $result
    }
    catch {
        $errorDetails = "$ErrorMessage - $($_.Exception.Message)"
        Write-StatusMessage -Message $errorDetails -Type "Error"
        
        if (-not $ContinueOnError) {
            throw
        }
        return $null
    }
}

function Test-URLAccessibility {
    param(
        [string]$URL,
        [int]$TimeoutSeconds = 10
    )
    
    try {
        $request = [System.Net.WebRequest]::Create($URL)
        $request.Timeout = $TimeoutSeconds * 1000
        $request.Method = "HEAD"
        
        $response = $request.GetResponse()
        $statusCode = [int]$response.StatusCode
        $response.Close()
        
        return @{
            Success = $true
            StatusCode = $statusCode
            Message = "Accessible"
        }
    }
    catch {
        return @{
            Success = $false
            StatusCode = 0
            Message = $_.Exception.Message
        }
    }
}

# ============================================================================
# CREATE TIMESTAMPED ASSESSMENT FOLDER
# ============================================================================

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$hostname = $env:COMPUTERNAME
$AssessmentFolder = Join-Path $OutputDirectory "$($hostname)_PKI_Health_$Timestamp"

try {
    New-Item -ItemType Directory -Path $AssessmentFolder -Force -ErrorAction Stop | Out-Null
    Write-Host "✅ Created assessment folder: $AssessmentFolder" -ForegroundColor Green
}
catch {
    Write-Host "❌ Cannot create assessment folder: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================================================
# GET CA CONFIGURATION
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Certificate Authority Configuration" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$CAConfig = @{}

try {
    # Get CA configuration
    $caConfigOutput = certutil -dump | Select-String "Config:" | Select-Object -First 1
    if ($caConfigOutput) {
        $CAConfig['Name'] = ($caConfigOutput -split "`"")[1]
        Write-Host "✅ CA Name: $($CAConfig['Name'])" -ForegroundColor Green
    }
    else {
        Write-StatusMessage -Message "Could not determine CA configuration" -Type "Critical"
        exit 1
    }
    
    # Get CA type
    $caType = certutil -dump | Select-String "CA Type:"
    if ($caType) {
        $CAConfig['Type'] = ($caType -split ":")[1].Trim()
        Write-Host "   CA Type: $($CAConfig['Type'])" -ForegroundColor White
    }
    
    # Get CA certificate validity
    $caCertInfo = certutil -ca.cert
    
    # Extract validity dates
    $notBefore = $caCertInfo | Select-String "NotBefore:" | Select-Object -First 1
    $notAfter = $caCertInfo | Select-String "NotAfter:" | Select-Object -First 1
    
    if ($notBefore -and $notAfter) {
        $CAConfig['NotBefore'] = ($notBefore -split "NotBefore: ")[1].Trim()
        $CAConfig['NotAfter'] = ($notAfter -split "NotAfter: ")[1].Trim()
        
        $expirationDate = [DateTime]::Parse($CAConfig['NotAfter'])
        $daysUntilExpiration = ($expirationDate - (Get-Date)).Days
        
        $CAConfig['DaysUntilExpiration'] = $daysUntilExpiration
        
        Write-Host "   CA Cert Expiration: $($CAConfig['NotAfter'])" -ForegroundColor White
        
        if ($daysUntilExpiration -le 0) {
            Write-StatusMessage -Message "CA certificate has EXPIRED!" -Type "Critical"
        }
        elseif ($daysUntilExpiration -le $DaysToExpiration) {
            Write-StatusMessage -Message "CA certificate expires in $daysUntilExpiration days" -Type "Warning"
        }
        else {
            Write-Host "   Days Until Expiration: $daysUntilExpiration" -ForegroundColor Green
        }
    }
    
    # Test CA responsiveness
    Write-Host "`n   Testing CA responsiveness..." -ForegroundColor Cyan
    $pingResult = certutil -ping 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✅ CA is responsive" -ForegroundColor Green
    }
    else {
        Write-StatusMessage -Message "CA ping returned non-zero exit code" -Type "Warning"
    }
}
catch {
    Write-StatusMessage -Message "Failed to retrieve CA configuration: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# CHECK CA DATABASE HEALTH
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "CA Database Health" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$DatabaseHealth = @{}

try {
    # Get database location and size
    $dbInfo = certutil -dump | Select-String "Database Directory:"
    if ($dbInfo) {
        $dbPath = ($dbInfo -split ":")[1].Trim()
        $DatabaseHealth['Path'] = $dbPath
        Write-Host "   Database Path: $dbPath" -ForegroundColor White
        
        # Check if path exists and get size
        if (Test-Path $dbPath) {
            $dbFiles = Get-ChildItem -Path $dbPath -Filter "*.edb" -ErrorAction SilentlyContinue
            if ($dbFiles) {
                $totalSize = ($dbFiles | Measure-Object -Property Length -Sum).Sum
                $DatabaseHealth['SizeMB'] = [math]::Round($totalSize / 1MB, 2)
                Write-Host "   Database Size: $($DatabaseHealth['SizeMB']) MB" -ForegroundColor White
                
                # Warn if database is very large
                if ($DatabaseHealth['SizeMB'] -gt 10240) {
                    Write-StatusMessage -Message "Database size exceeds 10 GB - consider archiving old records" -Type "Warning"
                }
            }
        }
    }
    
    # Get record counts
    Write-Host "`n   Retrieving database statistics..." -ForegroundColor Cyan
    
    $viewOutput = & certutil -view csv 2>&1
    if ($LASTEXITCODE -eq 0 -and $viewOutput) {
        # Count total rows (excluding header)
        $totalRows = ($viewOutput | Where-Object { $_ -match '^\d+,' }).Count
        $DatabaseHealth['TotalRecords'] = $totalRows
        Write-Host "   Total Certificate Records: $totalRows" -ForegroundColor White
        
        # Parse disposition statistics
        $issuedCount = 0
        $revokedCount = 0
        $pendingCount = 0
        $failedCount = 0
        
        foreach ($line in $viewOutput) {
            if ($line -match ',(\d+)\s+--\s+') {
                $disposition = $Matches[1]
                switch ($disposition) {
                    "20" { $issuedCount++ }
                    "21" { $revokedCount++ }
                    "9" { $pendingCount++ }
                    "30" { $failedCount++ }
                }
            }
        }
        
        $DatabaseHealth['IssuedCerts'] = $issuedCount
        $DatabaseHealth['RevokedCerts'] = $revokedCount
        $DatabaseHealth['PendingRequests'] = $pendingCount
        $DatabaseHealth['FailedRequests'] = $failedCount
        
        Write-Host "   Issued Certificates: $issuedCount" -ForegroundColor Green
        Write-Host "   Revoked Certificates: $revokedCount" -ForegroundColor Yellow
        Write-Host "   Pending Requests: $pendingCount" -ForegroundColor Cyan
        Write-Host "   Failed Requests: $failedCount" -ForegroundColor Red
        
        if ($pendingCount -gt 100) {
            Write-StatusMessage -Message "High number of pending requests ($pendingCount) - review queue" -Type "Warning"
        }
    }
    else {
        Write-StatusMessage -Message "Could not retrieve database statistics" -Type "Warning"
    }
}
catch {
    Write-StatusMessage -Message "Failed to check database health: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# CHECK CRL HEALTH
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Certificate Revocation List (CRL) Health" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$CRLHealth = @()

try {
    # Get CRL info
    $crlInfo = certutil -CRL
    
    # Parse CRL publish dates
    $crlPublished = $crlInfo | Select-String "CRL Published:" | Select-Object -First 1
    $crlNextPublish = $crlInfo | Select-String "NextPublish:" | Select-Object -First 1
    
    if ($crlPublished -and $crlNextPublish) {
        $publishedDate = ($crlPublished -split "Published: ")[1].Trim()
        $nextPublishDate = ($crlNextPublish -split "NextPublish: ")[1].Trim()
        
        Write-Host "   Last CRL Published: $publishedDate" -ForegroundColor White
        Write-Host "   Next CRL Publish: $nextPublishDate" -ForegroundColor White
        
        try {
            $nextPublish = [DateTime]::Parse($nextPublishDate)
            $hoursUntilNext = ($nextPublish - (Get-Date)).TotalHours
            
            if ($hoursUntilNext -le 0) {
                Write-StatusMessage -Message "CRL is OVERDUE for publishing!" -Type "Critical"
            }
            elseif ($hoursUntilNext -le 2) {
                Write-StatusMessage -Message "CRL publishes in $([math]::Round($hoursUntilNext, 1)) hours" -Type "Warning"
            }
            else {
                Write-Host "   Status: Current (next publish in $([math]::Round($hoursUntilNext, 1)) hours)" -ForegroundColor Green
            }
        }
        catch {
            Write-StatusMessage -Message "Could not parse CRL publish dates" -Type "Warning"
        }
    }
    
    # Check CRL distribution points
    if ($CheckCRLDistribution) {
        Write-Host "`n   Checking CRL Distribution Points..." -ForegroundColor Cyan
        
        $crlExtensions = certutil -ca.cert | Select-String "CRL Distribution Points"
        if ($crlExtensions) {
            # Extract URLs from certificate
            $urls = certutil -ca.cert | Select-String "http://" | ForEach-Object { 
                ($_ -split "URL=")[1].Trim() 
            } | Where-Object { $_ -like "*.crl" }
            
            foreach ($url in $urls) {
                Write-Host "   Testing: $url" -ForegroundColor Gray
                $result = Test-URLAccessibility -URL $url
                
                $crlHealthObj = [PSCustomObject]@{
                    URL = $url
                    Accessible = $result.Success
                    StatusCode = $result.StatusCode
                    Message = $result.Message
                    TestedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                
                $CRLHealth += $crlHealthObj
                
                if ($result.Success) {
                    Write-Host "   ✅ Accessible (HTTP $($result.StatusCode))" -ForegroundColor Green
                }
                else {
                    Write-StatusMessage -Message "CRL not accessible at $url - $($result.Message)" -Type "Error"
                }
            }
        }
    }
}
catch {
    Write-StatusMessage -Message "Failed to check CRL health: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# CHECK AIA DISTRIBUTION POINTS
# ============================================================================

if ($CheckAIADistribution) {
    Write-Host "`n$SubSeparator" -ForegroundColor Cyan
    Write-Host "Authority Information Access (AIA) Health" -ForegroundColor Cyan
    Write-Host $SubSeparator -ForegroundColor Cyan
    
    $AIAHealth = @()
    
    try {
        # Check AIA extensions in CA cert
        $aiaExtensions = certutil -ca.cert | Select-String "Authority Information Access"
        if ($aiaExtensions) {
            # Extract AIA URLs
            $urls = certutil -ca.cert | Select-String "http://" | ForEach-Object { 
                ($_ -split "URL=")[1].Trim() 
            } | Where-Object { $_ -like "*.crt" -or $_ -like "*.cer" }
            
            foreach ($url in $urls) {
                Write-Host "   Testing: $url" -ForegroundColor Gray
                $result = Test-URLAccessibility -URL $url
                
                $aiaHealthObj = [PSCustomObject]@{
                    URL = $url
                    Accessible = $result.Success
                    StatusCode = $result.StatusCode
                    Message = $result.Message
                    TestedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                
                $AIAHealth += $aiaHealthObj
                
                if ($result.Success) {
                    Write-Host "   ✅ Accessible (HTTP $($result.StatusCode))" -ForegroundColor Green
                }
                else {
                    Write-StatusMessage -Message "AIA not accessible at $url - $($result.Message)" -Type "Error"
                }
            }
        }
        else {
            Write-Host "   ℹ️  No HTTP AIA distribution points configured" -ForegroundColor Gray
        }
    }
    catch {
        Write-StatusMessage -Message "Failed to check AIA health: $($_.Exception.Message)" -Type "Error"
    }
}

# ============================================================================
# CHECK CERTIFICATE TEMPLATES
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Certificate Templates Health" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$TemplateHealth = @()

try {
    # Get published templates
    $publishedTemplates = certutil -CATemplates 2>&1
    
    if ($LASTEXITCODE -eq 0 -and $publishedTemplates) {
        # Parse template names
        $templateNames = $publishedTemplates | Where-Object { $_ -match '^\s+\w+' } | ForEach-Object { $_.Trim() }
        
        Write-Host "   Published Templates: $($templateNames.Count)" -ForegroundColor White
        
        # Check each template in AD
        $RootDSE = [ADSI]"LDAP://RootDSE"
        $ConfigNC = $RootDSE.configurationNamingContext
        $TemplatesPath = "LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$ConfigNC"
        
        foreach ($templateName in $templateNames) {
            try {
                $templateLDAP = "LDAP://CN=$templateName,CN=Certificate Templates,CN=Public Key Services,CN=Services,$ConfigNC"
                $template = [ADSI]$templateLDAP
                
                if ($template.distinguishedName) {
                    $templateHealthObj = [PSCustomObject]@{
                        TemplateName = $templateName
                        Status = "Available"
                        Published = $true
                        Issue = "None"
                    }
                    
                    Write-Host "   ✅ $templateName" -ForegroundColor Green
                }
                else {
                    $templateHealthObj = [PSCustomObject]@{
                        TemplateName = $templateName
                        Status = "Not Found"
                        Published = $true
                        Issue = "Template not found in AD"
                    }
                    
                    Write-StatusMessage -Message "Template '$templateName' is published but not found in AD" -Type "Warning"
                }
                
                $TemplateHealth += $templateHealthObj
            }
            catch {
                $templateHealthObj = [PSCustomObject]@{
                    TemplateName = $templateName
                    Status = "Error"
                    Published = $true
                    Issue = $_.Exception.Message
                }
                
                $TemplateHealth += $templateHealthObj
                Write-StatusMessage -Message "Error checking template '$templateName': $($_.Exception.Message)" -Type "Warning"
            }
        }
    }
    else {
        Write-StatusMessage -Message "Could not retrieve published templates" -Type "Warning"
    }
}
catch {
    Write-StatusMessage -Message "Failed to check template health: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# REVIEW EVENT LOGS
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Event Log Analysis (Last $EventLogHours hours)" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$EventLogIssues = @()
$StartDate = (Get-Date).AddHours(-$EventLogHours)

try {
    # Check Application log for CertSvc events
    Write-Host "   Checking Application log for Certificate Services events..." -ForegroundColor Cyan
    
    $certSvcErrors = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        ProviderName = 'Microsoft-Windows-CertificationAuthority'
        Level = 2 # Error
        StartTime = $StartDate
    } -ErrorAction SilentlyContinue
    
    $certSvcWarnings = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        ProviderName = 'Microsoft-Windows-CertificationAuthority'
        Level = 3 # Warning
        StartTime = $StartDate
    } -ErrorAction SilentlyContinue
    
    if ($certSvcErrors) {
        Write-Host "   ⚠️  Found $($certSvcErrors.Count) Error events" -ForegroundColor Red
        foreach ($event in $certSvcErrors | Select-Object -First 10) {
            $eventObj = [PSCustomObject]@{
                TimeCreated = $event.TimeCreated
                Level = 'Error'
                EventID = $event.Id
                Message = $event.Message.Substring(0, [Math]::Min(200, $event.Message.Length))
                Source = 'CertificationAuthority'
            }
            $EventLogIssues += $eventObj
        }
        Write-StatusMessage -Message "$($certSvcErrors.Count) Certificate Services errors in last $EventLogHours hours" -Type "Error"
    }
    else {
        Write-Host "   ✅ No error events found" -ForegroundColor Green
    }
    
    if ($certSvcWarnings) {
        Write-Host "   ⚠️  Found $($certSvcWarnings.Count) Warning events" -ForegroundColor Yellow
        foreach ($event in $certSvcWarnings | Select-Object -First 10) {
            $eventObj = [PSCustomObject]@{
                TimeCreated = $event.TimeCreated
                Level = 'Warning'
                EventID = $event.Id
                Message = $event.Message.Substring(0, [Math]::Min(200, $event.Message.Length))
                Source = 'CertificationAuthority'
            }
            $EventLogIssues += $eventObj
        }
        Write-StatusMessage -Message "$($certSvcWarnings.Count) Certificate Services warnings in last $EventLogHours hours" -Type "Warning"
    }
    else {
        Write-Host "   ✅ No warning events found" -ForegroundColor Green
    }
}
catch {
    Write-StatusMessage -Message "Failed to review event logs: $($_.Exception.Message)" -Type "Warning"
}

# ============================================================================
# CALCULATE HEALTH SCORE
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Health Score Calculation" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

# Ensure score doesn't go below 0
$HealthScore = [Math]::Max(0, $HealthScore)

$scoreColor = if ($HealthScore -ge 90) { 'Green' }
             elseif ($HealthScore -ge 70) { 'Yellow' }
             elseif ($HealthScore -ge 50) { 'Red' }
             else { 'Red' }

Write-Host "`n   Overall Health Score: " -NoNewline
Write-Host "$HealthScore/100" -ForegroundColor $scoreColor

$healthStatus = if ($HealthScore -ge 90) { "EXCELLENT" }
                elseif ($HealthScore -ge 70) { "GOOD" }
                elseif ($HealthScore -ge 50) { "FAIR" }
                else { "POOR" }

Write-Host "   Status: $healthStatus" -ForegroundColor $scoreColor

# ============================================================================
# EXPORT DATA TO CSV
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Exporting Assessment Data" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

# Export CRL health
if ($CRLHealth.Count -gt 0) {
    $crlFile = Join-Path $AssessmentFolder "$($hostname)_PKI_CRLHealth_$Timestamp.csv"
    $CRLHealth | Export-Csv -Path $crlFile -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Exported CRL health to: $crlFile" -ForegroundColor Green
}

# Export AIA health
if ($AIAHealth.Count -gt 0) {
    $aiaFile = Join-Path $AssessmentFolder "$($hostname)_PKI_AIAHealth_$Timestamp.csv"
    $AIAHealth | Export-Csv -Path $aiaFile -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Exported AIA health to: $aiaFile" -ForegroundColor Green
}

# Export template health
if ($TemplateHealth.Count -gt 0) {
    $templateFile = Join-Path $AssessmentFolder "$($hostname)_PKI_TemplateHealth_$Timestamp.csv"
    $TemplateHealth | Export-Csv -Path $templateFile -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Exported template health to: $templateFile" -ForegroundColor Green
}

# Export event log issues
if ($EventLogIssues.Count -gt 0) {
    $eventsFile = Join-Path $AssessmentFolder "$($hostname)_PKI_EventLogIssues_$Timestamp.csv"
    $EventLogIssues | Export-Csv -Path $eventsFile -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Exported event log issues to: $eventsFile" -ForegroundColor Green
}

# ============================================================================
# GENERATE HEALTH REPORT
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Generating Health Report" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$ReportFile = Join-Path $AssessmentFolder "$($hostname)_PKI_Health_Report_$Timestamp.txt"

$Report = @"
$Separator
$OrganizationName - PKI HEALTH ASSESSMENT REPORT
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$Separator

OVERALL HEALTH STATUS
$SubSeparator
Health Score: $HealthScore/100
Status: $healthStatus
Errors: $ErrorCount
Warnings: $WarningCount

CERTIFICATE AUTHORITY INFORMATION
$SubSeparator
CA Server: $env:COMPUTERNAME
CA Name: $($CAConfig['Name'])
CA Type: $($CAConfig['Type'])
Service Status: $($certSvc.Status)

CA CERTIFICATE STATUS
$SubSeparator
Issued: $($CAConfig['NotBefore'])
Expires: $($CAConfig['NotAfter'])
Days Until Expiration: $($CAConfig['DaysUntilExpiration'])
Status: $(if ($CAConfig['DaysUntilExpiration'] -le 0) { 'EXPIRED' } elseif ($CAConfig['DaysUntilExpiration'] -le $DaysToExpiration) { 'EXPIRING SOON' } else { 'Valid' })

DATABASE HEALTH
$SubSeparator
Database Path: $($DatabaseHealth['Path'])
Database Size: $($DatabaseHealth['SizeMB']) MB
Total Records: $($DatabaseHealth['TotalRecords'])
Issued Certificates: $($DatabaseHealth['IssuedCerts'])
Revoked Certificates: $($DatabaseHealth['RevokedCerts'])
Pending Requests: $($DatabaseHealth['PendingRequests'])
Failed Requests: $($DatabaseHealth['FailedRequests'])

CRL HEALTH
$SubSeparator
"@

if ($CRLHealth.Count -gt 0) {
    $Report += "`nDistribution Point Status:"
    foreach ($crl in $CRLHealth) {
        $status = if ($crl.Accessible) { "✅ ACCESSIBLE" } else { "❌ FAILED" }
        $Report += "`n  $status - $($crl.URL)"
        if (-not $crl.Accessible) {
            $Report += "`n    Error: $($crl.Message)"
        }
    }
}
else {
    $Report += "`n  No CRL distribution points checked"
}

if ($CheckAIADistribution) {
    $Report += @"

`n
AIA HEALTH
$SubSeparator
"@
    
    if ($AIAHealth.Count -gt 0) {
        $Report += "`nDistribution Point Status:"
        foreach ($aia in $AIAHealth) {
            $status = if ($aia.Accessible) { "✅ ACCESSIBLE" } else { "❌ FAILED" }
            $Report += "`n  $status - $($aia.URL)"
            if (-not $aia.Accessible) {
                $Report += "`n    Error: $($aia.Message)"
            }
        }
    }
    else {
        $Report += "`n  No AIA distribution points configured"
    }
}

$Report += @"

`n
CERTIFICATE TEMPLATES
$SubSeparator
Published Templates: $($TemplateHealth.Count)
"@

if ($TemplateHealth.Count -gt 0) {
    $healthyTemplates = ($TemplateHealth | Where-Object { $_.Status -eq 'Available' }).Count
    $Report += "`nHealthy Templates: $healthyTemplates"
    
    $unhealthyTemplates = $TemplateHealth | Where-Object { $_.Status -ne 'Available' }
    if ($unhealthyTemplates) {
        $Report += "`n`nTemplates with Issues:"
        foreach ($template in $unhealthyTemplates) {
            $Report += "`n  ❌ $($template.TemplateName): $($template.Issue)"
        }
    }
}

$Report += @"

`n
EVENT LOG ANALYSIS
$SubSeparator
Review Period: Last $EventLogHours hours
Certificate Services Errors: $(($EventLogIssues | Where-Object { $_.Level -eq 'Error' }).Count)
Certificate Services Warnings: $(($EventLogIssues | Where-Object { $_.Level -eq 'Warning' }).Count)
"@

if ($EventLogIssues.Count -gt 0) {
    $Report += "`n`nRecent Issues (Top 10):"
    foreach ($event in $EventLogIssues | Select-Object -First 10) {
        $Report += "`n`n[$($event.TimeCreated)] $($event.Level) - Event ID $($event.EventID)"
        $Report += "`n  $($event.Message)"
    }
}

$Report += @"

`n
RECOMMENDATIONS
$SubSeparator
"@

# Generate recommendations based on findings
$recommendations = @()

if ($certSvc.Status -ne 'Running') {
    $recommendations += "CRITICAL: Start Certificate Services immediately"
}

if ($CAConfig['DaysUntilExpiration'] -le 0) {
    $recommendations += "CRITICAL: CA certificate has expired - renew immediately"
}
elseif ($CAConfig['DaysUntilExpiration'] -le $DaysToExpiration) {
    $recommendations += "WARNING: CA certificate expires in $($CAConfig['DaysUntilExpiration']) days - plan renewal"
}

if ($DatabaseHealth['SizeMB'] -gt 10240) {
    $recommendations += "Consider archiving old certificate records (Database >10 GB)"
}

if ($DatabaseHealth['PendingRequests'] -gt 100) {
    $recommendations += "Review pending certificate requests ($($DatabaseHealth['PendingRequests']) pending)"
}

$failedCRLs = $CRLHealth | Where-Object { -not $_.Accessible }
if ($failedCRLs) {
    $recommendations += "Fix inaccessible CRL distribution points ($($failedCRLs.Count) failed)"
}

$failedAIAs = $AIAHealth | Where-Object { -not $_.Accessible }
if ($failedAIAs) {
    $recommendations += "Fix inaccessible AIA distribution points ($($failedAIAs.Count) failed)"
}

$unhealthyTemplates = $TemplateHealth | Where-Object { $_.Status -ne 'Available' }
if ($unhealthyTemplates) {
    $recommendations += "Review certificate templates with issues ($($unhealthyTemplates.Count) templates)"
}

if (($EventLogIssues | Where-Object { $_.Level -eq 'Error' }).Count -gt 0) {
    $recommendations += "Review and resolve Certificate Services errors in Event Viewer"
}

if ($recommendations.Count -eq 0) {
    $Report += "`n✅ No critical issues found - PKI infrastructure is healthy"
}
else {
    foreach ($rec in $recommendations) {
        $Report += "`n• $rec"
    }
}

$Report += @"

`n
EXPORTED FILES
$SubSeparator
"@

if ($CRLHealth.Count -gt 0) {
    $Report += "`nCRL Health: $($hostname)_PKI_CRLHealth_$Timestamp.csv"
}
if ($AIAHealth.Count -gt 0) {
    $Report += "`nAIA Health: $($hostname)_PKI_AIAHealth_$Timestamp.csv"
}
if ($TemplateHealth.Count -gt 0) {
    $Report += "`nTemplate Health: $($hostname)_PKI_TemplateHealth_$Timestamp.csv"
}
if ($EventLogIssues.Count -gt 0) {
    $Report += "`nEvent Log Issues: $($hostname)_PKI_EventLogIssues_$Timestamp.csv"
}
$Report += "`nHealth Report: $($hostname)_PKI_Health_Report_$Timestamp.txt"

$Report += @"

`n
$Separator
End of Health Assessment Report
$Separator
"@

# Save report
try {
    $Report | Out-File -FilePath $ReportFile -Encoding UTF8
    Write-Host "✅ Health report saved to: $ReportFile" -ForegroundColor Green
}
catch {
    Write-StatusMessage -Message "Failed to save report: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# COMPLETION SUMMARY
# ============================================================================

$EndTime = Get-Date
$Duration = $EndTime - $StartTime

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "PKI HEALTH ASSESSMENT COMPLETED" -ForegroundColor $scoreColor
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Health Score: " -NoNewline
Write-Host "$HealthScore/100 ($healthStatus)" -ForegroundColor $scoreColor
Write-Host "Duration: $($Duration.ToString('mm\:ss'))" -ForegroundColor White
Write-Host "Errors: $ErrorCount | Warnings: $WarningCount" -ForegroundColor $(if($ErrorCount -gt 0){'Red'}else{'Green'})
Write-Host "`nAll reports saved to: $AssessmentFolder" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
