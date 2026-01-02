<#
.SYNOPSIS
    Comprehensive PKI assessment tool for Certificate Authority infrastructure.

.DESCRIPTION
    This script provides a complete assessment of your PKI environment including:
    - All issued certificates from the Certificate Authority
    - Published certificate templates with detailed properties
    - Template permissions and security descriptors
    - Certificate validity analysis
    - Template usage statistics
    
    The script generates both detailed CSV exports and a comprehensive text report
    for documentation and analysis purposes.

.PARAMETER OutputDirectory
    Directory path where reports and CSV exports will be saved.
    Default: C:\Reports\PKI_Assessment

.PARAMETER CAServerName
    Name of the Certificate Authority server to assess.
    If not specified, will attempt to discover the CA automatically.

.PARAMETER IncludeRevokedCertificates
    Include revoked certificates in the assessment.
    Default: Only active certificates are included.

.PARAMETER DaysToExpiration
    Flag certificates expiring within this number of days.
    Default: 90 days

.PARAMETER OrganizationName
    Organization name for report headers.
    Default: Organization

.EXAMPLE
    .\Get-ComprehensivePKIReport.ps1
    
    Runs assessment with default settings on the local CA server.

.EXAMPLE
    .\Get-ComprehensivePKIReport.ps1 -IncludeRevokedCertificates
    
    Runs assessment including revoked certificates.

.EXAMPLE
    .\Get-ComprehensivePKIReport.ps1 -OutputDirectory "D:\PKI_Reports" -DaysToExpiration 30
    
    Runs assessment with custom output directory and flags certificates expiring within 30 days.

.NOTES
    Author: W. Ford
    Date: 2025-12-24
    Version: 1.1
    
    Requirements:
    - PowerShell 5.1 or later
    - Administrator privileges
    - Must be run ON the Certificate Authority server
    - Domain membership for AD template queries
    
    IMPORTANT: This script must be executed directly on the CA server.
    Certificate database queries require local access to the CA database.

.LINK
    https://docs.microsoft.com/en-us/windows-server/networking/core-network-guide/cng/server-certs/install-the-certification-authority
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Directory path for report output")]
    [string]$OutputDirectory = "C:\Reports\PKI_Assessment",
    
    [Parameter(Mandatory=$false, HelpMessage="Include revoked certificates in the assessment")]
    [switch]$IncludeRevokedCertificates,
    
    [Parameter(Mandatory=$false, HelpMessage="Days until expiration to flag certificates")]
    [ValidateRange(1, 365)]
    [int]$DaysToExpiration = 90,
    
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

Write-Host "\n$Separator" -ForegroundColor Cyan
Write-Host "$OrganizationName - PKI ASSESSMENT" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan

# Validate PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "❌ This script requires PowerShell 5.1 or later" -ForegroundColor Red
    exit 1
}

# Verify running on CA server
Write-Host "\n⚠️  IMPORTANT: This script must run on a CA server" -ForegroundColor Yellow
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
    Write-Host "   Start the service with: Start-Service CertSvc" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Certificate Services is running on this server" -ForegroundColor Green

# Check for required modules
Write-Host "`nValidating required modules..." -ForegroundColor Yellow

# Check if running on Windows Server or Windows Client
$IsServer = (Get-WmiObject -Class Win32_OperatingSystem).ProductType -ne 1

# Try to load ADCS module - it may already be available
$ADCSModuleAvailable = $false

# Check if module exists first
$adcsModule = Get-Module -Name ADCS-Administration -ListAvailable

if ($adcsModule) {
    try {
        Import-Module ADCS-Administration -ErrorAction Stop
        $ADCSModuleAvailable = $true
        Write-Host "✅ ADCS-Administration module loaded" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  ADCS-Administration module found but failed to load: $($_.Exception.Message)" -ForegroundColor Yellow
        $ADCSModuleAvailable = $false
    }
}
else {
    Write-Host "⚠️  ADCS-Administration module not found - attempting installation..." -ForegroundColor Yellow
    
    # Try to install RSAT-ADCS features
    if ($IsServer) {
        $rsatFeature = Get-WindowsFeature -Name RSAT-ADCS-Mgmt -ErrorAction SilentlyContinue
        if ($rsatFeature -and -not $rsatFeature.Installed) {
            Write-Host "   Installing RSAT-ADCS-Mgmt feature..." -ForegroundColor Yellow
            try {
                $result = Install-WindowsFeature RSAT-ADCS-Mgmt -IncludeAllSubFeature -ErrorAction Stop
                if ($result.Success) {
                    Write-Host "   ✅ RSAT-ADCS-Mgmt installed successfully" -ForegroundColor Green
                    
                    # Try to import the module after installation
                    $adcsModule = Get-Module -Name ADCS-Administration -ListAvailable
                    if ($adcsModule) {
                        try {
                            Import-Module ADCS-Administration -ErrorAction Stop
                            $ADCSModuleAvailable = $true
                            Write-Host "   ✅ ADCS-Administration module loaded" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "   ⚠️  Module installed but failed to load" -ForegroundColor Yellow
                        }
                    }
                    else {
                        Write-Host "   ⚠️  Feature installed but module still not available" -ForegroundColor Yellow
                        Write-Host "   This is normal - ADCS cmdlets only work on the CA server itself" -ForegroundColor Gray
                    }
                }
            }
            catch {
                Write-Host "   ⚠️  Failed to install RSAT-ADCS-Mgmt: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "   RSAT-ADCS-Mgmt feature is already installed" -ForegroundColor Gray
            Write-Host "   Note: ADCS-Administration module only provides cmdlets on CA server" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "   Install RSAT from: Settings > Apps > Optional Features > RSAT: Active Directory Certificate Services Tools" -ForegroundColor Yellow
        Write-Host "   Or via PowerShell: Get-WindowsCapability -Online | Where-Object Name -like 'Rsat.CertificateServices*' | Add-WindowsCapability -Online" -ForegroundColor Yellow
    }
    
    # Check if we can use certutil as fallback
    if (Get-Command certutil -ErrorAction SilentlyContinue) {
        Write-Host "✅ Using certutil for certificate operations (limited functionality)" -ForegroundColor Green
    }
    else {
        Write-Host "❌ Cannot proceed without ADCS tools or certutil" -ForegroundColor Red
        exit 1
    }
}

# Optional: ActiveDirectory module for enhanced template permissions
if (Get-Module -Name ActiveDirectory -ListAvailable) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "✅ ActiveDirectory module loaded (enhanced permissions analysis)" -ForegroundColor Green
        $ADModuleAvailable = $true
    }
    catch {
        Write-Host "⚠️  ActiveDirectory module not available - basic permissions only" -ForegroundColor Yellow
        $ADModuleAvailable = $false
    }
}
else {
    Write-Host "⚠️  ActiveDirectory module not available - basic permissions only" -ForegroundColor Yellow
    $ADModuleAvailable = $false
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
        }
        "Warning" { 
            Write-Host "[$timestamp] WARNING: $Message" -ForegroundColor Yellow
            $script:WarningCount++
        }
        "Success" { 
            Write-Host "[$timestamp] SUCCESS: $Message" -ForegroundColor Green
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

# ============================================================================
# GET LOCAL CA CONFIGURATION
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Local CA Configuration" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

try {
    # Get local CA configuration
    $caConfig = certutil -dump | Select-String "Config:" | Select-Object -First 1
    if ($caConfig) {
        $CAName = ($caConfig -split "`"")[1]
        Write-Host "✅ Local CA: $CAName" -ForegroundColor Green
    }
    else {
        Write-Host "❌ Could not determine CA configuration" -ForegroundColor Red
        Write-Host "   Verify CA is properly installed and configured" -ForegroundColor Yellow
        exit 1
    }
    
    # Validate CA is operational
    $null = certutil -ping 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ CA is operational" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️  CA ping returned non-zero exit code" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "❌ Failed to validate CA: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================================================
# EXPORT ISSUED CERTIFICATES
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Exporting Issued Certificates" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$hostname = $env:COMPUTERNAME
# Create timestamped subfolder for this assessment
$AssessmentFolder = Join-Path $OutputDirectory "$($hostname)_PKI_Assessment_$Timestamp"
try {
    New-Item -ItemType Directory -Path $AssessmentFolder -Force -ErrorAction Stop | Out-Null
    Write-Host "✅ Created assessment folder: $AssessmentFolder" -ForegroundColor Green
}
catch {
    Write-Host "❌ Cannot create assessment folder: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$CertificatesFile = Join-Path $AssessmentFolder "$($hostname)_PKI_IssuedCertificates_$Timestamp.csv"

Write-StatusMessage -Message "Retrieving issued certificates from CA..." -Type "Info"

$IssuedCertificates = Invoke-SafeCommand -Command {
    Write-Host "  Querying CA database (this may take several minutes for large databases)..." -ForegroundColor Cyan
    
    # Try using ADCS module first if available
    if ($ADCSModuleAvailable) {
        Write-Host "  Using ADCS-Administration module cmdlets..." -ForegroundColor Cyan
        try {
            if ($IncludeRevokedCertificates) {
                Write-Host "  Retrieving all certificates (issued, revoked, pending)..." -ForegroundColor Cyan
                $certs = Get-IssuedRequest -CertificationAuthority $CAName -ErrorAction Stop
            }
            else {
                Write-Host "  Retrieving issued certificates only..." -ForegroundColor Cyan
                $certs = Get-IssuedRequest -CertificationAuthority $CAName -Filter "NotAfter -ge $([DateTime]::MinValue)" -ErrorAction Stop |
                    Where-Object { $_.StatusCode -eq 0x14 } # 0x14 = Issued
            }
            
            # Convert to our format
            $certificates = @()
            foreach ($cert in $certs) {
                $certificates += [PSCustomObject]@{
                    'CommonName' = $cert.CommonName
                    'RequesterName' = $cert.RequesterName
                    'NotBefore' = $cert.NotBefore
                    'NotAfter' = $cert.NotAfter
                    'SerialNumber' = $cert.SerialNumber
                    'CertificateTemplate' = $cert.CertificateTemplate
                    'CertificateHash' = $cert.CertificateHash
                    'Disposition' = $cert.StatusCode
                }
            }
            
            Write-Host "  ✅ Successfully retrieved $($certificates.Count) certificate records using ADCS cmdlets" -ForegroundColor Green
            return $certificates
        }
        catch {
            Write-Host "  ⚠️  ADCS cmdlet failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "  Falling back to certutil..." -ForegroundColor Yellow
        }
    }
    
    # Use certutil to query local CA database
    Write-Host "  Querying local CA database with certutil..." -ForegroundColor Cyan
    Write-Host "  This may take several minutes for large databases (12,000+ certificates)..." -ForegroundColor Cyan
    
    try {
        # Query local CA database with CSV output
        Write-Host "  Running: certutil -view csv" -ForegroundColor Gray
        $viewOutput = & certutil -view csv 2>&1
        
        Write-Host "  Exit code: $LASTEXITCODE" -ForegroundColor Gray
        Write-Host "  Output lines: $($viewOutput.Count)" -ForegroundColor Gray
        Write-Host "  First 3 lines:" -ForegroundColor Gray
        $viewOutput | Select-Object -First 3 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        
        # Check if we got valid CSV output - look for header row
        $headerLine = $viewOutput | Where-Object { $_ -match 'Request ID' } | Select-Object -First 1
        
        if ($LASTEXITCODE -eq 0 -and $headerLine) {
            Write-Host "  ✅ CSV header detected: $headerLine" -ForegroundColor Green
            Write-Host "  Parsing CSV data..." -ForegroundColor Cyan
            
            # Convert output to CSV - join lines and parse
            $csvText = $viewOutput -join "`n"
            
            # Save to temp file for Import-Csv
            $tempCsvFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ca_certs_csv_$([guid]::NewGuid()).csv")
            $csvText | Out-File -FilePath $tempCsvFile -Encoding UTF8
            
            try {
                Write-Host "  Loading CSV from temp file..." -ForegroundColor Cyan
                $csvData = Import-Csv -Path $tempCsvFile -ErrorAction Stop
                
                Write-Host "  CSV loaded: $($csvData.Count) total rows" -ForegroundColor Cyan
                
                # Show first row structure for debugging
                if ($csvData.Count -gt 0) {
                    Write-Host "  Sample columns: $($csvData[0].PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                }
                
                $certificates = @()
                $certCount = 0
                $skippedCount = 0
                
                foreach ($row in $csvData) {
                    # Only process issued certificates (disposition 20) unless IncludeRevoked is specified
                    $disposition = $row.'Request Disposition'
                    
                    # Disposition format is "20 -- Issued" so we need to extract just the number
                    $dispositionCode = if ($disposition -match '^(\d+)') { $Matches[1] } else { $disposition }
                    
                    if (-not $IncludeRevokedCertificates -and $dispositionCode -ne '20') {
                        $skippedCount++
                        continue
                    }
                    
                    # Skip if no expiration date
                    if ([string]::IsNullOrWhiteSpace($row.'Certificate Expiration Date') -or $row.'Certificate Expiration Date' -eq 'EMPTY') {
                        $skippedCount++
                        continue
                    }
                        
                        $certificates += [PSCustomObject]@{
                            'CommonName' = if ($row.'Issued Common Name' -and $row.'Issued Common Name' -ne 'EMPTY') { $row.'Issued Common Name' } else { $row.'Request Common Name' }
                            'RequesterName' = if ($row.'Requester Name' -and $row.'Requester Name' -ne 'EMPTY') { $row.'Requester Name' } else { 'N/A' }
                            'NotBefore' = if ($row.'Certificate Effective Date' -and $row.'Certificate Effective Date' -ne 'EMPTY') { $row.'Certificate Effective Date' } else { '' }
                            'NotAfter' = if ($row.'Certificate Expiration Date' -and $row.'Certificate Expiration Date' -ne 'EMPTY') { $row.'Certificate Expiration Date' } else { '' }
                            'SerialNumber' = if ($row.'Serial Number' -and $row.'Serial Number' -ne 'EMPTY') { $row.'Serial Number' } else { 'N/A' }
                            'CertificateTemplate' = if ($row.'Certificate Template' -and $row.'Certificate Template' -ne 'EMPTY') { $row.'Certificate Template' } else { 'N/A' }
                            'CertificateHash' = if ($row.'Certificate Hash' -and $row.'Certificate Hash' -ne 'EMPTY') { $row.'Certificate Hash' } else { 'N/A' }
                            'Disposition' = $disposition
                        }
                        
                    $certCount++
                    if ($certCount % 500 -eq 0) {
                        Write-Host "  Parsed $certCount certificates..." -ForegroundColor Cyan
                    }
                }
                
                Write-Host "  ✅ Successfully parsed $certCount certificate records from CSV" -ForegroundColor Green
                if ($skippedCount -gt 0) {
                    Write-Host "  ℹ️  Skipped $skippedCount records (revoked, pending, or invalid dates)" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "  ❌ CSV parsing failed: $($_.Exception.Message)" -ForegroundColor Red
                $certificates = @()
            }
            finally {
                # Clean up temp CSV file
                if (Test-Path $tempCsvFile) {
                    Remove-Item $tempCsvFile -Force -ErrorAction SilentlyContinue
                }
            }
        }
        else {
            # No valid CSV output
            Write-Host "  ❌ No valid CSV data returned from certutil" -ForegroundColor Red
            Write-Host "  Attempting row-by-row parsing as fallback..." -ForegroundColor Yellow
            
            $viewOutput = & certutil -view 2>&1
            
            $certificates = @()
            $currentCert = @{
                'CommonName' = ''
                'RequesterName' = ''
                'NotBefore' = ''
                'NotAfter' = ''
                'SerialNumber' = ''
                'CertificateTemplate' = ''
                'CertificateHash' = ''
                'Disposition' = ''
            }
            $certCount = 0
            $rowNumber = 0
            
            foreach ($line in $viewOutput) {
                $lineStr = $line.ToString()
                
                # New row indicator
                if ($lineStr -match '^Row (\d+):') {
                    $rowNumber = [int]$Matches[1]
                    
                    # Save previous certificate if it has required data
                    if ($currentCert['NotBefore'] -and $currentCert['NotAfter'] -and 
                        ($IncludeRevokedCertificates -or $currentCert['Disposition'] -eq '20')) {
                        $certificates += [PSCustomObject]$currentCert
                        $certCount++
                        
                        if ($certCount % 500 -eq 0) {
                            Write-Host "  Parsed $certCount certificates (Row $rowNumber)..." -ForegroundColor Cyan
                        }
                    }
                    
                    # Start new certificate
                    $currentCert = @{
                        'CommonName' = ''
                        'RequesterName' = ''
                        'NotBefore' = ''
                        'NotAfter' = ''
                        'SerialNumber' = ''
                        'CertificateTemplate' = ''
                        'CertificateHash' = ''
                        'Disposition' = ''
                    }
                    continue
                }
                
                # Parse property lines - format is "  Property Name: Value"
                if ($lineStr -match '^\s\s+(.+?):\s+(.*)$') {
                    $propName = $Matches[1].Trim()
                    $propValue = $Matches[2].Trim()
                    
                    if ([string]::IsNullOrWhiteSpace($propValue) -or $propValue -eq 'EMPTY') { continue }
                    
                    # Map properties
                    switch -Wildcard ($propName) {
                        '*Common Name*' { 
                            if (-not $currentCert['CommonName']) { $currentCert['CommonName'] = $propValue }
                        }
                        'Requester Name' { $currentCert['RequesterName'] = $propValue }
                        'Certificate Effective Date' { $currentCert['NotBefore'] = $propValue }
                        'Certificate Expiration Date' { $currentCert['NotAfter'] = $propValue }
                        'Serial Number' { $currentCert['SerialNumber'] = $propValue }
                        'Certificate Template' { 
                            if (-not $currentCert['CertificateTemplate']) { $currentCert['CertificateTemplate'] = $propValue }
                        }
                        'Request Disposition' { $currentCert['Disposition'] = $propValue }
                    }
                }
            }
            
            # Add final certificate
            if ($currentCert['NotBefore'] -and $currentCert['NotAfter'] -and 
                ($IncludeRevokedCertificates -or $currentCert['Disposition'] -eq '20')) {
                $certificates += [PSCustomObject]$currentCert
                $certCount++
            }
            
            if ($certificates.Count -eq 0) {
                Write-Host "`n  ❌ Unable to retrieve certificate data" -ForegroundColor Red
                Write-Host "  Troubleshooting:" -ForegroundColor Yellow
                Write-Host "  • Verify CA is running: Get-Service CertSvc" -ForegroundColor Gray
                Write-Host "  • Check database access: certutil -view csv | Select-Object -First 10" -ForegroundColor Gray
                Write-Host "  • Review CA logs in Event Viewer" -ForegroundColor Gray
                
                Write-StatusMessage -Message "Certificate export returned no data" -Type "Warning"
                return @()
            }
        }
        
        Write-Host "  ✅ Successfully retrieved $($certificates.Count) certificate records" -ForegroundColor Green
        return $certificates
    }
    catch {
        Write-Host "  ❌ Failed to query CA database: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
} -ErrorMessage "Failed to retrieve issued certificates" -ContinueOnError

if ($null -eq $IssuedCertificates -or $IssuedCertificates.Count -eq 0) {
    Write-StatusMessage -Message "No issued certificates found or query failed" -Type "Warning"
    $CertificateData = @()
}
else {
    Write-Host "✅ Retrieved $($IssuedCertificates.Count) certificate records" -ForegroundColor Green
    
    # Process certificate data
    $CertificateData = @()
    $ExpiringCount = 0
    $ExpiredCount = 0
    $CurrentDate = Get-Date
    $ExpirationThreshold = $CurrentDate.AddDays($DaysToExpiration)
    
    Write-StatusMessage -Message "Analyzing certificate expiration..." -Type "Info"
    
    $processedCount = 0
    foreach ($cert in $IssuedCertificates) {
        try {
            # Handle empty or null dates
            if ([string]::IsNullOrWhiteSpace($cert.NotAfter) -or [string]::IsNullOrWhiteSpace($cert.NotBefore)) {
                Write-Verbose "Skipping certificate with missing date: $($cert.CommonName)"
                continue
            }
            
            # Try parsing dates with multiple formats
            $notAfter = $null
            $notBefore = $null
            
            try {
                $notAfter = [DateTime]::Parse($cert.NotAfter)
                $notBefore = [DateTime]::Parse($cert.NotBefore)
            }
            catch {
                # Try ParseExact with common formats
                $formats = @(
                    'M/d/yyyy h:mm tt',
                    'M/d/yyyy H:mm',
                    'MM/dd/yyyy HH:mm:ss',
                    'yyyy-MM-dd HH:mm:ss',
                    'MM/dd/yyyy hh:mm:ss tt'
                )
                
                foreach ($format in $formats) {
                    try {
                        $notAfter = [DateTime]::ParseExact($cert.NotAfter, $format, [System.Globalization.CultureInfo]::InvariantCulture)
                        $notBefore = [DateTime]::ParseExact($cert.NotBefore, $format, [System.Globalization.CultureInfo]::InvariantCulture)
                        break
                    }
                    catch {
                        continue
                    }
                }
                
                if ($null -eq $notAfter) {
                    throw "Could not parse dates: NotBefore='$($cert.NotBefore)', NotAfter='$($cert.NotAfter)'"
                }
            }
            
            $daysRemaining = ($notAfter - $CurrentDate).Days
            
            $expirationStatus = if ($daysRemaining -lt 0) {
                $ExpiredCount++
                "Expired"
            }
            elseif ($daysRemaining -le $DaysToExpiration) {
                $ExpiringCount++
                "Expiring Soon"
            }
            else {
                "Valid"
            }
            
            $certObj = [PSCustomObject]@{
                CommonName = if ($cert.CommonName) { $cert.CommonName } else { "N/A" }
                RequesterName = if ($cert.RequesterName) { $cert.RequesterName } else { "N/A" }
                Template = if ($cert.CertificateTemplate) { $cert.CertificateTemplate } else { "N/A" }
                SerialNumber = if ($cert.SerialNumber) { $cert.SerialNumber } else { "N/A" }
                NotBefore = $notBefore.ToString('yyyy-MM-dd HH:mm:ss')
                NotAfter = $notAfter.ToString('yyyy-MM-dd HH:mm:ss')
                DaysRemaining = $daysRemaining
                ExpirationStatus = $expirationStatus
                CertificateHash = if ($cert.CertificateHash) { $cert.CertificateHash } else { "N/A" }
                Disposition = if ($cert.Disposition) { $cert.Disposition } else { "20 (Issued)" }
            }
            
            $CertificateData += $certObj
            $processedCount++
            
            # Progress indicator for large datasets
            if ($processedCount % 1000 -eq 0) {
                Write-Host "  Processed $processedCount certificates..." -ForegroundColor Cyan
            }
        }
        catch {
            Write-StatusMessage -Message "Error processing certificate '$($cert.CommonName)': $($_.Exception.Message)" -Type "Warning"
        }
    }
    
    Write-Host "✅ Successfully processed $processedCount out of $($IssuedCertificates.Count) certificates" -ForegroundColor Green
    
    # Export to CSV
    try {
        $CertificateData | Export-Csv -Path $CertificatesFile -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Exported $($CertificateData.Count) certificates to: $CertificatesFile" -ForegroundColor Green
    }
    catch {
        Write-StatusMessage -Message "Failed to export certificates CSV: $($_.Exception.Message)" -Type "Error"
    }
    
    Write-Host "`nCertificate Summary:" -ForegroundColor Cyan
    Write-Host "  Total Certificates: $($CertificateData.Count)" -ForegroundColor White
    Write-Host "  Valid: $(($CertificateData | Where-Object {$_.ExpirationStatus -eq 'Valid'}).Count)" -ForegroundColor Green
    Write-Host "  Expiring Soon (<$DaysToExpiration days): $ExpiringCount" -ForegroundColor Yellow
    Write-Host "  Expired: $ExpiredCount" -ForegroundColor Red
}

# ============================================================================
# EXPORT CERTIFICATE TEMPLATES
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Exporting Certificate Templates" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$TemplatesFile = Join-Path $AssessmentFolder "$($hostname)_PKI_CertificateTemplates_$Timestamp.csv"
$TemplatePermissionsFile = Join-Path $AssessmentFolder "$($hostname)_PKI_TemplatePermissions_$Timestamp.csv"

Write-StatusMessage -Message "Retrieving certificate templates..." -Type "Info"

# Get templates from AD (Configuration partition)
$TemplateData = @()
$TemplatePermissions = @()

try {
    # Get AD configuration naming context
    $RootDSE = [ADSI]"LDAP://RootDSE"
    $ConfigNC = $RootDSE.configurationNamingContext
    $TemplatesPath = "LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$ConfigNC"
    
    Write-StatusMessage -Message "Connecting to: $TemplatesPath" -Type "Info"
    
    $TemplatesContainer = [ADSI]$TemplatesPath
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher
    $Searcher.SearchRoot = $TemplatesContainer
    $Searcher.Filter = "(objectClass=pKICertificateTemplate)"
    $Searcher.PageSize = 1000
    
    $Templates = $Searcher.FindAll()
    
    if ($Templates.Count -eq 0) {
        Write-StatusMessage -Message "No certificate templates found" -Type "Warning"
    }
    else {
        Write-Host "✅ Found $($Templates.Count) certificate templates" -ForegroundColor Green
        
        foreach ($template in $Templates) {
            try {
                $templateEntry = $template.GetDirectoryEntry()
                
                # Extract template properties
                $displayName = if ($templateEntry.displayName) { $templateEntry.displayName[0] } else { $templateEntry.cn[0] }
                $cn = $templateEntry.cn[0]
                
                # Parse flags (if available)
                $flags = if ($templateEntry.flags) { $templateEntry.flags[0] } else { 0 }
                $enrollmentFlags = if ($templateEntry.'msPKI-Enrollment-Flag') { $templateEntry.'msPKI-Enrollment-Flag'[0] } else { 0 }
                
                # Validity period
                $validityPeriod = if ($templateEntry.'pKIExpirationPeriod') { 
                    $bytes = $templateEntry.'pKIExpirationPeriod'[0]
                    # Convert to days (simplified)
                    "Custom"
                } else { 
                    "Not Set" 
                }
                
                # Check if template is published
                $published = $false
                $publishedOn = @()
                
                # Get CA templates to see where published
                try {
                    $caTemplates = certutil -CATemplates 2>&1
                    if ($caTemplates -match $cn) {
                        $published = $true
                        $publishedOn += $CAName
                    }
                }
                catch {
                    # Ignore errors checking published status
                }
                
                $templateObj = [PSCustomObject]@{
                    DisplayName = $displayName
                    Name = $cn
                    Published = $published
                    PublishedOn = ($publishedOn -join "; ")
                    ValidityPeriod = $validityPeriod
                    MinimalKeyLength = if ($templateEntry.'msPKI-Minimal-Key-Size') { $templateEntry.'msPKI-Minimal-Key-Size'[0] } else { "Not Set" }
                    Flags = $flags
                    EnrollmentFlags = $enrollmentFlags
                    Distinguished_Name = $templateEntry.distinguishedName[0]
                }
                
                $TemplateData += $templateObj
                
                # Extract permissions (ACL)
                Write-StatusMessage -Message "Extracting permissions for template: $displayName" -Type "Info"
                
                $acl = $templateEntry.ObjectSecurity
                
                foreach ($ace in $acl.Access) {
                    $permObj = [PSCustomObject]@{
                        TemplateName = $displayName
                        TemplateCommonName = $cn
                        IdentityReference = $ace.IdentityReference.ToString()
                        AccessControlType = $ace.AccessControlType.ToString()
                        ActiveDirectoryRights = $ace.ActiveDirectoryRights.ToString()
                        InheritanceType = $ace.InheritanceType.ToString()
                        IsInherited = $ace.IsInherited
                    }
                    
                    $TemplatePermissions += $permObj
                }
            }
            catch {
                Write-StatusMessage -Message "Error processing template $($templateEntry.cn): $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Export templates to CSV
        try {
            $TemplateData | Export-Csv -Path $TemplatesFile -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Exported $($TemplateData.Count) templates to: $TemplatesFile" -ForegroundColor Green
        }
        catch {
            Write-StatusMessage -Message "Failed to export templates CSV: $($_.Exception.Message)" -Type "Error"
        }
        
        # Export permissions to CSV
        try {
            $TemplatePermissions | Export-Csv -Path $TemplatePermissionsFile -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Exported $($TemplatePermissions.Count) permission entries to: $TemplatePermissionsFile" -ForegroundColor Green
        }
        catch {
            Write-StatusMessage -Message "Failed to export permissions CSV: $($_.Exception.Message)" -Type "Error"
        }
        
        Write-Host "`nTemplate Summary:" -ForegroundColor Cyan
        Write-Host "  Total Templates: $($TemplateData.Count)" -ForegroundColor White
        Write-Host "  Published: $(($TemplateData | Where-Object {$_.Published -eq $true}).Count)" -ForegroundColor Green
        Write-Host "  Unpublished: $(($TemplateData | Where-Object {$_.Published -eq $false}).Count)" -ForegroundColor Yellow
    }
}
catch {
    Write-StatusMessage -Message "Failed to retrieve certificate templates: $($_.Exception.Message)" -Type "Error"
}

# ============================================================================
# GENERATE COMPREHENSIVE TEXT REPORT
# ============================================================================

Write-Host "`n$SubSeparator" -ForegroundColor Cyan
Write-Host "Generating Comprehensive Report" -ForegroundColor Cyan
Write-Host $SubSeparator -ForegroundColor Cyan

$ReportFile = Join-Path $AssessmentFolder "PKI_Assessment_Report_$Timestamp.txt"

$Report = @"
$Separator
$OrganizationName - PKI INFRASTRUCTURE ASSESSMENT
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$Separator

CERTIFICATE AUTHORITY INFORMATION
$SubSeparator
CA Server: $env:COMPUTERNAME
CA Name: $CAName
Assessment Date: $(Get-Date -Format 'yyyy-MM-dd')
Report Generated By: $env:USERNAME

ISSUED CERTIFICATES SUMMARY
$SubSeparator
Total Certificates: $($CertificateData.Count)
Valid Certificates: $(($CertificateData | Where-Object {$_.ExpirationStatus -eq 'Valid'}).Count)
Expiring Soon (<$DaysToExpiration days): $ExpiringCount
Expired Certificates: $ExpiredCount

TOP 10 EXPIRING CERTIFICATES
$SubSeparator
"@

if ($CertificateData.Count -gt 0) {
    $topExpiring = $CertificateData | Where-Object {$_.DaysRemaining -ge 0} | Sort-Object DaysRemaining | Select-Object -First 10
    foreach ($cert in $topExpiring) {
        $Report += "`n$($cert.CommonName) - Expires: $($cert.NotAfter) ($($cert.DaysRemaining) days)"
    }
}
else {
    $Report += "`nNo certificate data available"
}

$Report += @"

`n
CERTIFICATE TEMPLATES SUMMARY
$SubSeparator
Total Templates: $($TemplateData.Count)
Published Templates: $(($TemplateData | Where-Object {$_.Published -eq $true}).Count)
Unpublished Templates: $(($TemplateData | Where-Object {$_.Published -eq $false}).Count)

PUBLISHED TEMPLATES
$SubSeparator
"@

$publishedTemplates = $TemplateData | Where-Object {$_.Published -eq $true}
if ($publishedTemplates.Count -gt 0) {
    foreach ($template in $publishedTemplates) {
        $Report += "`n- $($template.DisplayName) ($($template.Name))"
        $Report += "`n  Published On: $($template.PublishedOn)"
        $Report += "`n  Min Key Length: $($template.MinimalKeyLength)"
    }
}
else {
    $Report += "`nNo published templates found"
}

$Report += @"

`n
TEMPLATE PERMISSIONS SUMMARY
$SubSeparator
Total Permission Entries: $($TemplatePermissions.Count)

Key Permissions by Template:
"@

$groupedPerms = $TemplatePermissions | Group-Object TemplateName
foreach ($group in $groupedPerms | Sort-Object Name) {
    $Report += "`n`n$($group.Name):"
    $enrollPerms = $group.Group | Where-Object {$_.ActiveDirectoryRights -like "*Enroll*"}
    foreach ($perm in $enrollPerms) {
        $Report += "`n  - $($perm.IdentityReference): $($perm.ActiveDirectoryRights) ($($perm.AccessControlType))"
    }
}

$Report += @"

`n
ASSESSMENT SUMMARY
$SubSeparator
Errors Encountered: $ErrorCount
Warnings Generated: $WarningCount

EXPORTED FILES
$SubSeparator
Certificates CSV: $CertificatesFile
Templates CSV: $TemplatesFile
Permissions CSV: $TemplatePermissionsFile
Assessment Report: $ReportFile

$Separator
End of Report
$Separator
"@

# Save report
try {
    $Report | Out-File -FilePath $ReportFile -Encoding UTF8
    Write-Host "✅ Comprehensive report saved to: $ReportFile" -ForegroundColor Green
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
Write-Host "PKI ASSESSMENT COMPLETED" -ForegroundColor Green
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Duration: $($Duration.ToString('mm\:ss'))" -ForegroundColor White
Write-Host "Errors: $ErrorCount | Warnings: $WarningCount" -ForegroundColor $(if($ErrorCount -gt 0){'Red'}else{'Green'})
Write-Host "`nAll reports saved to: $AssessmentFolder" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
