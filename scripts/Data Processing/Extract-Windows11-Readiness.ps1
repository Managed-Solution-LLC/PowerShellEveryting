# Excel Column Data Extractor - Final Working Version
# This script processes Windows 11 Readiness Check results from Excel files

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$WorksheetName = "Sheet1",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "Windows11_Readiness_Results.csv"
)

# Import ImportExcel module
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}
Import-Module ImportExcel -Force

function Get-StatusData {
    param([string]$RawData)
    
    $result = @{
        ReadinessStatus = ""
        StorageResult = ""
        MemoryResult = ""
        TPMResult = ""
        ProcessorResult = ""
        SecureBootResult = ""
        OSVersionResult = ""
        StorageDetails = ""
        MemoryDetails = ""
        TPMDetails = ""
        ProcessorDetails = ""
        SecureBootDetails = ""
        OSDetails = ""
    }
    
    if ([string]::IsNullOrWhiteSpace($RawData)) { return $result }
    
    # Extract overall status
    if ($RawData -match "Status\s*:\s*(\w+)") {
        $result.ReadinessStatus = $matches[1]
    }
    
    # Extract individual component results using regex
    if ($RawData -match "Storage:[^:]*::\s*(\w+)") {
        $result.StorageResult = $matches[1]
        if ($RawData -match "(Storage:[^:]*::\s*\w+)") {
            $result.StorageDetails = $matches[1]
        }
    }
    
    if ($RawData -match "Memory:[^:]*::\s*(\w+)") {
        $result.MemoryResult = $matches[1]
        if ($RawData -match "(Memory:[^:]*::\s*\w+)") {
            $result.MemoryDetails = $matches[1]
        }
    }
    
    if ($RawData -match "TPM:[^:]*::\s*(\w+)") {
        $result.TPMResult = $matches[1]
        if ($RawData -match "(TPM:[^:]*::\s*\w+)") {
            $result.TPMDetails = $matches[1]
        }
    }
    
    if ($RawData -match "SecureBoot:[^:]*::\s*(\w+)") {
        $result.SecureBootResult = $matches[1]
        if ($RawData -match "(SecureBoot:[^:]*::\s*\w+)") {
            $result.SecureBootDetails = $matches[1]
        }
    }
    
    if ($RawData -match "OsVersion:[^:]*::\s*(\w+)") {
        $result.OSVersionResult = $matches[1]
        if ($RawData -match "(OsVersion:[^:]*::\s*\w+)") {
            $result.OSDetails = $matches[1]
        }
    }
    
    # Extract processor info (more complex due to multi-line format)
    if ($RawData -match "Processor:\s*\{([^}]*)\}\s*::\s*(\w+)") {
        $result.ProcessorResult = $matches[2]
        $result.ProcessorDetails = "Processor: {$($matches[1].Trim())} :: $($matches[2])"
    }
    
    return $result
}

try {
    Write-Host "Processing Excel file: $ExcelFilePath" -ForegroundColor Cyan
    
    # Import data
    $data = Import-Excel -Path $ExcelFilePath -WorksheetName $WorksheetName
    Write-Host "Found $($data.Count) rows" -ForegroundColor Green
    
    # Process each row
    $results = foreach ($row in $data) {
        $statusData = Get-StatusData -RawData $row.Output
        
        [PSCustomObject]@{
            MachineName = $row.'Machine name'
            LastRunTime = $row.'Last run time'
            Windows11ReadinessStatus = $statusData.ReadinessStatus
            CommandExecutionStatus = $row.Status
            FailureSource = $row.'Failure Source'
            StorageResult = $statusData.StorageResult
            MemoryResult = $statusData.MemoryResult
            TPMResult = $statusData.TPMResult
            ProcessorResult = $statusData.ProcessorResult
            SecureBootResult = $statusData.SecureBootResult
            OSVersionResult = $statusData.OSVersionResult
            StorageDetails = $statusData.StorageDetails
            MemoryDetails = $statusData.MemoryDetails
            TPMDetails = $statusData.TPMDetails
            ProcessorDetails = $statusData.ProcessorDetails
            SecureBootDetails = $statusData.SecureBootDetails
            OSDetails = $statusData.OSDetails
        }
    }
    
    # Export results
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
    
    # Show summary
    Write-Host "`nWindows 11 Readiness Summary:" -ForegroundColor Yellow
    $results | Group-Object Windows11ReadinessStatus | Select-Object Name, Count | Format-Table
    
    Write-Host "Component Failure Analysis:" -ForegroundColor Yellow
    $failures = @(
        @{Component="Storage"; Failed=($results | Where-Object StorageResult -eq "FAIL").Count}
        @{Component="Memory"; Failed=($results | Where-Object MemoryResult -eq "FAIL").Count}
        @{Component="TPM"; Failed=($results | Where-Object TPMResult -eq "FAIL").Count}
        @{Component="Processor"; Failed=($results | Where-Object ProcessorResult -eq "FAIL").Count}
        @{Component="SecureBoot"; Failed=($results | Where-Object SecureBootResult -eq "FAIL").Count}
        @{Component="OS"; Failed=($results | Where-Object OSVersionResult -eq "FAIL").Count}
    )
    
    $failures | ForEach-Object { [PSCustomObject]$_ } | Format-Table
    
    Write-Host "Sample of machines with issues:" -ForegroundColor Yellow
    $results | Where-Object Windows11ReadinessStatus -eq "Unsupported" | 
        Select-Object MachineName, FailureSource, TPMResult, ProcessorResult, SecureBootResult |
        Format-Table -AutoSize
    
    return $results
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
}
