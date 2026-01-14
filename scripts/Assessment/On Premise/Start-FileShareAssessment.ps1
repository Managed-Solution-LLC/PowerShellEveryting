<#
.SYNOPSIS
    Comprehensive file share assessment tool that analyzes SMB shares and generates an Excel report.

.DESCRIPTION
    This script performs a complete file share assessment on the local server including:
    - Share inventory and size analysis
    - NTFS permissions for all folders
    - Files with SharePoint/OneDrive unsupported characters
    - Consolidates all data into a formatted Excel workbook
    
    Designed to run directly on the file server with administrative privileges.

.PARAMETER Domain
    The domain name for the assessment. Used in report naming.

.PARAMETER OutputDirectory
    Directory where CSV files and Excel report will be saved. Defaults to current directory.

.PARAMETER ExcludeShares
    Array of share names to exclude from assessment (e.g., administrative shares).

.PARAMETER SkipPermissions
    Skip the permissions analysis (faster but less complete).

.PARAMETER Workers
    Number of parallel workers for permission scanning. Default is 50.

.EXAMPLE
    .\Start-FileShareAssessment.ps1 -Domain "contoso"
    
    Runs assessment on all shares and creates contoso_File_Share_Assessment.xlsx

.EXAMPLE
    .\Start-FileShareAssessment.ps1 -Domain "contoso" -ExcludeShares "ADMIN$","IPC$" -Workers 100
    
    Excludes administrative shares and uses 100 parallel workers for faster processing.

.EXAMPLE
    .\Start-FileShareAssessment.ps1 -Domain "contoso" -OutputDirectory "C:\Reports" -SkipPermissions
    
    Saves output to C:\Reports and skips permission analysis.

.NOTES
    Author: W. Ford
    Company: Managed Solution LLC
    Date: 2026-01-05
    Version: 1.0
    
    Requirements:
    - PowerShell 5.1 or later
    - Administrator privileges on file server
    - ImportExcel module (auto-installed if missing)
    - Long paths enabled (script will enable if needed)
    
    Performance Notes:
    - Permissions analysis can be slow on large shares
    - Consider using -SkipPermissions for initial quick assessment
    - Increase -Workers parameter for better performance on large environments

.LINK
    https://github.com/Managed-Solution-LLC/PowerShellEveryting
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Domain name for the assessment")]
    [string]$Domain,
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for reports")]
    [string]$OutputDirectory = ".",
    
    [Parameter(Mandatory=$false, HelpMessage="Share names to exclude")]
    [string[]]$ExcludeShares = @("ADMIN$", "IPC$", "C$", "D$", "E$", "F$"),
    
    [Parameter(Mandatory=$false, HelpMessage="Skip permission analysis")]
    [switch]$SkipPermissions,
    
    [Parameter(Mandatory=$false, HelpMessage="Number of parallel workers")]
    [int]$Workers = 50
)

#Requires -RunAsAdministrator

# Initialize variables
$script:ErrorCount = 0
$script:WarningCount = 0
$StartTime = Get-Date
$Separator = "=" * 80
$SubSeparator = "-" * 60

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Level) {
        'Error' { 
            Write-Host "[$timestamp] ❌ $Message" -ForegroundColor Red
            $script:ErrorCount++
        }
        'Warning' { 
            Write-Host "[$timestamp] ⚠️  $Message" -ForegroundColor Yellow
            $script:WarningCount++
        }
        'Success' { 
            Write-Host "[$timestamp] ✅ $Message" -ForegroundColor Green
        }
        'Info' { 
            Write-Host "[$timestamp] ℹ️  $Message" -ForegroundColor Cyan
        }
    }
}

function Test-LongPathsEnabled {
    try {
        $longPathsEnabled = (Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -ErrorAction Stop).LongPathsEnabled
        return ($longPathsEnabled -eq 1)
    }
    catch {
        return $false
    }
}

function Enable-LongPaths {
    try {
        Set-ItemProperty 'HKLM:\System\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1 -ErrorAction Stop
        Write-Log "Long paths enabled successfully" -Level Success
        Write-Log "Note: A system restart may be required for changes to take full effect" -Level Warning
        return $true
    }
    catch {
        Write-Log "Failed to enable long paths: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Get-LocalSMBShares {
    try {
        Write-Log "Discovering SMB shares on local server..." -Level Info
        $shares = Get-SmbShare | Where-Object { 
            $_.Name -notin $ExcludeShares -and 
            $_.Special -eq $false 
        }
        
        Write-Log "Found $($shares.Count) shares to assess" -Level Success
        return $shares
    }
    catch {
        Write-Log "Failed to enumerate SMB shares: $($_.Exception.Message)" -Level Error
        return @()
    }
}

function Export-FileShareSize {
    param(
        [Parameter(Mandatory=$true)]
        $Share
    )
    
    $sharePath = $Share.Path
    $shareName = $Share.Name
    $results = @()

    Write-Log "Analyzing share: $shareName ($sharePath)" -Level Info
    
    try {
        $folders = Get-ChildItem -Path $sharePath -Directory -ErrorAction Stop
        
        foreach ($folder in $folders) {
            try {
                Write-Host "  |-- Processing $($folder.Name)..." -NoNewline
                $size = (Get-ChildItem -Path $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue | 
                        Measure-Object -Property Length -Sum).Sum
                $folderSizeInGB = [math]::Round($size / 1GB, 2)
                $files = @(Get-ChildItem -Path $folder.FullName -File -Recurse -ErrorAction SilentlyContinue)
                $subfolders = @(Get-ChildItem -Path $folder.FullName -Directory -Recurse -ErrorAction SilentlyContinue)
                
                Write-Host " $folderSizeInGB GB" -ForegroundColor Green
                
                $results += [PSCustomObject]@{
                    FolderName = $folder.Name
                    FolderSizeGB = $folderSizeInGB
                    TotalFolders = $subfolders.Count
                    TotalFiles = $files.Count
                }
            }
            catch {
                Write-Host " Error" -ForegroundColor Red
                Write-Log "  Error processing folder $($folder.Name): $($_.Exception.Message)" -Level Warning
            }
        }
        
        # Process root level files
        Write-Host "  |-- Processing Root..." -NoNewline
        try {
            $preSum = ($results | Measure-Object -Property FolderSizeGB -Sum).Sum
            $rootFiles = @(Get-ChildItem -Path $sharePath -File -ErrorAction Stop)
            $rootFolders = @(Get-ChildItem -Path $sharePath -Directory -ErrorAction Stop)
            $rootSize = (Get-ChildItem -Path $sharePath -Recurse -Force -ErrorAction Stop | 
                        Measure-Object -Property Length -Sum).Sum
            $rootFolderSizeInGB = [math]::Round([decimal]$rootSize / 1GB, 2)
            
            $results += [PSCustomObject]@{
                FolderName = 'Root'
                FolderSizeGB = $rootFolderSizeInGB - $preSum
                TotalFolders = $rootFolders.Count
                TotalFiles = $rootFiles.Count
            }
            Write-Host " $($rootFolderSizeInGB)GB" -ForegroundColor Green
        }
        catch {
            Write-Host " Error" -ForegroundColor Red
            Write-Log "  Error processing root: $($_.Exception.Message)" -Level Warning
        }
        
        # Add totals
        $totalFolderSizeGB = ($results | Measure-Object -Property FolderSizeGB -Sum).Sum
        $totalFolders = ($results | Measure-Object -Property TotalFolders -Sum).Sum
        $totalFiles = ($results | Measure-Object -Property TotalFiles -Sum).Sum
        
        $results += [PSCustomObject]@{
            FolderName = 'Total'
            FolderSizeGB = $totalFolderSizeGB
            TotalFolders = $totalFolders
            TotalFiles = $totalFiles
        }
        
        # Export to CSV
        $outputPath = Join-Path $OutputDirectory "fileaudit_$($shareName).csv"
        $results | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Log "Share size analysis completed: $outputPath" -Level Success
    }
    catch {
        Write-Log "Failed to analyze share ${shareName} : $($_.Exception.Message)" -Level Error
    }
}

function Export-FileSharePermissions {
    param(
        [Parameter(Mandatory=$true)]
        $Share
    )
    
    $sharePath = $Share.Path
    $shareName = $Share.Name
    
    Write-Log "Analyzing permissions for share: $shareName" -Level Info
    
    try {
        $folders = Get-ChildItem -Path $sharePath -Directory -ErrorAction Stop
        
        foreach ($folder in $folders) {
            Write-Host "  |-- Processing $($folder.Name) permissions..." -NoNewline
            
            $folderPermissions = @()
            $folderShare = Get-ChildItem -Path $folder.FullName -Directory -Recurse -ErrorAction SilentlyContinue
            
            $totalFolders = $folderShare.Count
            $currentFolder = 0
            
            foreach ($subfolder in $folderShare) {
                $currentFolder++
                $percentComplete = [math]::Round(($currentFolder / $totalFolders) * 100, 2)
                Write-Progress -Activity "Processing folder permissions" -Status "$percentComplete% Complete" -PercentComplete $percentComplete -CurrentOperation "Processing $($subfolder.FullName)"
                
                try {
                    $acl = Get-Acl -Path $subfolder.FullName -ErrorAction Stop
                    $acl.Access | ForEach-Object {
                        $folderPermissions += [PSCustomObject]@{
                            SharePath = $sharePath
                            FolderName = $subfolder.FullName.Substring($sharePath.Length + 1)
                            IdentityReference = $_.IdentityReference
                            FileSystemRights = $_.FileSystemRights
                            AccessControlType = $_.AccessControlType
                            IsInherited = $_.IsInherited
                        }
                    }
                }
                catch {
                    Write-Verbose "Error getting ACL for: $($subfolder.FullName)"
                }
            }
            
            Write-Progress -Activity "Processing folder permissions" -Completed
            
            # Export permissions
            $rawDataPath = Join-Path $OutputDirectory "RawData"
            if (-not (Test-Path $rawDataPath)) {
                New-Item -ItemType Directory -Path $rawDataPath -Force | Out-Null
            }
            
            $outputPath = Join-Path $rawDataPath "permissions_$($shareName)_$($folder.Name).csv"
            $folderPermissions | Export-Csv -Path $outputPath -NoTypeInformation
            Write-Host " Completed" -ForegroundColor Green
        }
        
        Write-Log "Permissions analysis completed for $shareName" -Level Success
    }
    catch {
        Write-Log "Failed to analyze permissions for ${shareName}: $($_.Exception.Message)" -Level Error
    }
}

function Export-FileShareUnsupportedFileNames {
    param(
        [Parameter(Mandatory=$true)]
        $Share
    )
    
    $sharePath = $Share.Path
    $shareName = $Share.Name
    
    # SharePoint/OneDrive unsupported characters
    $unsupportedChars = '[~#%&*{}\\:<>?/|"]'
    
    Write-Log "Scanning for unsupported filenames in: $shareName" -Level Info
    Write-Host "  |-- Scanning files..." -NoNewline
    
    try {
        $files = Get-ChildItem -Path $sharePath -Recurse -File -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -match $unsupportedChars } |
                Select-Object Name, Directory, FullName
        
        $outputPath = Join-Path $OutputDirectory "unsupported_filenames_$($shareName).csv"
        
        if ($files) {
            $files | Export-Csv -Path $outputPath -NoTypeInformation
            Write-Host " Found $($files.Count) files" -ForegroundColor Yellow
            Write-Log "Found $($files.Count) files with unsupported characters" -Level Warning
        }
        else {
            # Export empty file to maintain consistency
            @() | Export-Csv -Path $outputPath -NoTypeInformation
            Write-Host " No issues found" -ForegroundColor Green
            Write-Log "No files with unsupported characters found" -Level Success
        }
    }
    catch {
        Write-Host " Error" -ForegroundColor Red
        Write-Log "Failed to scan for unsupported filenames: $($_.Exception.Message)" -Level Error
    }
}

function Install-RequiredModules {
    Write-Log "Checking required modules..." -Level Info
    
    # Check for ImportExcel module
    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        Write-Log "ImportExcel module not found. Installing..." -Level Info
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "ImportExcel module installed successfully" -Level Success
        }
        catch {
            Write-Log "Failed to install ImportExcel module: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    else {
        Write-Log "ImportExcel module already installed" -Level Success
    }
    
    # Import the module
    try {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Log "ImportExcel module loaded" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to load ImportExcel module: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Export-UniquePermissions {
    Write-Log "Extracting unique permissions from RawData..." -Level Info
    
    $rawDataPath = Join-Path $OutputDirectory "RawData"
    
    if (-not (Test-Path $rawDataPath)) {
        Write-Log "No RawData folder found, skipping unique permissions extraction" -Level Warning
        return
    }
    
    $permissionFiles = Get-ChildItem -Path $rawDataPath -Filter "permissions_*.csv" -File -ErrorAction SilentlyContinue
    
    if ($permissionFiles.Count -eq 0) {
        Write-Log "No permission files found in RawData folder" -Level Warning
        return
    }
    
    Write-Log "Processing $($permissionFiles.Count) permission files..." -Level Info
    
    try {
        # Collect all permissions
        $allPermissions = @()
        foreach ($file in $permissionFiles) {
            $data = Import-Csv -Path $file.FullName -ErrorAction Stop
            if ($data) {
                $allPermissions += $data
            }
        }
        
        if ($allPermissions.Count -eq 0) {
            Write-Log "No permissions data found in CSV files" -Level Warning
            return
        }
        
        Write-Log "Collected $($allPermissions.Count) total permission entries" -Level Info
        
        # Analyze folder-level permissions for SharePoint mapping
        Write-Log "Analyzing folder hierarchy for SharePoint mapping..." -Level Info
        
        # Group permissions by folder path
        $folderPermissions = $allPermissions | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Path) } | Group-Object Path | ForEach-Object {
            $folderPath = $_.Name
            $folderPerms = $_.Group
            
            # Create permission signature for this folder (non-inherited permissions)
            $explicitPerms = $folderPerms | Where-Object { $_.IsInherited -eq $false }
            
            if ($explicitPerms) {
                $permSignature = ($explicitPerms | 
                    Sort-Object IdentityReference, FileSystemRights, AccessControlType | 
                    ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)" }) -join ';'
            } else {
                $permSignature = "INHERITED_ONLY"
            }
            
            [PSCustomObject]@{
                Path = $folderPath
                Depth = ($folderPath -split '\\').Count
                PermissionSignature = $permSignature
                ExplicitPermissionCount = $explicitPerms.Count
                TotalPermissionCount = $folderPerms.Count
                UniqueIdentities = ($explicitPerms | Select-Object -ExpandProperty IdentityReference -Unique | Measure-Object).Count
            }
        }
        
        Write-Log "Analyzed $($folderPermissions.Count) unique folder paths" -Level Info
        
        # Find folders with unique permission boundaries (lowest level with explicit permissions)
        $uniquePermissionBoundaries = @()
        $processedSignatures = @{}
        
        foreach ($folder in ($folderPermissions | Sort-Object Depth -Descending)) {
            if ($folder.ExplicitPermissionCount -gt 0 -and $folder.PermissionSignature -ne "INHERITED_ONLY") {
                # Track unique permission signatures at lowest levels
                if (-not $processedSignatures.ContainsKey($folder.PermissionSignature)) {
                    $uniquePermissionBoundaries += $folder
                    $processedSignatures[$folder.PermissionSignature] = $folder.Path
                }
            }
        }
        
        Write-Log "Identified $($uniquePermissionBoundaries.Count) unique permission boundaries" -Level Success
        
        # Create SharePoint mapping recommendations
        $sharepointMapping = foreach ($boundary in $uniquePermissionBoundaries) {
            $boundaryPath = $boundary.Path
            $perms = $allPermissions | Where-Object { $_.Path -eq $boundaryPath -and $_.IsInherited -eq $false }
            
            $identityList = ($perms | Select-Object -ExpandProperty IdentityReference -Unique | Sort-Object) -join '; '
            $rightsList = ($perms | Select-Object FileSystemRights, AccessControlType -Unique | 
                ForEach-Object { "$($_.AccessControlType): $($_.FileSystemRights)" }) -join '; '
            
            # Suggest SharePoint site structure
            $pathParts = $boundaryPath -split '\\'
            $shareName = if ($pathParts.Count -gt 2) { $pathParts[2] } else { "Root" }
            $relativePath = if ($pathParts.Count -gt 3) { 
                ($pathParts[3..($pathParts.Count-1)] -join '/')
            } else { 
                "Root" 
            }
            
            [PSCustomObject]@{
                FolderPath = $boundaryPath
                ShareName = $shareName
                RelativePath = $relativePath
                FolderDepth = $boundary.Depth
                UniqueIdentities = $boundary.UniqueIdentities
                Identities = $identityList
                Permissions = $rightsList
                PermissionSignature = $boundary.PermissionSignature
                RecommendedAction = if ($boundary.Depth -le 4) { 
                    "Create as SharePoint Site or Document Library" 
                } else { 
                    "Set as folder-level permission in SharePoint" 
                }
            }
        }
        
        # Create detailed unique permissions list
        $detailedUniquePermissions = $allPermissions |
            Group-Object IdentityReference, FileSystemRights, AccessControlType |
            Select-Object @{
                Name='Identity'; Expression={($_.Name -split ', ')[0]}
            }, @{
                Name='FileSystemRights'; Expression={($_.Name -split ', ')[1]}
            }, @{
                Name='AccessControlType'; Expression={($_.Name -split ', ')[2]}
            }, @{
                Name='Occurrences'; Expression={$_.Count}
            }, @{
                Name='SamplePath'; Expression={$_.Group[0].Path}
            }, @{
                Name='IsInherited'; Expression={$_.Group[0].IsInherited}
            }, @{
                Name='InheritanceFlags'; Expression={$_.Group[0].InheritanceFlags}
            }, @{
                Name='PropagationFlags'; Expression={$_.Group[0].PropagationFlags}
            } |
            Sort-Object Identity, FileSystemRights
        
        # Create identity summary
        $identitySummary = $allPermissions | 
            Group-Object IdentityReference | 
            Select-Object @{
                Name='Identity'; Expression={$_.Name}
            }, @{
                Name='TotalOccurrences'; Expression={$_.Count}
            }, @{
                Name='UniqueRightsCombinations'; Expression={
                    ($_.Group | Select-Object FileSystemRights, AccessControlType -Unique).Count
                }
            } |
            Sort-Object TotalOccurrences -Descending
        
        # Export results
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $uniquePermissionsFile = Join-Path $OutputDirectory "$($Domain)_Unique_Permissions_Detailed_$timestamp.csv"
        $identitySummaryFile = Join-Path $OutputDirectory "$($Domain)_Unique_Identities_Summary_$timestamp.csv"
        $sharepointMappingFile = Join-Path $OutputDirectory "$($Domain)_SharePoint_Permission_Mapping_$timestamp.csv"
        $folderPermissionsFile = Join-Path $OutputDirectory "$($Domain)_Folder_Permissions_Analysis_$timestamp.csv"
        
        $detailedUniquePermissions | Export-Csv -Path $uniquePermissionsFile -NoTypeInformation -Encoding UTF8
        $identitySummary | Export-Csv -Path $identitySummaryFile -NoTypeInformation -Encoding UTF8
        $sharepointMapping | Export-Csv -Path $sharepointMappingFile -NoTypeInformation -Encoding UTF8
        $folderPermissions | Export-Csv -Path $folderPermissionsFile -NoTypeInformation -Encoding UTF8
        
        Write-Log "Exported unique permissions: $($detailedUniquePermissions.Count) combinations" -Level Success
        Write-Log "Exported identity summary: $($identitySummary.Count) identities" -Level Success
        Write-Log "Exported SharePoint mapping: $($sharepointMapping.Count) permission boundaries" -Level Success
        
        return @{
            UniquePermissions = $detailedUniquePermissions
            IdentitySummary = $identitySummary
            SharePointMapping = $sharepointMapping
            FolderPermissions = $folderPermissions
            UniquePermissionsFile = $uniquePermissionsFile
            IdentitySummaryFile = $identitySummaryFile
            SharePointMappingFile = $sharepointMappingFile
            FolderPermissionsFile = $folderPermissionsFile
        }
    }
    catch {
        Write-Log "Failed to extract unique permissions: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function New-FileShareAssessmentExcel {
    Write-Log "Creating Excel workbook..." -Level Info
    
    $excelFile = Join-Path $OutputDirectory "$($Domain)_File_Share_Assessment.xlsx"
    
    # Remove existing file if it exists
    if (Test-Path $excelFile) {
        try {
            Remove-Item $excelFile -Force -ErrorAction Stop
            Write-Log "Removed existing Excel file" -Level Info
        }
        catch {
            Write-Log "Could not remove existing Excel file. It may be open." -Level Warning
        }
    }
    
    try {
        # Get all fileaudit CSV files
        $fileSharesCSVFiles = Get-ChildItem -Path $OutputDirectory -Filter "fileaudit*.csv" -ErrorAction SilentlyContinue
        
        if ($fileSharesCSVFiles) {
            Write-Log "Processing $($fileSharesCSVFiles.Count) share audit files" -Level Info
            foreach ($file in $fileSharesCSVFiles) {
                $shareName = $file.BaseName.Replace("fileaudit_", "")
                $data = Import-Csv -Path $file.FullName
                $data | Export-Excel -Path $excelFile -WorksheetName $shareName -TableName $shareName -AutoSize -AutoFilter -FreezeTopRow
                Write-Log "  Added worksheet: $shareName" -Level Info
            }
        }
        
        # Get all unsupported filenames CSV files
        $unsupportedCSVFiles = Get-ChildItem -Path $OutputDirectory -Filter "unsupported_filenames*.csv" -ErrorAction SilentlyContinue
        
        if ($unsupportedCSVFiles) {
            Write-Log "Processing $($unsupportedCSVFiles.Count) unsupported filename reports" -Level Info
            foreach ($file in $unsupportedCSVFiles) {
                $shareName = $file.BaseName.Replace("unsupported_filenames_", "")
                $data = Import-Csv -Path $file.FullName
                
                if ($data) {
                    $worksheetName = "USC - $shareName"
                    $data | Export-Excel -Path $excelFile -WorksheetName $worksheetName -TableName $worksheetName -AutoSize -AutoFilter -FreezeTopRow
                    Write-Log "  Added worksheet: $worksheetName ($($data.Count) files)" -Level Info
                }
            }
        }
        
        # Get all permissions CSV files from RawData folder
        $rawDataPath = Join-Path $OutputDirectory "RawData"
        if (Test-Path $rawDataPath) {
            $permissionFiles = Get-ChildItem -Path $rawDataPath -Filter "permissions*.csv" -ErrorAction SilentlyContinue
            
            if ($permissionFiles) {
                Write-Log "Processing $($permissionFiles.Count) permission files" -Level Info
                foreach ($file in $permissionFiles) {
                    $worksheetName = $file.BaseName
                    if ($worksheetName.Length -gt 31) {
                        $worksheetName = $worksheetName.Substring(0, 31)
                    }
                    $data = Import-Csv -Path $file.FullName
                    if ($data) {
                        $data | Export-Excel -Path $excelFile -WorksheetName $worksheetName -TableName $worksheetName -AutoSize -AutoFilter -FreezeTopRow
                        Write-Log "  Added worksheet: $worksheetName" -Level Info
                    }
                }
            }
        }
        
        Write-Log "Excel workbook created successfully: $excelFile" -Level Success
        return $excelFile
    }
    catch {
        Write-Log "Failed to create Excel workbook: $($_.Exception.Message)" -Level Error
        return $null
    }
}

#endregion

#region Main Script

Clear-Host
Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "FILE SHARE ASSESSMENT TOOL" -ForegroundColor Cyan
Write-Host "Domain: $Domain" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Log "Assessment started" -Level Info
Write-Host ""

# Step 1: Check prerequisites
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 1: Checking Prerequisites" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

# Check long paths
if (-not (Test-LongPathsEnabled)) {
    Write-Log "Long paths are not enabled" -Level Warning
    $enable = Read-Host "Enable long paths now? (Y/N)"
    if ($enable -eq 'Y' -or $enable -eq 'y') {
        if (-not (Enable-LongPaths)) {
            Write-Log "Failed to enable long paths. Assessment may fail on deep folder structures." -Level Error
        }
    }
}
else {
    Write-Log "Long paths are enabled" -Level Success
}

# Ensure output directory exists
if (-not (Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
        Write-Log "Created output directory: $OutputDirectory" -Level Success
    }
    catch {
        Write-Log "Failed to create output directory: $($_.Exception.Message)" -Level Error
        exit 1
    }
}

# Install required modules
if (-not (Install-RequiredModules)) {
    Write-Log "Failed to install required modules. Cannot continue." -Level Error
    exit 1
}

# Step 2: Discover shares
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 2: Discovering SMB Shares" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

$shares = Get-LocalSMBShares

if ($shares.Count -eq 0) {
    Write-Log "No shares found to assess. Exiting." -Level Error
    exit 1
}

foreach ($share in $shares) {
    Write-Host "  - $($share.Name): $($share.Path)" -ForegroundColor White
}

# Step 3: Analyze shares
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 3: Analyzing Shares" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

foreach ($share in $shares) {
    Write-Host "`n--- Processing Share: $($share.Name) ---" -ForegroundColor Magenta
    
    # Size analysis
    Export-FileShareSize -Share $share
    
    # Permissions analysis
    if (-not $SkipPermissions) {
        Export-FileSharePermissions -Share $share
    }
    else {
        Write-Log "Permissions analysis skipped per parameter" -Level Info
    }
    
    # Unsupported filenames
    Export-FileShareUnsupportedFileNames -Share $share
    
    Write-Host ""
}

# Step 4: Extract unique permissions
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 4: Extracting Unique Permissions" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

if (-not $SkipPermissions) {
    $uniquePermissionsResult = Export-UniquePermissions
}
else {
    Write-Log "Unique permissions extraction skipped (no permissions data collected)" -Level Info
}

# Step 5: Create Excel report
Write-Host "`n$SubSeparator" -ForegroundColor Yellow
Write-Host "STEP 5: Creating Excel Report" -ForegroundColor Yellow
Write-Host $SubSeparator -ForegroundColor Yellow

$excelPath = New-FileShareAssessmentExcel

# Summary
$EndTime = Get-Date
$Duration = $EndTime - $StartTime

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "ASSESSMENT COMPLETE" -ForegroundColor Cyan
Write-Host $Separator -ForegroundColor Cyan
Write-Host "Duration: $($Duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "Shares Assessed: $($shares.Count)" -ForegroundColor White
Write-Host "Errors: $script:ErrorCount | Warnings: $script:WarningCount" -ForegroundColor $(if($script:ErrorCount -gt 0){'Red'}else{'Green'})

if ($uniquePermissionsResult) {
    Write-Host "`n📋 Unique Permissions Extracted:" -ForegroundColor Cyan
    Write-Host "  • $($uniquePermissionsResult.UniquePermissions.Count) unique permission combinations" -ForegroundColor White
    Write-Host "  • $($uniquePermissionsResult.IdentitySummary.Count) unique identities (users/groups)" -ForegroundColor White
    Write-Host "`n📍 SharePoint Migration Planning:" -ForegroundColor Yellow
    Write-Host "  • $($uniquePermissionsResult.FolderPermissions.Count) folder paths analyzed" -ForegroundColor White
    Write-Host "  • $($uniquePermissionsResult.SharePointMapping.Count) permission boundaries (lowest level)" -ForegroundColor Cyan
    Write-Host "    └─ Folders requiring unique SharePoint permissions" -ForegroundColor Gray
}

if ($excelPath -and (Test-Path $excelPath)) {
    Write-Host "`n📊 Excel Report: $excelPath" -ForegroundColor Green
    
    $openExcel = Read-Host "`nOpen Excel report now? (Y/N)"
    if ($openExcel -eq 'Y' -or $openExcel -eq 'y') {
        Start-Process $excelPath
    }
}
else {
    Write-Log "Excel report was not created successfully" -Level Error
}

Write-Host ""
Write-Log "Assessment completed" -Level Success

#endregion
