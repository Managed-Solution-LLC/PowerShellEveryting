<#
.SYNOPSIS
    This script uses AzCopy to archive files to Azure Blob Storage with specified parameters.

.DESCRIPTION
    The script allows users to specify a folder, tier, Blob Storage SAS token, and the path to the files to be archived.
    It checks for the presence of azcopy.exe in the current directory and constructs the appropriate Blob Storage URL with SAS token.
    The script then uses AzCopy to copy files recursively to the specified Azure Blob Storage location.

.PARAMETER ParameterName
    [folder]
    Specifies the folder name to which files will be archived. Valid options are Billing, LA Office, NC Office.
    [tier]
    Specifies the storage tier for the blobs. Valid options are Hot, Cold, Archive.
    [blobSasToken]
    Specifies the SAS token for Azure Blob Storage, which provides access to the storage account.
    [path]
    Specifies the local path to the files that need to be archived.

.EXAMPLE
    .\AzCopyCommand.ps1 -folder "Billing" -tier "cold" -blobSasToken "https://mystorageaccount.blob.core.windows.net/archiveblob?sv=2020-08-04&ss=b&srt=sco&sp=rwdlacup&se=2021-12-31T23:59:59Z&st=2021-01-01T00:00:00Z&spr=https&sig=exampleSASToken" -path "C:\FilesToArchive"
    This example archives files from the specified path to the Billing folder in the Cold tier of Azure Blob Storage.

    $blobToken = "https://mystorageaccount.blob.core.windows.net/archiveblob?sv=2020-08-04&ss=b&srt=sco&sp=rwdlacup&se=2021-12-31T23:59:59Z&st=2021-01-01T00:00:00Z&spr=https&sig=exampleSASToken"
    $multipath = @("C:\FilesToArchive1", "C:\FilesToArchive2")
    foreach ($item in $multipath) {
        .\AzCopyCommand.ps1 -folder "Billing" -tier "cold" -blobSasToken $blobToken -path $item
    }
    This example demonstrates how to use the script in a loop to archive multiple paths to the Billing folder in the Cold tier of Azure Blob Storage.

    .\AzCopyCommand.ps1 "NC Office" "Archive" "https://mystorageaccount.blob.core.windows.net/archiveblob?sv=2020-08-04&ss=b&srt=sco&sp=rwdlacup&se=2021-12-31T23:59:59Z&st=2021-01-01T00:00:00Z&spr=https&sig=exampleSASToken" "C:\FilesToArchive"
# AzCopyCommand.ps1 - A PowerShell script to archive files to Azure Blob Storage using AzCopy
.NOTES
    Author: William Ford
    Date: 2025-06-25
    Version: 1.2
    Required Modules: AzCopy
    Output: Files are copied to Azure Blob Storage in the specified folder and tier.
    History:
        - 2025-06-25: Initial version with basic functionality.
        - 2025-06-27: Added error handling for AzCopy download and extraction, and improved folder handling.
#>
# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-27
# =============================
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0,HelpMessage="Enter the tier (Hot, Cool, Cold, Archive)")]
    [ValidateSet("Hot", "Cool", "Cold", "Archive")]
    [ValidateNotNullOrEmpty()]
    [string]$tier,

    [Parameter(Mandatory=$true,Position=1,HelpMessage="Enter the Blob Storage SAS Token")]
    [ValidateNotNullOrEmpty()]
    [string]$blobSasToken,

    [Parameter(Mandatory=$true,Position=2,HelpMessage="Enter the path to the files to be archived")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$path,

    [Parameter(Mandatory=$false,Position=3,HelpMessage="Enter the AzCopy executable path if not in current directory")]
    [string]$azCopyPath = ".\azcopy.exe",

    [Parameter(Mandatory=$true,Position=4,HelpMessage="Enter the folder name you wish to upload to.")]
    [string]$folder,

    [Parameter(Mandatory=$false,Position=5,HelpMessage="Enter the Root Folder for the archive, default is none")]
    [string]$rootFolder = "/"

)
#region AzCopy Downlaod
# Check if working directory has azcopy.exe, if not download it
if (-not (Test-Path -Path "$azCopyPath")) {
    Write-Host "Downloading AzCopy..."
    try {
        # Get the actual download URL
        $downloadUrl = (Invoke-WebRequest -Uri https://aka.ms/downloadazcopy-v10-windows -MaximumRedirection 0 -ErrorAction SilentlyContinue).headers.location
        
        if (-not $downloadUrl) {
            Write-Error "Failed to get AzCopy download URL"
            exit 1
        }
        
        # Download the ZIP file
        Invoke-WebRequest -Uri $downloadUrl -OutFile "azcopy.zip" -ErrorAction Stop
        
        # Verify the ZIP file was downloaded and is valid
        if (-not (Test-Path "azcopy.zip") -or (Get-Item "azcopy.zip").Length -eq 0) {
            Write-Error "Failed to download AzCopy ZIP file or file is empty"
            exit 1
        }
        
        # Extract the ZIP file
        Expand-Archive -Path '.\azcopy.zip' -DestinationPath '.\'
        Remove-Item -Path '.\azcopy.zip' -Force
        
        # Find the azcopy.exe file
        $azCopyPath = (Get-ChildItem -Path '.\' -Recurse -File -Filter 'azcopy.exe').FullName
        
        if (-not $azCopyPath -or -not (Test-Path $azCopyPath)) {
            Write-Error "AzCopy executable not found after extraction"
            exit 1
        }
        
        Write-Host "AzCopy downloaded and extracted successfully to: $azCopyPath"
    }
    catch {
        Write-Error "Failed to download or extract AzCopy: $($_.Exception.Message)"
        exit 1
    }
}
#endregion

#Build Folder
if ($folder -eq "" -and $rootFolder -eq "/") {
    Write-Host "Uploading to root folder, no subfolder specified."
} else {
    $folder = $rootFolder + $folder
    # Ensure the folder begins with a slash
    if ($folder -notlike "/*") {
        $folder = "/" + $folder
    }
}
# Check if the tier is valid
$tier = $tier.ToLower()
if ($tier -notin @("hot","cool", "cold", "archive")) {
    Write-Host "Invalid tier specified. Valid options are: Hot, Cool, Cold, Archive."
    exit 1
}

# Rebuild the Blob Storage URL with SAS token if folder is specified
if ($folder -ne "") {
    # Split after the archiveblob part
    $blobSasTokenParts = $blobSasToken -split '\?'
    $blobSasTokenBase = $blobSasTokenParts[0]
    $blobSasTokenQuery = $blobSasTokenParts[1]
    # Check if the folder has a space
    if ($folder -like "* *") {
        # Replace spaces with %20
        $folder = $folder -replace " ", "%20"
    }
    # Rebuild the SAS token without the folder
    $blobSasToken = $blobSasTokenBase + $folder + "?" + $blobSasTokenQuery
}

& $azCopyPath copy $path $blobSasToken --recursive=true --blob-type=BlockBlob --block-blob-tier=$tier
