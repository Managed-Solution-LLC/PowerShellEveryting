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
    Version: 1.0
    Required Modules: AzCopy
    Output: Files are copied to Azure Blob Storage in the specified folder and tier.
#>
# =============================
# VALIDATED FOR PUBLIC RELEASE
# Date: 2025-06-25
# =============================
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0,HelpMessage="Enter the folder name you wish to upload to.")]
    [string]$folder,

    [Parameter(Mandatory=$true,Position=1,HelpMessage="Enter the tier (Hot, Cool, Cold, Archive)")]
    [string]$tier,

    [Parameter(Mandatory=$true,Position=2,HelpMessage="Enter the Blob Storage SAS Token")]
    [string]$blobSasToken,

    [Parameter(Mandatory=$true,Position=3,HelpMessage="Enter the path to the files to be archived")]
    [string]$path,

    [Parameter(Mandatory=$false,Position=4,HelpMessage="Enter the AzCopy executable path if not in current directory")]
    [string]$azCopyPath = ".\azcopy.exe",

    [Parameter(Mandatory=$false,Position=5,HelpMessage="Enter the Root Folder for the archive, default is /Archive/")]
    [string]$rootFolder = "/Archive/"

)
# Check if working directory has azcopy.exe, if not download it
if (-not (Test-Path -Path "$azCopyPath")) {
    Write-Host "Downloading AzCopy..."
    (Invoke-WebRequest -Uri https://aka.ms/downloadazcopy-v10-windows -MaximumRedirection 0 -ErrorAction SilentlyContinue).headers.location | Out-File -FilePath "azcopy.zip" -Encoding ASCII
    Expand-archive -Path '.\azcopyv10.zip' -Destinationpath '.\'
    Remove-Item -Path '.\azcopyv10.zip' -Force
    $azCopyPath = (Get-ChildItem -path '.\' -Recurse -File -Filter 'azcopy.exe').FullName
    exit 1
}

#Build Folder
$folder = $rootFolder + $folder
# Check if the tier is valid
$tier = $tier.ToLower()
if ($tier -notin @("hot","cool", "cold", "archive")) {
    Write-Host "Invalid tier specified. Valid options are: Hot, Cool, Cold, Archive."
    exit 1
}
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
