<#!
    .SYNOPSIS
    Uninstall legacy (MSI) Office 2016 or lower, then install Microsoft 365 Apps (Business or Enterprise / ProPlus).

    .DESCRIPTION
    Detects and removes MSI-based Office 2016 / 2013 / 2010 (and related Visio/Project) prior to installing Microsoft 365 Apps
    using the Office Deployment Tool (ODT). Generates a configuration XML tailored to parameters, downloads/extracts the latest
    ODT if not supplied, and executes a silent install. Supports choosing Business (O365BusinessRetail) or Enterprise (O365ProPlusRetail)
    SKU, language, channel, architecture, shared computer licensing, and logging path.

    .NOTES
    Run elevated. Requires outbound HTTP(S) to download ODT unless -ODTPath provided. Uses only registry queries (no Win32_Product).

    .EXAMPLE
    .\Install-M365Apps.ps1 -Edition Enterprise -Channel Current -Language en-us -Verbose

    .EXAMPLE
    .\Install-M365Apps.ps1 -Edition Business -SharedComputerLicensing -Channel MonthlyEnterprise -LogPath C:\Logs

    .EXAMPLE
    .\Install-M365Apps.ps1 -Edition Enterprise -Architecture x86 -WhatIf
!>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet('Enterprise','Business')]
    [string]$Edition = 'Enterprise',

    [ValidateSet('Current','MonthlyEnterprise','SemiAnnualEnterprise','BetaChannel')]
    [string]$Channel = 'Current',

    [ValidateSet('Auto','x64','x86')]
    [string]$Architecture = 'Auto',

    [string]$Language = 'en-us',

    [switch]$SharedComputerLicensing,

    [string]$TempPath = (Join-Path -Path $env:TEMP -ChildPath 'M365Install'),

    [string]$LogPath = 'C:\ProgramData\Microsoft\OfficeDeploymentLogs',

    [string]$ODTDownloadUrl = 'https://download.microsoft.com/download/2/6/5/265CAFD5-AD32-4B00-B40E-3E25CC98D4A9/officedeploymenttool.exe',

    [string]$ODTPath,  # Optional pre-downloaded ODT folder containing setup.exe

    [switch]$Force,

    [switch]$ForceReboot,

    [switch]$SkipRemovalCheck  # If set, always attempt removal of MSI Office via RemoveMSI
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $current = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    return $current.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    Write-Error 'This script must be run elevated (Administrator).'
    exit 1
}

Write-Verbose "Edition: $Edition | Channel: $Channel | Arch: $Architecture | Lang: $Language"

# Map edition to Product ID
$productId = switch ($Edition) {
    'Enterprise' { 'O365ProPlusRetail' }
    'Business'   { 'O365BusinessRetail' }
}

function Get-OfficeMsiInstalls {
    <#
        Returns objects representing MSI-based legacy Office (<=2016) products: Office core, Visio, Project.
        Detection heuristic: DisplayName contains 'Microsoft Office' OR Visio/Project and does NOT contain '365' or 'Click-to-Run'.
        Uses registry uninstall keys only (no Win32_Product).
    #>
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $results = foreach ($p in $paths) {
        if (Test-Path $p) {
            Get-ChildItem $p | ForEach-Object {
                try {
                    $item = Get-ItemProperty $_.PsPath -ErrorAction Stop
                    $name = $item.DisplayName
                    if ([string]::IsNullOrWhiteSpace($name)) { continue }
                    if ($name -match 'Visio|Project|Microsoft Office') {
                        if ($name -notmatch '365' -and $name -notmatch 'Click-?to-?Run') {
                            # Filter for likely 2016 or lower (coarse: presence of 2016/2013/2010 OR absence of year but MSI stub)
                            if ($name -match '2016|2013|2010' -or $item.DisplayVersion -lt '17.0') {
                                [PSCustomObject]@{
                                    DisplayName    = $name
                                    DisplayVersion = $item.DisplayVersion
                                    UninstallString= $item.UninstallString
                                    PSPath         = $item.PSPath
                                }
                            }
                        }
                    }
                } catch {}
            }
        }
    }
    return $results | Sort-Object -Property DisplayName -Unique
}

function Ensure-Path {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        Write-Verbose "Creating directory: $Path"
        if ($PSCmdlet.ShouldProcess($Path,'Create Directory')) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
}

Ensure-Path -Path $TempPath
Ensure-Path -Path $LogPath

function Get-TargetArchitecture {
    param([string]$Requested)
    if ($Requested -and $Requested -ne 'Auto') { return $Requested }
    # Auto: Prefer x64 unless existing 32-bit Office found
    $clickToRunKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $clickToRunKey) {
        try { (Get-ItemProperty $clickToRunKey -Name Platform -ErrorAction Stop).Platform } catch { 'x64' }
    } else {
        # Check installed MSI (rare 32-bit on 64-bit OS) via uninstall strings containing x86
        $msi = Get-OfficeMsiInstalls
        if ($msi.UninstallString -match 'Program Files \(x86\)') { return 'x86' }
        return 'x64'
    }
}

$resolvedArch = Get-TargetArchitecture -Requested $Architecture
Write-Verbose "Resolved architecture: $resolvedArch"

function Download-ODT {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$DestinationFolder
    )
    Ensure-Path -Path $DestinationFolder
    $exe = Join-Path $DestinationFolder 'officedeploymenttool.exe'
    if (Test-Path (Join-Path $DestinationFolder 'setup.exe')) {
        Write-Verbose 'ODT already extracted.'
        return $DestinationFolder
    }
    Write-Verbose "Downloading ODT: $Url"
    if ($PSCmdlet.ShouldProcess($exe,'Download ODT')) {
        Invoke-WebRequest -Uri $Url -OutFile $exe -UseBasicParsing
        & $exe /quiet /extract:$DestinationFolder | Out-Null
        Remove-Item $exe -Force -ErrorAction SilentlyContinue
    }
    if (-not (Test-Path (Join-Path $DestinationFolder 'setup.exe'))) {
        throw 'Failed to obtain Office Deployment Tool (setup.exe not found).'
    }
    return $DestinationFolder
}

if (-not $ODTPath) {
    $ODTPath = Join-Path $TempPath 'ODT'
}

if (-not (Test-Path (Join-Path $ODTPath 'setup.exe'))) {
    $null = Download-ODT -Url $ODTDownloadUrl -DestinationFolder $ODTPath
}

$setupExe = Join-Path $ODTPath 'setup.exe'
if (-not (Test-Path $setupExe)) { throw 'setup.exe not found after download.' }

$legacyOffice = Get-OfficeMsiInstalls
if ($legacyOffice -and -not $SkipRemovalCheck) {
    Write-Host 'Detected legacy Office MSI products:' -ForegroundColor Cyan
    $legacyOffice | Format-Table DisplayName, DisplayVersion -AutoSize | Out-String | Write-Host
} elseif (-not $legacyOffice) {
    Write-Verbose 'No legacy MSI Office products detected.'
}

if (-not $legacyOffice -and -not $SkipRemovalCheck -and -not $Force) {
    Write-Verbose 'Skipping removal (none found). Use -SkipRemovalCheck to force RemoveMSI anyway.'
}

function New-ConfigXmlContent {
    param(
        [string]$ProductId,
        [string]$Channel,
        [string]$Language,
        [string]$Arch,
        [switch]$SharedComputerLicensing,
        [switch]$AlwaysRemoveMsi
    )
    $channelAttr = $Channel
    $props = @()
    if ($SharedComputerLicensing) { $props += '<Property Name="SharedComputerLicensing" Value="1" />' }
    $props += '<Property Name="AUTOACTIVATE" Value="1" />'

    $removeMsiBlock = if ($AlwaysRemoveMsi -or $legacyOffice) { '<RemoveMSI />' } else { '' }

    @"
<Configuration>
  <Add OfficeClientEdition="$Arch" Channel="$channelAttr">
    <Product ID="$ProductId">
      <Language ID="$Language" />
    </Product>
  </Add>
  $removeMsiBlock
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  $($props -join "`n  ")
  <Logging Level="Standard" Path="$LogPath" />
  <Updates Enabled="TRUE" />
</Configuration>
"@
}

$configXmlPath = Join-Path $TempPath 'M365Apps-Install.xml'
$xmlContent = New-ConfigXmlContent -ProductId $productId -Channel $Channel -Language $Language -Arch $resolvedArch -SharedComputerLicensing:$SharedComputerLicensing -AlwaysRemoveMsi:($SkipRemovalCheck)

if ($PSCmdlet.ShouldProcess($configXmlPath,'Write configuration XML')) {
    $xmlContent | Out-File -FilePath $configXmlPath -Encoding UTF8 -Force
}
Write-Verbose "Generated configuration XML at $configXmlPath"

function Invoke-OfficeDeployment {
    param(
        [Parameter(Mandatory)][string]$SetupExe,
        [Parameter(Mandatory)][string]$ConfigPath
    )
    $args = "/configure `"$ConfigPath`""
    Write-Host "Running: $SetupExe $args" -ForegroundColor Green
    if ($PSCmdlet.ShouldProcess('Microsoft 365 Apps','Install/Configure')) {
        $process = Start-Process -FilePath $SetupExe -ArgumentList $args -PassThru -Wait -WindowStyle Hidden
        if ($process.ExitCode -ne 0) { throw "Office Deployment Tool returned exit code $($process.ExitCode)" }
    }
}

Invoke-OfficeDeployment -SetupExe $setupExe -ConfigPath $configXmlPath

Write-Host 'Microsoft 365 Apps deployment complete.' -ForegroundColor Green

if ($ForceReboot) {
    Write-Host 'Rebooting system (ForceReboot specified)...' -ForegroundColor Yellow
    Restart-Computer -Force
}

<#
    EXIT CODES / RESULT
    0   Success (script completed without throwing)
    >0  Exception thrown (non-terminating exceptions promoted due to $ErrorActionPreference)
*#>

exit 0
