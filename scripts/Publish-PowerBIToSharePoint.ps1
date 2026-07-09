<#
.SYNOPSIS
Publishes this Power BI repo to a locally synced SharePoint folder.

.DESCRIPTION
Copies the dashboard contents to a SharePoint/OneDrive sync folder that has the
same folder layout. Git metadata, local Power BI cache/settings files, and
development tooling folders are excluded. By default, files are copied and
updated but extra files in the destination are left alone. Use -Mirror to remove
destination files that no longer exist in the repo.

.EXAMPLE
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -DryRun

.EXAMPLE
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -Mirror

.EXAMPLE
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -DryRun -ShowDetails
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationRoot,

    [Parameter()]
    [string]$SourceRoot,

    [Parameter()]
    [switch]$Mirror,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [switch]$IncludeDevelopmentTools,

    [Parameter()]
    [switch]$ShowDetails,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
    $SourceRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $tempPath = if ([string]::IsNullOrWhiteSpace($env:TEMP)) { [System.IO.Path]::GetTempPath() } else { $env:TEMP }
    $LogPath = Join-Path $tempPath ("PowerBI-SharePoint-Publish-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
}

function Resolve-FullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $executionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Test-IsSubPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChildPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    $child = [System.IO.Path]::GetFullPath($ChildPath).TrimEnd('\')
    $parent = [System.IO.Path]::GetFullPath($ParentPath).TrimEnd('\')

    return $child.Equals($parent, [System.StringComparison]::OrdinalIgnoreCase) -or
        $child.StartsWith($parent + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

$source = Resolve-FullPath $SourceRoot
$destination = Resolve-FullPath $DestinationRoot

if (-not (Test-Path -LiteralPath $source -PathType Container)) {
    throw "SourceRoot does not exist or is not a folder: $source"
}

if (Test-IsSubPath -ChildPath $destination -ParentPath $source) {
    throw "DestinationRoot cannot be inside SourceRoot. Source: $source Destination: $destination"
}

if (-not (Test-Path -LiteralPath $destination -PathType Container)) {
    if ($DryRun) {
        Write-Host "Dry run: destination folder would be created: $destination"
    }
    else {
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
    }
}

$copyMode = if ($Mirror) { '/MIR' } else { '/E' }
$dryRunMode = if ($DryRun) { '/L' } else { $null }

$excludeDirectories = @(
    '.git',
    '.agents',
    '.codex',
    '.vs',
    '.vscode',
    'node_modules'
)

if (-not $IncludeDevelopmentTools) {
    $excludeDirectories += @(
        'scripts',
        'skills-for-fabric'
    )
}

$excludeFiles = @(
    'cache.abf',
    'localSettings.json',
    'Thumbs.db',
    'desktop.ini',
    '.DS_Store'
)

$robocopyArgs = @(
    $source,
    $destination,
    $copyMode,
    '/DCOPY:DAT',
    '/COPY:DAT',
    '/R:2',
    '/W:5',
    '/FFT',
    "/LOG:$LogPath",
    '/XD'
) + $excludeDirectories + @('/XF') + $excludeFiles

if ($ShowDetails) {
    $robocopyArgs += '/TEE'
}

if ($dryRunMode) {
    $robocopyArgs += $dryRunMode
}

Write-Host "Publishing Power BI repo"
Write-Host "Source:      $source"
Write-Host "Destination: $destination"
Write-Host "Mode:        $(if ($Mirror) { 'Mirror destination, including deletes' } else { 'Copy/update only' })"
Write-Host "Dry run:     $([bool]$DryRun)"
Write-Host "Log:         $LogPath"
Write-Host ''

& robocopy @robocopyArgs
$exitCode = $LASTEXITCODE

if ($exitCode -ge 8) {
    throw "Robocopy failed with exit code $exitCode. Review the log: $LogPath"
}

Write-Host ''
Write-Host "Publish completed successfully. Robocopy exit code: $exitCode"
