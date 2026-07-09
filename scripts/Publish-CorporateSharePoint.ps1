<#
.SYNOPSIS
Publishes the corporate Power-BI checkout to the ELCC SharePoint PBI folder.

.EXAMPLE
.\scripts\Publish-CorporateSharePoint.ps1 -DryRun

.EXAMPLE
.\scripts\Publish-CorporateSharePoint.ps1 -Mirror
#>

[CmdletBinding()]
param(
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

$sourceRoot = 'C:\GitHub\Power BI'
$destinationRoot = 'C:\Users\michael.louie\ESDC EDSC\Federal Secretariat on Early Learning and Child Care - Quants\PBI'
$publishScript = Join-Path $PSScriptRoot 'Publish-PowerBIToSharePoint.ps1'

if (-not (Test-Path -LiteralPath $publishScript -PathType Leaf)) {
    throw "Missing publish engine script: $publishScript"
}

$publishArgs = @{
    SourceRoot = $sourceRoot
    DestinationRoot = $destinationRoot
}

if ($Mirror) {
    $publishArgs.Mirror = $true
}

if ($DryRun) {
    $publishArgs.DryRun = $true
}

if ($IncludeDevelopmentTools) {
    $publishArgs.IncludeDevelopmentTools = $true
}

if ($ShowDetails) {
    $publishArgs.ShowDetails = $true
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
    $publishArgs.LogPath = $LogPath
}

& $publishScript @publishArgs

