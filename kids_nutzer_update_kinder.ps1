<#
.SYNOPSIS
Updates Microsoft 365 users from a CSV by writing the Kinder column into an on-premises extension attribute.

.DESCRIPTION
Reads a CSV with columns UserPrincipalName and Kinder (children names separated by ";").
For each row, sets OnPremisesExtensionAttributes.ExtensionAttributeN to the formatted value.
Example: "Jesper; Brea" becomes "Jesper & Brea" when using the default JoinSeparator.

.PARAMETER CsvPath
Path to the CSV file with columns UserPrincipalName and Kinder.

.PARAMETER ExtensionAttribute
Number of the extension attribute to update (1-15). Defaults to 1.

.PARAMETER ListSeparator
Separator used in the Kinder column. Defaults to ';'.

.PARAMETER JoinSeparator
Separator used when joining multiple names for the extension attribute value. Defaults to ' & '.

.PARAMETER ClearEmpty
If set, clears the extension attribute when Kinder is empty. Otherwise empty Kinder rows are skipped.

.EXAMPLE
PS> .\kids_nutzer_update_kinder.ps1 -CsvPath .\export-clean.csv

.EXAMPLE
PS> .\kids_nutzer_update_kinder.ps1 -CsvPath .\export-clean.csv -WhatIf

.EXAMPLE
PS> .\kids_nutzer_update_kinder.ps1 -CsvPath .\export-clean.csv -ClearEmpty
#>
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]$CsvPath,
    [ValidateRange(1, 15)]
    [int]$ExtensionAttribute = 1,
    [ValidateNotNullOrEmpty()]
    [string]$ListSeparator = ';',
    [ValidateNotNullOrEmpty()]
    [string]$JoinSeparator = ' & ',
    [switch]$ClearEmpty
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param (
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "PowerShell module '$Name' is not installed. Install it before running this script."
    }
}

function Get-TrimmedValue {
    param (
        [Parameter(Mandatory)]
        $Record,
        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $property = $Record.PSObject.Properties[$PropertyName]
    if (-not $property) {
        return ''
    }

    $value = $property.Value
    if ($null -eq $value) {
        return ''
    }

    return $value.ToString().Trim()
}

function Convert-KinderToAttributeValue {
    param (
        [AllowEmptyString()]
        [string]$Kinder,
        [Parameter(Mandatory)]
        [string]$ListSeparator,
        [Parameter(Mandatory)]
        [string]$JoinSeparator
    )

    if (-not $Kinder) {
        return $null
    }

    $pattern = "\s*" + [regex]::Escape($ListSeparator) + "\s*"
    $parts = $Kinder -split $pattern
    $trimmed = @()
    foreach ($part in $parts) {
        $name = $part.Trim()
        if ($name) {
            $trimmed += $name
        }
    }

    if ($trimmed.Count -eq 0) {
        return $null
    }

    return ($trimmed -join $JoinSeparator)
}

$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users'
)

foreach ($module in $graphModules) {
    Ensure-Module -Name $module
    Import-Module $module -ErrorAction Stop
}

if (Get-Command -Name 'Select-MgProfile' -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name 'beta'
}

$attributeName = "ExtensionAttribute$ExtensionAttribute"

Write-Verbose 'Connecting to Microsoft Graph...'
Connect-MgGraph -Scopes 'User.ReadWrite.All'

$records = Import-Csv -LiteralPath $CsvPath
if (-not $records) {
    throw "No records found in CSV: $CsvPath"
}

$updated = 0
$skipped = 0
$failed = 0

foreach ($record in $records) {
    $upn = Get-TrimmedValue -Record $record -PropertyName 'UserPrincipalName'
    if (-not $upn) {
        Write-Warning 'Skipping row with missing UserPrincipalName.'
        $skipped++
        continue
    }

    $kinder = Get-TrimmedValue -Record $record -PropertyName 'Kinder'
    $value = Convert-KinderToAttributeValue -Kinder $kinder -ListSeparator $ListSeparator -JoinSeparator $JoinSeparator

    if ($null -eq $value) {
        if (-not $ClearEmpty) {
            Write-Verbose "Skipping $upn because Kinder is empty."
            $skipped++
            continue
        }
    }

    $payload = @{}
    $payload[$attributeName] = $value

    $action = "Update $attributeName to '$value'"
    if ($PSCmdlet.ShouldProcess($upn, $action)) {
        try {
            Update-MgUser -UserId $upn -OnPremisesExtensionAttributes $payload -ErrorAction Stop
            $updated++
        } catch {
            Write-Warning "Failed to update ${upn}: $($_.Exception.Message)"
            $failed++
        }
    }
}

Write-Host "Done. Updated: $updated. Skipped: $skipped. Failed: $failed."
