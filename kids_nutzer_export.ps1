<#
.SYNOPSIS
Exports Microsoft 365 users to a CSV matching the column layout of test_daten.csv.

.DESCRIPTION
Connects to Microsoft Graph and Exchange Online, retrieves user profile data, mailbox forwarding settings,
and group mail addresses, then writes the results to a CSV with the columns:
Name,Surname,UserPrincipalName,ForwardTo,TargetGroupMailAddress,StreetAddress,PostalCode,City,MobilePhone.

.PARAMETER CsvPath
Path to the CSV file that will be created.

.PARAMETER UserFilter
Optional OData filter for Graph (example: "accountEnabled eq true").

.PARAMETER TargetGroupMailRegex
Optional regex to restrict which group mail addresses are exported.

.PARAMETER IncludeAllGroupMails
If set, TargetGroupMailAddress contains all matching group mails joined by ';'.
Otherwise, only the first matching group mail is used (with a warning if multiple match).

.EXAMPLE
PS> .\kids_nutzer_export.ps1 -CsvPath .\export.csv

Exports all users to export.csv with the required columns.
#>
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Groups,ExchangeOnlineManagement
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,
    [string]$UserFilter,
    [string]$TargetGroupMailRegex,
    [switch]$IncludeAllGroupMails
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

function Get-SmtpAddressString {
    param (
        $Value
    )

    if (-not $Value) {
        return ''
    }

    $text = $Value.ToString()
    if ($text -match '^smtp:(.+)$') {
        return $Matches[1]
    }
    if ($text -match '^SMTP:(.+)$') {
        return $Matches[1]
    }

    return $text
}

function Resolve-GroupMail {
    param (
        [Parameter(Mandatory)]
        [string]$GroupId,
        [Parameter(Mandatory)]
        $GroupObject,
        [Parameter(Mandatory)]
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey($GroupId)) {
        return $Cache[$GroupId]
    }

    $mail = $null
    if ($GroupObject.PSObject.Properties['Mail'] -and $GroupObject.Mail) {
        $mail = $GroupObject.Mail
    } elseif ($GroupObject.PSObject.Properties['AdditionalProperties'] -and $GroupObject.AdditionalProperties.ContainsKey('mail')) {
        $mail = $GroupObject.AdditionalProperties['mail']
    }

    if (-not $mail) {
        $groupDetails = Get-MgGroup -GroupId $GroupId -Property 'mail' -ErrorAction Stop
        $mail = $groupDetails.Mail
    }

    $Cache[$GroupId] = $mail
    return $mail
}

function Get-ODataType {
    param (
        [Parameter(Mandatory)]
        $Entry
    )

    $property = $Entry.PSObject.Properties['@odata.type']
    if ($property) {
        return $property.Value
    }

    $property = $Entry.PSObject.Properties['OdataType']
    if ($property) {
        return $property.Value
    }

    if ($Entry.PSObject.Properties['AdditionalProperties']) {
        $additional = $Entry.AdditionalProperties
        if ($additional -and $additional.ContainsKey('@odata.type')) {
            return $additional['@odata.type']
        }
    }

    return $null
}

function Test-IsGroupEntry {
    param (
        [Parameter(Mandatory)]
        $Entry
    )

    $odataType = Get-ODataType -Entry $Entry
    if ($odataType -eq '#microsoft.graph.group') {
        return $true
    }

    $typeNames = $Entry.PSObject.TypeNames
    if ($typeNames -match 'MicrosoftGraphGroup' -or $typeNames -match 'IMicrosoftGraphGroup') {
        return $true
    }

    return $false
}

$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups'
)

foreach ($module in $graphModules) {
    Ensure-Module -Name $module
    Import-Module $module -ErrorAction Stop
}

Ensure-Module -Name 'ExchangeOnlineManagement'
Import-Module ExchangeOnlineManagement -ErrorAction Stop

if (Get-Command -Name 'Select-MgProfile' -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name 'beta'
}

Write-Verbose 'Connecting to Microsoft Graph...'
Connect-MgGraph -Scopes 'User.Read.All', 'Group.Read.All'

Write-Verbose 'Connecting to Exchange Online...'
Connect-ExchangeOnline

try {
    Write-Verbose 'Loading mailboxes for forwarding settings...'
    $mailboxForwarding = @{}
    $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
    foreach ($mailbox in $mailboxes) {
        if (-not $mailbox.UserPrincipalName) {
            continue
        }

        $mailboxForwarding[$mailbox.UserPrincipalName.ToLowerInvariant()] = Get-SmtpAddressString -Value $mailbox.ForwardingSmtpAddress
    }

    $userProperties = @(
        'id',
        'givenName',
        'surname',
        'userPrincipalName',
        'streetAddress',
        'postalCode',
        'city',
        'mobilePhone'
    )

    Write-Verbose 'Loading users from Microsoft Graph...'
    if ($UserFilter) {
        $users = Get-MgUser -All -Filter $UserFilter -Property $userProperties -ErrorAction Stop
    } else {
        $users = Get-MgUser -All -Property $userProperties -ErrorAction Stop
    }

    if (-not $users) {
        throw 'No users returned from Microsoft Graph.'
    }

    $groupMailCache = @{}
    $exportRows = foreach ($user in $users) {
        $userPrincipalName = $user.UserPrincipalName
        $forwardTo = ''
        if ($userPrincipalName -and $mailboxForwarding.ContainsKey($userPrincipalName.ToLowerInvariant())) {
            $forwardTo = $mailboxForwarding[$userPrincipalName.ToLowerInvariant()]
        }

        $groupMails = @()
        try {
            $memberOf = Get-MgUserMemberOf -UserId $user.Id -All -Property 'id,mail,displayName' -ErrorAction Stop
            foreach ($entry in $memberOf) {
                if (-not (Test-IsGroupEntry -Entry $entry)) {
                    continue
                }

                $mail = Resolve-GroupMail -GroupId $entry.Id -GroupObject $entry -Cache $groupMailCache
                if (-not $mail) {
                    continue
                }
                if ($TargetGroupMailRegex -and ($mail -notmatch $TargetGroupMailRegex)) {
                    continue
                }

                $groupMails += $mail
            }
        } catch {
            Write-Warning "Unable to load group membership for '$userPrincipalName': $($_.Exception.Message)"
        }

        $groupMails = @($groupMails | Sort-Object -Unique)
        $targetGroupMailAddress = ''
        if ($groupMails.Count -gt 0) {
            if ($IncludeAllGroupMails) {
                $targetGroupMailAddress = $groupMails -join ';'
            } else {
                if ($groupMails.Count -gt 1) {
                    Write-Warning "User '$userPrincipalName' is in multiple matching groups: $($groupMails -join ', '). Using the first match."
                }
                $targetGroupMailAddress = $groupMails[0]
            }
        }

        [PSCustomObject]@{
            Name                   = $user.GivenName
            Surname                = $user.Surname
            UserPrincipalName      = $userPrincipalName
            ForwardTo              = $forwardTo
            TargetGroupMailAddress = $targetGroupMailAddress
            StreetAddress          = $user.StreetAddress
            PostalCode             = $user.PostalCode
            City                   = $user.City
            MobilePhone            = $user.MobilePhone
        }
    }

    $exportRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $(@($exportRows).Count) users to '$CsvPath'."
}
finally {
    Write-Verbose 'Disconnecting from Exchange Online.'
    Disconnect-ExchangeOnline -Confirm:$false
    if (Get-Command -Name 'Disconnect-MgGraph' -ErrorAction SilentlyContinue) {
        Write-Verbose 'Disconnecting from Microsoft Graph.'
        Disconnect-MgGraph | Out-Null
    }
}
