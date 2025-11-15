<#
.SYNOPSIS
Creates Microsoft 365 users from a CSV file and configures licensing, password reset contact, and email forwarding.

.DESCRIPTION
Reads user details (name, surname, email, alternate email) from a CSV file, creates the accounts in Microsoft 365 by using the Microsoft Graph
PowerShell SDK, assigns the Microsoft 365 Business Essentials (O365_BUSINESS_ESSENTIALS) license, registers the alternate email address for
self-service password reset, and enables email forwarding to the alternate address while retaining a copy in the mailbox.

The script requires the Microsoft Graph modules (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`,
`Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Identity.SignIns`) and the `ExchangeOnlineManagement` module.

.PARAMETER CsvPath
Path to the CSV file containing user entries. Required columns are Name, Surname, UserPrincipalName, and TargetGroupMailAddress. Optional columns include ForwardTo, StreetAddress, PostalCode, City, and MobilePhone.

.EXAMPLE
PS> .\kids_nutzer_erstellen.ps1 -CsvPath .\users.csv

Connects to Microsoft Graph and Exchange Online, creates or updates the listed users, assigns licenses, configures password reset contact,
and sets email forwarding to the alternative address.
#>
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Identity.SignIns,ExchangeOnlineManagement
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]$CsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$usageLocation = 'DE'
$licenseSkuPartNumber = 'O365_BUSINESS_ESSENTIALS'


function Test-EmailAddress {
    param (
        [Parameter(Mandatory)]
        [string]$Address
    )

    try {
        $null = [System.Net.Mail.MailAddress]::new($Address)
        return $true
    } catch {
        return $false
    }
}

function Resolve-LicenseSkuId {
    param (
        [Parameter(Mandatory)]
        [string]$SkuPartNumber
    )

    $subscribedSkus = Get-MgSubscribedSku -All
    $match = $subscribedSkus | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }

    if (-not $match) {
        throw "Unable to find license SKU '$SkuPartNumber'. Use Get-MgSubscribedSku to confirm availability."
    }

    $enabledUnits = $match.PrepaidUnits.Enabled
    $consumedUnits = $match.ConsumedUnits

    if ($enabledUnits -le $consumedUnits) {
        throw "No available licenses for SKU '$($match.SkuId)'."
    }

    return $match.SkuId
}

function Ensure-Module {
    param (
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "PowerShell module '$Name' is not installed. Install it before running this script."
    }
}

function New-TemporaryPassword {
    param (
        [int]$Length = 16
    )

    if ($Length -lt 8) {
        throw 'Password length must be at least 8 characters.'
    }

    $charSets = @{
        Upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
        Lower   = 'abcdefghijkmnopqrstuvwxyz'
        Digit   = '23456789'
        Special = '!@$%&*?'
    }

    $allCharacters = ($charSets.Values -join '')
    $passwordChars = @()

    foreach ($set in $charSets.Values) {
        $passwordChars += $set[(Get-Random -Minimum 0 -Maximum $set.Length)]
    }

    for ($i = $passwordChars.Count; $i -lt $Length; $i++) {
        $passwordChars += $allCharacters[(Get-Random -Minimum 0 -Maximum $allCharacters.Length)]
    }

    $shuffled = $passwordChars | Sort-Object { Get-Random }
    return -join $shuffled
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

function Get-GraphUserByUserPrincipalName {
    param (
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $escapedUpn = $UserPrincipalName.Replace("'", "''")

    try {
        $result = Get-MgUser `
            -Filter "userPrincipalName eq '$escapedUpn'" `
            -ConsistencyLevel eventual `
            -Top 1 `
            -ErrorAction Stop
    } catch {
        throw
    }

    if (-not $result) {
        return $null
    }

    if ($result -is [System.Array]) {
        return $result[0]
    }

    return $result
}

function Get-MailNicknameFromUpn {
    param (
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $nickname = $UserPrincipalName.Split('@')[0]
    $nickname = $nickname -replace '[^a-zA-Z0-9]', '_'

    if (-not $nickname) {
        $nickname = "user$([Guid]::NewGuid().ToString('N').Substring(0, 6))"
    }

    return $nickname.ToLowerInvariant()
}

function Get-GraphGroupByMailAddress {
    param (
        [Parameter(Mandatory)]
        [string]$MailAddress
    )

    $escapedMail = $MailAddress.Replace("'", "''")

    try {
        $result = Get-MgGroup `
            -Filter "mail eq '$escapedMail'" `
            -ConsistencyLevel eventual `
            -Top 2 `
            -ErrorAction Stop
    } catch {
        throw
    }

    if (-not $result) {
        throw "Unable to find group with mail address '$MailAddress'."
    }

    if ($result -is [System.Array]) {
        if ($result.Count -gt 1) {
            throw "Multiple groups found with mail address '$MailAddress'. Use the group's object Id to disambiguate."
        }

        return $result[0]
    }

    return $result
}

function Get-GroupFromCache {
    param (
        [Parameter(Mandatory)]
        [hashtable]$Cache,
        [Parameter(Mandatory)]
        [string]$MailAddress
    )

    $key = $MailAddress.ToLowerInvariant()

    if (-not $Cache.ContainsKey($key)) {
        $Cache[$key] = Get-GraphGroupByMailAddress -MailAddress $MailAddress
    }

    return $Cache[$key]
}

function Add-UserToGroup {
    param (
        [Parameter(Mandatory)]
        [string]$GroupId,
        [Parameter(Mandatory)]
        [string]$UserId,
        [string]$GroupName
    )

    try {
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $UserId -ErrorAction Stop | Out-Null
    } catch {
        $message = $_.Exception.Message

        if ($message -and ($message -match 'added object references already exist' -or $message -match 'already exist')) {
            Write-Verbose "User '$UserId' is already a member of group '$GroupName'."
            return
        }

        Write-Warning "Unable to add user '$UserId' to group '$GroupName': $message"
    }
}

function Set-AlternateEmailMethod {
    param (
        [Parameter(Mandatory)]
        [string]$UserId,
        [Parameter(Mandatory)]
        [string]$AlternateEmail
    )

    $newEmailCommand = Get-Command -Name 'New-MgUserAuthenticationEmailMethod' -ErrorAction SilentlyContinue
    $getEmailCommand = Get-Command -Name 'Get-MgUserAuthenticationEmailMethod' -ErrorAction SilentlyContinue
    $removeEmailCommand = Get-Command -Name 'Remove-MgUserAuthenticationEmailMethod' -ErrorAction SilentlyContinue

    if (-not $newEmailCommand -or -not $getEmailCommand -or -not $removeEmailCommand) {
        Write-Warning 'Graph authentication email methods cmdlets not available; skipping SSPR alternate email registration.'
        return
    }

    try {
        $existingMethods = Get-MgUserAuthenticationEmailMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($existingMethods) {
            $matching = $existingMethods | Where-Object { $_.EmailAddress -eq $AlternateEmail }
            if ($matching) {
                return
            }

            foreach ($method in $existingMethods) {
                if ($method.EmailAddress) {
                    Remove-MgUserAuthenticationEmailMethod -UserId $UserId -EmailAuthenticationMethodId $method.Id -ErrorAction Stop
                }
            }
        }

        New-MgUserAuthenticationEmailMethod -UserId $UserId -EmailAddress $AlternateEmail -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Unable to register alternate email for '$UserId': $($_.Exception.Message)"
    }
}

function Wait-ForMailbox {
    param (
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,
        [ValidateRange(30, 3600)]
        [int]$TimeoutSeconds = 600,
        [ValidateRange(5, 120)]
        [int]$CheckIntervalSeconds = 15
    )

    Write-Verbose "Waiting for mailbox '$UserPrincipalName' to be provisioned (timeout ${TimeoutSeconds}s)..."

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $null = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            Write-Verbose "Mailbox '$UserPrincipalName' is now available."
            return $true
        } catch {
            $message = $_.Exception.Message
            $shouldRetry = $false

            if ($_.FullyQualifiedErrorId -and $_.FullyQualifiedErrorId -match 'ManagementObjectNotFoundException') {
                $shouldRetry = $true
            } elseif ($message) {
                $lower = $message.ToLowerInvariant()
                if ($lower -match 'not.*found' -or $lower -match 'nicht.*gefunden') {
                    $shouldRetry = $true
                }
            }

            if (-not $shouldRetry) {
                throw
            }
        }

        Write-Verbose "Mailbox '$UserPrincipalName' not ready yet; retrying in $CheckIntervalSeconds seconds..."
        Start-Sleep -Seconds $CheckIntervalSeconds
    }

    throw "Mailbox for '$UserPrincipalName' was not provisioned within $TimeoutSeconds seconds."
}

$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Identity.SignIns'
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
Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Directory.ReadWrite.All', 'UserAuthenticationMethod.ReadWrite.All'

Write-Verbose 'Connecting to Exchange Online...'
Connect-ExchangeOnline

try {
    $licenseSkuId = Resolve-LicenseSkuId -SkuPartNumber $licenseSkuPartNumber
    Write-Verbose "Using license SKU '$licenseSkuId'."

    $groupCache = @{}
    Write-Verbose 'Target group mail addresses will be read from the CSV rows.'

    $rows = Import-Csv -LiteralPath $CsvPath
    if (-not $rows) {
        throw "CSV '$CsvPath' does not contain any rows."
    }

    $requiredColumns = 'Name', 'Surname', 'UserPrincipalName', 'TargetGroupMailAddress'
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $rows[0].PsObject.Properties.Name }
    if ($missingColumns) {
        throw "CSV file is missing required column(s): $($missingColumns -join ', ')."
    }

    $processed = @()

    foreach ($row in $rows) {
        $firstName = Get-TrimmedValue -Record $row -PropertyName 'Name'
        $lastName = Get-TrimmedValue -Record $row -PropertyName 'Surname'
        $userPrincipalName = (Get-TrimmedValue -Record $row -PropertyName 'UserPrincipalName').ToLowerInvariant()
        $alternateEmail = Get-TrimmedValue -Record $row -PropertyName 'ForwardTo'
        $streetAddress = Get-TrimmedValue -Record $row -PropertyName 'StreetAddress'
        $postalCode = Get-TrimmedValue -Record $row -PropertyName 'PostalCode'
        $city = Get-TrimmedValue -Record $row -PropertyName 'City'
        $mobilePhone = Get-TrimmedValue -Record $row -PropertyName 'MobilePhone'
        $rowTargetGroupMailAddress = Get-TrimmedValue -Record $row -PropertyName 'TargetGroupMailAddress'

        if (-not $firstName -or -not $lastName -or -not $userPrincipalName) {
            Write-Warning "Skipping row with missing data: $($row | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)"
            continue
        }

        if (-not (Test-EmailAddress -Address $userPrincipalName)) {
            Write-Warning "Skipping row because the email address '$userPrincipalName' is not valid."
            continue
        }

        if ($alternateEmail -and -not (Test-EmailAddress -Address $alternateEmail)) {
            Write-Warning "Skipping row because the alternate email address '$alternateEmail' is not valid."
            continue
        }

        $displayName = "$firstName $lastName"
        $existingUser = Get-GraphUserByUserPrincipalName -UserPrincipalName $userPrincipalName

        if ($existingUser) {
            Write-Error "User '$userPrincipalName' already exists. Skipping."
            $processed += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                Action            = 'Skipped (already exists)'
                Password          = $null
            }
            continue
        }

        if (-not $rowTargetGroupMailAddress) {
            Write-Warning "Skipping '$userPrincipalName' because TargetGroupMailAddress is missing or empty."
            $processed += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                Action            = 'Skipped (missing target group)'
                Password          = $null
            }
            continue
        }

        try {
            $targetGroup = Get-GroupFromCache -Cache $groupCache -MailAddress $rowTargetGroupMailAddress
        } catch {
            Write-Error "Unable to resolve target group '$rowTargetGroupMailAddress' for '$userPrincipalName': $($_.Exception.Message)"
            $processed += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                Action            = 'Skipped (group lookup failed)'
                Password          = $null
            }
            continue
        }

        $mailNickname = Get-MailNicknameFromUpn -UserPrincipalName $userPrincipalName
        $temporaryPassword = New-TemporaryPassword

        Write-Verbose "Creating new user '$userPrincipalName'."

        if ($PSCmdlet.ShouldProcess($userPrincipalName, 'Create user')) {
            $creationParams = @{
                AccountEnabled = $true
                DisplayName = $displayName
                GivenName = $firstName
                Surname = $lastName
                UserPrincipalName = $userPrincipalName
                MailNickname = $mailNickname
                UsageLocation = $usageLocation
                PasswordProfile = @{
                    Password                      = $temporaryPassword
                    ForceChangePasswordNextSignIn = $true
                }
            }

            if ($alternateEmail) {
                $creationParams.OtherMails = @($alternateEmail)
            }
            if ($streetAddress) {
                $creationParams.StreetAddress = $streetAddress
            }
            if ($postalCode) {
                $creationParams.PostalCode = $postalCode
            }
            if ($city) {
                $creationParams.City = $city
            }
            if ($mobilePhone) {
                $creationParams.MobilePhone = $mobilePhone
            }

            $newUser = New-MgUser @creationParams

            Set-MgUserLicense `
                -UserId $newUser.Id `
                -AddLicenses @{ SkuId = $licenseSkuId; DisabledPlans = @() } `
                -RemoveLicenses @()

            if ($alternateEmail) {
                Set-AlternateEmailMethod -UserId $newUser.Id -AlternateEmail $alternateEmail
            }

            try {
                Wait-ForMailbox -UserPrincipalName $userPrincipalName -TimeoutSeconds 600 -CheckIntervalSeconds 15
                if ($alternateEmail) {
                    Set-Mailbox -Identity $userPrincipalName -ForwardingSmtpAddress $alternateEmail -DeliverToMailboxAndForward:$true
                }
            } catch {
                Write-Warning "Unable to configure mailbox forwarding for '$userPrincipalName': $($_.Exception.Message)"
            }

            Add-UserToGroup -GroupId $targetGroup.Id -UserId $newUser.Id -GroupName $targetGroup.DisplayName

            $processed += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                Action            = 'Created user'
                Password          = $temporaryPassword
            }
        }
    }

    if ($processed) {
        $processed | Format-Table -AutoSize
    } else {
        Write-Warning 'No users were processed.'
    }
}
finally {
    Write-Verbose 'Disconnecting from Exchange Online.'
    Disconnect-ExchangeOnline -Confirm:$false
    if (Get-Command -Name 'Disconnect-MgGraph' -ErrorAction SilentlyContinue) {
        Write-Verbose 'Disconnecting from Microsoft Graph.'
        Disconnect-MgGraph | Out-Null
    }
}
