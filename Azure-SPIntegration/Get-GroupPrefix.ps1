#Requires -Version 7.0
<#
.SYNOPSIS
    Builds / parses Shared-Mailbox group naming (US_SMBOX_<Name>_Caj(AD)) and checks if the group exists.

.DESCRIPTION
    Your admin center "Groups" entries (e.g. US_SMBOX_MyFunding1_Caj(AD) /
    MyFunding@keiratheapp.com) are Exchange distribution / mail-enabled security
    groups — use Get-DistributionGroup, not Shared Mailbox / Unified M365 groups.

    Non-interactive by default via Managed Identity.
    Use -Interactive for a browser login locally.

.NOTES
    MI: Exchange.ManageAsApp + Entra Exchange Administrator (or Global Reader for read-only)
    Module: ExchangeOnlineManagement
    Organization: e.g. keiratheapp.com (Automation variable EXCHANGE_ORGANIZATION)

.EXAMPLE
    .\Get-GroupPrefix.ps1 -Organization 'keiratheapp.com' -MailboxName 'MyFunding1'

.EXAMPLE
    .\Get-GroupPrefix.ps1 -Interactive -MailboxName 'MyFunding1'
#>
[CmdletBinding(DefaultParameterSetName = 'ManagedIdentity')]
param(
    [string]$MailboxName = 'MyFunding1',
    [string]$Prefix = 'US_SMBOX',
    [string]$Suffix = 'Caj(AD)',

    [Parameter(ParameterSetName = 'ManagedIdentity')]
    [string]$Organization,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'

# Parse: US_SMBOX_MyFunding1_Caj(AD) -> MyFunding1
$encoded = "${Prefix}_$($MailboxName.Replace(' ', '_'))_$Suffix"
if ($encoded -match "(?<=${Prefix}_).*(?=_$([regex]::Escape($Suffix)))") {
    Write-Output "Parsed display fragment: $($Matches[0].Replace('_', ' ').ToUpper())"
}

$result = "${Prefix}_$($MailboxName.Replace(' ', '_'))_$Suffix"
Write-Output "Encoded group name: $result"

Import-Module ExchangeOnlineManagement -ErrorAction Stop

$alreadyConnected = $false
try {
    $null = Get-ConnectionInformation -ErrorAction Stop
    $alreadyConnected = $true
}
catch { }

if (-not $alreadyConnected) {
    if ($Interactive) {
        Write-Output 'Connecting to Exchange Online (interactive)...'
        Connect-ExchangeOnline -ShowBanner:$false
    }
    else {
        if ([string]::IsNullOrWhiteSpace($Organization)) {
            try {
                $Organization = Get-AutomationVariable -Name 'EXCHANGE_ORGANIZATION' -ErrorAction Stop
            }
            catch {
                $Organization = $env:EXCHANGE_ORGANIZATION
            }
        }

        if ([string]::IsNullOrWhiteSpace($Organization)) {
            throw "Non-interactive mode needs -Organization (e.g. keiratheapp.com) or EXCHANGE_ORGANIZATION. Use -Interactive for browser login."
        }

        Write-Output "Connecting to Exchange Online (Managed Identity / $Organization)..."
        Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false
    }
}

function Find-MailEnabledGroup([string]$Identity) {
    Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
}

# Match by encoded name first (US_SMBOX_MyFunding1_Caj(AD)), then friendly name / SMTP
$found = Find-MailEnabledGroup -Identity $result
if (-not $found) { $found = Find-MailEnabledGroup -Identity $MailboxName }
if (-not $found) {
    $found = Get-DistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -eq $result -or
            $_.DisplayName -eq $MailboxName -or
            $_.Name -eq $result -or
            $_.Alias -eq $MailboxName -or
            $_.PrimarySmtpAddress -like "$MailboxName@*" -or
            $_.PrimarySmtpAddress -like 'MyFunding@*'
        } |
        Select-Object -First 1
}

if ($found) {
    Write-Output ("Group exists: {0} <{1}> ({2})" -f $found.DisplayName, $found.PrimarySmtpAddress, $found.RecipientTypeDetails)
}
else {
    Write-Output "Group does not exist for '$MailboxName' or '$result'"
}

Write-Output ''
Write-Output 'Mail-enabled groups matching US_SMBOX_*_Caj(AD):'
Get-DistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "${Prefix}_*_${Suffix}" -or $_.Name -like "${Prefix}_*_${Suffix}" } |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails |
    Format-Table -AutoSize

if (-not $Interactive -and -not $alreadyConnected) {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
}
