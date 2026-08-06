<#
.SYNOPSIS
    Adds a user to a distribution list or mail-enabled security group.

.DESCRIPTION
    Simple Azure Automation runbook: pass the user and the group they requested.
    Uses Exchange Online app-only (Managed Identity) to add membership.

.NOTES
    Required
    --------
    - System-assigned Managed Identity
    - Exchange.ManageAsApp + Entra role "Exchange Administrator" on the MI
    - Module: ExchangeOnlineManagement
    - Automation variable: EXCHANGE_ORGANIZATION (e.g. contoso.onmicrosoft.com)
      Optional if Microsoft.Graph.Authentication is available (auto-detects domain).

    After editing this runbook in Azure Automation: Save, then Publish.
    Job logs: open the job -> Output (click Refresh) or All Logs.

.EXAMPLE
    .\Request-GroupAccess.ps1 -UserUpn 'jane@contoso.com' -GroupName 'Finance DL'
#>
[CmdletBinding()]
Param(
   
    [string]$UserUpn = "AaGswM.JZesPk@cajesharepoint.onmicrosoft.com",

    [string]$GroupName = "Z_Infra"
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'

function Write-RunbookLog {
    param([string]$Message)

    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Output -InputObject $line
    Write-Warning -Message $line
}

function Get-ExchangeOrganization {
    $fromVar = Get-AutomationVariable -Name 'EXCHANGE_ORGANIZATION' -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($fromVar)) {
        return , $fromVar.Trim()
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Identity -NoWelcome
    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains'
        $org = @($response.value) | Select-Object -First 1
        $domains = @($org.verifiedDomains)
        $initial = $domains | Where-Object { $_.isInitial -eq $true } | Select-Object -First 1
        if ($initial) { return , [string]$initial.name }
        $first = $domains | Select-Object -First 1
        if ($first) { return , [string]$first.name }
        return , $null
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

function Test-IsAlreadyGroupMember {
    param(
        [Parameter(Mandatory)][string]$GroupName,
        [Parameter(Mandatory)][string]$UserUpn
    )

    $members = @(Get-DistributionGroupMember -Identity $GroupName -ResultSize Unlimited -ErrorAction Stop)
    $target = $UserUpn.Trim().ToLowerInvariant()

    foreach ($member in $members) {
        $candidates = @(
            [string]$member.PrimarySmtpAddress
            [string]$member.WindowsLiveID
            [string]$member.ExternalEmailAddress
            [string]$member.Alias
            [string]$member.Name
            [string]$member.DisplayName
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidate in $candidates) {
            $normalized = $candidate.Trim().ToLowerInvariant() -replace '^smtp:', ''
            if ($normalized -eq $target) {
                return $true
            }
        }
    }

    return $false
}

Write-RunbookLog 'Request-GroupAccess starting...'
Write-RunbookLog ("Parameters: UserUpn='{0}'; GroupName='{1}'" -f $UserUpn, $GroupName)

try {
    $UserUpn = $UserUpn.Trim()
    $GroupName = $GroupName.Trim()

    if ([string]::IsNullOrWhiteSpace($UserUpn) -or [string]::IsNullOrWhiteSpace($GroupName)) {
        throw 'Both UserUpn and GroupName are required.'
    }

    Write-RunbookLog 'Resolving Exchange organization...'
    $organization = Get-ExchangeOrganization
    if ($organization -is [System.Array]) {
        $organization = $organization | Select-Object -Last 1
    }
    if ([string]::IsNullOrWhiteSpace($organization)) {
        throw 'Set automation variable EXCHANGE_ORGANIZATION (e.g. contoso.onmicrosoft.com).'
    }
    Write-RunbookLog "Exchange organization: $organization"

    Write-RunbookLog 'Importing ExchangeOnlineManagement...'
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Write-RunbookLog "Connecting to Exchange Online as '$organization'..."
    Connect-ExchangeOnline -ManagedIdentity -Organization $organization -ShowBanner:$false
    Write-RunbookLog 'Connected to Exchange Online.'

    Write-RunbookLog "Checking if '$UserUpn' is already a member of '$GroupName'..."
    if (Test-IsAlreadyGroupMember -GroupName $GroupName -UserUpn $UserUpn) {
        Write-RunbookLog "Already a member: '$UserUpn' is already in '$GroupName'."
        Write-RunbookLog 'Request-GroupAccess finished.'
        return
    }

    Write-RunbookLog "Adding '$UserUpn' to group '$GroupName'..."
    $addErr = $null
    Add-DistributionGroupMember -Identity $GroupName -Member $UserUpn -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue -ErrorVariable addErr | Out-Null

    if ((-not $addErr) -and $?) {
        Write-RunbookLog "Success: added '$UserUpn' to '$GroupName'."
    }
    else {
        $message = if ($addErr -and $addErr.Count -gt 0) {
            [string]$addErr[0].Exception.Message
        }
        else {
            'Unknown error (Exchange did not return error details).'
        }

        if ($message -match 'already exist|already a member|is already a member|already present') {
            Write-RunbookLog "Already a member: '$UserUpn' is already in '$GroupName'."
        }
        else {
            throw "Failed to add '$UserUpn' to '$GroupName'. $message"
        }
    }

    Write-RunbookLog 'Request-GroupAccess finished.'
}
catch {
    Write-RunbookLog "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    Write-RunbookLog 'Cleaning up connections...'
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
