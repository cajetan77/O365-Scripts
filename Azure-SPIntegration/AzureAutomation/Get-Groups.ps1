<#
.SYNOPSIS
    Syncs Entra DLs / mail-enabled security groups into SharePoint choice fields,
    and optionally adds a user to all of those groups (app-only / Managed Identity).

.DESCRIPTION
    Ensures the SharePoint list "Groups" exists (creates it if missing), ensures the
    choice fields exist, then syncs Entra mail-enabled groups into those fields:
      - "Distribution List"  <- mail-enabled, non-security groups (classic DLs)
      - "Mail Enabled"       <- mail-enabled security groups (not Microsoft 365 groups)

    If automation variable GROUP_MEMBER_USER_UPN is set, adds that user to every
    synced group using application permissions:
      - Mail-enabled security groups -> Microsoft Graph (GroupMember.ReadWrite.All)
      - Classic distribution lists    -> Exchange Online app-only (Exchange.ManageAsApp)

.NOTES
    Azure Automation setup
    ----------------------
    1. System-assigned Managed Identity on the Automation Account.
    2. Application permissions (admin consent) — re-run Set-SystemManagedId.ps1:
         Group.Read.All
         GroupMember.ReadWrite.All
         User.Read.All
         Organization.Read.All
         Sites.ReadWrite.All (or SharePoint Sites.FullControl.All)
         Exchange Online: Exchange.ManageAsApp
       Also assign the Managed Identity the Entra role "Exchange Administrator"
       (required for Exchange Online app-only recipient management).
    3. Runtime modules (same Graph version, e.g. 2.38.1):
         Microsoft.Graph.Authentication
         Microsoft.Graph.Groups
         Microsoft.Graph.Users
         ExchangeOnlineManagement
         PnP.PowerShell
       Load order matters: Graph -> Exchange -> PnP (avoid MSAL assembly clashes).
    4. Automation variables:
         SHAREPOINT_SITE_URL       (required)
         GROUP_MEMBER_USER_UPN     (optional; e.g. admin@contoso.com)
         EXCHANGE_ORGANIZATION     (optional; e.g. contoso.onmicrosoft.com —
                                   auto-detected from Graph if unset)
#>
[CmdletBinding()]
Param(
    [string]$ListTitle = 'Groups',
    [string]$DistributionListField = 'Distribution List',
    [string]$DistributionListInternalName = 'DistributionList',
    [string]$MailEnabledField = 'Mail Enabled',
    [string]$MailEnabledInternalName = 'MailEnabled',
    [string]$MemberUserUpn
)

$ErrorActionPreference = 'Stop'

$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
    throw 'Automation variable SHAREPOINT_SITE_URL is required (e.g. https://contoso.sharepoint.com/sites/IT).'
}

if ([string]::IsNullOrWhiteSpace($MemberUserUpn)) {
    $MemberUserUpn = Get-AutomationVariable -Name 'GROUP_MEMBER_USER_UPN' -ErrorAction SilentlyContinue
}

$ExchangeOrganization = Get-AutomationVariable -Name 'EXCHANGE_ORGANIZATION' -ErrorAction SilentlyContinue

function Write-RunbookLog {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    # Use Write-Host so log lines do not become function return values (Write-Output would).
    Write-Host "[$stamp] $Message"
}

function Import-MatchingGraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Users'
    )

    $authVersions = @(
        Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
            Select-Object -ExpandProperty Version -Unique |
            Sort-Object -Descending
    )

    if ($authVersions.Count -eq 0) {
        throw 'Microsoft.Graph.Authentication is not installed in this Automation Account.'
    }

    $commonVersion = $null
    foreach ($version in $authVersions) {
        $missing = @()
        foreach ($moduleName in $requiredModules) {
            $match = Get-Module -ListAvailable -Name $moduleName |
                Where-Object { $_.Version -eq $version } |
                Select-Object -First 1
            if (-not $match) {
                $missing += $moduleName
            }
        }

        if ($missing.Count -eq 0) {
            $commonVersion = $version
            break
        }
    }

    if (-not $commonVersion) {
        $installed = foreach ($moduleName in $requiredModules) {
            $versions = @(
                Get-Module -ListAvailable -Name $moduleName |
                    Select-Object -ExpandProperty Version -Unique |
                    Sort-Object -Descending
            )
            "  $moduleName : $(if ($versions) { $versions -join ', ' } else { '(not installed)' })"
        }

        throw @"
Microsoft Graph modules must all be the SAME version in Azure Automation.
Install matching versions of:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Groups
  - Microsoft.Graph.Users

Currently available:
$($installed -join [Environment]::NewLine)
"@
    }

    foreach ($moduleName in $requiredModules) {
        Import-Module $moduleName -RequiredVersion $commonVersion -Force -ErrorAction Stop
    }

    Write-RunbookLog "Loaded Microsoft Graph modules version $commonVersion"
}

function Get-InitialDomainName {
    $response = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains'
    $org = @($response.value) | Select-Object -First 1
    if (-not $org) { return $null }

    $domains = @($org.verifiedDomains)
    $initial = $domains | Where-Object { $_.isInitial -eq $true } | Select-Object -First 1
    if ($initial) { return [string]$initial.name }

    $first = $domains | Select-Object -First 1
    if ($first) { return [string]$first.name }
    return $null
}

function Test-AlreadyGroupMemberError {
    param($ErrorRecord)
    $message = [string]$ErrorRecord.Exception.Message
    return ($message -match 'already exist|already a member|added object references already exist')
}

function Add-UserToMailEnabledSecurityGroups {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$UserUpn,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Groups
    )

    $added = 0
    $skipped = 0
    $failed = 0

    foreach ($group in $Groups) {
        try {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $UserId -ErrorAction Stop | Out-Null
            Write-RunbookLog "Graph: added $UserUpn to mail-enabled security group '$($group.DisplayName)'."
            $added++
        }
        catch {
            if (Test-AlreadyGroupMemberError -ErrorRecord $_) {
                Write-RunbookLog "Graph: $UserUpn already in '$($group.DisplayName)'."
                $skipped++
            }
            else {
                Write-RunbookLog "Graph ERROR on '$($group.DisplayName)': $($_.Exception.Message)"
                $failed++
            }
        }
    }

    Write-RunbookLog "Graph membership summary: added=$added skipped=$skipped failed=$failed"
    return [int]$failed
}

function Add-UserToDistributionGroups {
    param(
        [Parameter(Mandatory)][string]$UserUpn,
        [Parameter(Mandatory)][string]$Organization,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Groups
    )

    if ($Groups.Count -eq 0) {
        Write-RunbookLog 'No distribution lists to update via Exchange.'
        return 0
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-RunbookLog "Connecting to Exchange Online (app-only / Managed Identity) as org '$Organization'..."
    Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false

    $added = 0
    $skipped = 0
    $failed = 0

    try {
        foreach ($group in $Groups) {
            $identity = if ($group.Mail) { $group.Mail } else { $group.Id }
            try {
                Add-DistributionGroupMember -Identity $identity -Member $UserUpn -BypassSecurityGroupManagerCheck -ErrorAction Stop
                Write-RunbookLog "Exchange: added $UserUpn to distribution list '$($group.DisplayName)'."
                $added++
            }
            catch {
                if (Test-AlreadyGroupMemberError -ErrorRecord $_) {
                    Write-RunbookLog "Exchange: $UserUpn already in '$($group.DisplayName)'."
                    $skipped++
                }
                else {
                    Write-RunbookLog "Exchange ERROR on '$($group.DisplayName)': $($_.Exception.Message)"
                    $failed++
                }
            }
        }
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }

    Write-RunbookLog "Exchange membership summary: added=$added skipped=$skipped failed=$failed"
    return [int]$failed
}

function Initialize-GroupsList {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$DistributionListField,
        [Parameter(Mandatory)][string]$DistributionListInternalName,
        [Parameter(Mandatory)][string]$MailEnabledField,
        [Parameter(Mandatory)][string]$MailEnabledInternalName
    )

    $list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-RunbookLog "Creating list '$ListTitle'..."
        $null = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch -ErrorAction Stop
        Write-RunbookLog "Created list '$ListTitle'."
    }
    else {
        Write-RunbookLog "List '$ListTitle' already exists."
    }

    Initialize-ChoiceField -ListTitle $ListTitle -DisplayName $DistributionListField -InternalName $DistributionListInternalName
    Initialize-ChoiceField -ListTitle $ListTitle -DisplayName $MailEnabledField -InternalName $MailEnabledInternalName
}

function Resolve-PnPListField {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$InternalName,
        [Parameter(Mandatory)][string]$DisplayName
    )

    $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        $field = Get-PnPField -List $ListTitle -Identity $DisplayName -ErrorAction SilentlyContinue
    }
    if (-not $field) {
        throw "Choice field not found on list '$ListTitle'. Tried '$InternalName' and '$DisplayName'."
    }
    return $field
}

function Initialize-ChoiceField {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$InternalName
    )

    $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        $field = Get-PnPField -List $ListTitle -Identity $DisplayName -ErrorAction SilentlyContinue
    }

    if (-not $field) {
        Write-RunbookLog "Creating choice field '$DisplayName' ($InternalName) on list '$ListTitle'..."
        Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Choice -AddToDefaultView -ErrorAction Stop | Out-Null
        Write-RunbookLog "Created field '$DisplayName'."
        return
    }

    Write-RunbookLog "Field '$DisplayName' already exists (InternalName=$($field.InternalName))."
}

function Set-PnPChoiceFieldValues {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$InternalName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Choices
    )

    $uniqueChoices = @(
        $Choices |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Sort-Object -Unique
    )

    $field = Resolve-PnPListField -ListTitle $ListTitle -InternalName $InternalName -DisplayName $DisplayName
    [xml]$schemaXml = $field.SchemaXml

    $choicesNode = $schemaXml.SelectSingleNode('/Field/CHOICES')
    if (-not $choicesNode) {
        $choicesNode = $schemaXml.CreateElement('CHOICES')
        [void]$schemaXml.DocumentElement.AppendChild($choicesNode)
    }
    else {
        $choicesNode.RemoveAll()
    }

    foreach ($choice in $uniqueChoices) {
        $choiceNode = $schemaXml.CreateElement('CHOICE')
        $choiceNode.InnerText = $choice
        [void]$choicesNode.AppendChild($choiceNode)
    }

    if ($schemaXml.Field.HasAttribute('FillInChoice')) {
        $schemaXml.Field.SetAttribute('FillInChoice', 'FALSE')
    }
    if ($schemaXml.Field.HasAttribute('Format')) {
        $schemaXml.Field.SetAttribute('Format', 'Dropdown')
    }

    Set-PnPField -List $ListTitle -Identity $field.InternalName -Values @{ SchemaXml = $schemaXml.OuterXml } -ErrorAction Stop
    Write-RunbookLog "Updated '$DisplayName' ($($field.InternalName)) with $($uniqueChoices.Count) choice(s)."
}

try {
    Import-MatchingGraphModules

    Write-RunbookLog 'Connecting to Microsoft Graph with Managed Identity...'
    Connect-MgGraph -Identity -NoWelcome

    Write-RunbookLog 'Retrieving groups...'
    $allGroups = @(Get-MgGroup -All -Property Id, DisplayName, MailEnabled, SecurityEnabled, GroupTypes, Mail)

    $distributionGroups = @(
        $allGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            $_.SecurityEnabled -eq $false
        }
    )

    $mailEnabledSecurityGroups = @(
        $allGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            $_.SecurityEnabled -eq $true -and
            (@($_.GroupTypes) -notcontains 'Unified')
        }
    )

    Write-RunbookLog "Found $($distributionGroups.Count) distribution list(s) and $($mailEnabledSecurityGroups.Count) mail-enabled security group(s)."

    $membershipFailures = 0
    if (-not [string]::IsNullOrWhiteSpace($MemberUserUpn)) {
        Write-RunbookLog "Resolving user '$MemberUserUpn' for group membership (app permissions)..."
        $memberUser = Get-MgUser -UserId $MemberUserUpn -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
        Write-RunbookLog "Resolved user: $($memberUser.DisplayName) ($($memberUser.UserPrincipalName))"

        $membershipFailures = [int]$membershipFailures + [int](Add-UserToMailEnabledSecurityGroups `
            -UserId $memberUser.Id `
            -UserUpn $memberUser.UserPrincipalName `
            -Groups $mailEnabledSecurityGroups)

        if ([string]::IsNullOrWhiteSpace($ExchangeOrganization)) {
            $ExchangeOrganization = Get-InitialDomainName
        }
        if ([string]::IsNullOrWhiteSpace($ExchangeOrganization)) {
            throw 'EXCHANGE_ORGANIZATION is required to add the user to classic distribution lists (e.g. contoso.onmicrosoft.com).'
        }

        # Disconnect Graph before Exchange to reduce MSAL conflicts.
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $membershipFailures = [int]$membershipFailures + [int](Add-UserToDistributionGroups `
            -UserUpn $memberUser.UserPrincipalName `
            -Organization $ExchangeOrganization `
            -Groups $distributionGroups)
    }
    else {
        Write-RunbookLog 'GROUP_MEMBER_USER_UPN not set — skipping group membership updates.'
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Import-Module PnP.PowerShell -ErrorAction Stop

    Write-RunbookLog "Connecting to SharePoint: $SharePointSiteUrl"
    Connect-PnPOnline -Url $SharePointSiteUrl -ManagedIdentity

    Initialize-GroupsList `
        -ListTitle $ListTitle `
        -DistributionListField $DistributionListField `
        -DistributionListInternalName $DistributionListInternalName `
        -MailEnabledField $MailEnabledField `
        -MailEnabledInternalName $MailEnabledInternalName

    Set-PnPChoiceFieldValues `
        -ListTitle $ListTitle `
        -InternalName $DistributionListInternalName `
        -DisplayName $DistributionListField `
        -Choices @($distributionGroups.DisplayName)

    Set-PnPChoiceFieldValues `
        -ListTitle $ListTitle `
        -InternalName $MailEnabledInternalName `
        -DisplayName $MailEnabledField `
        -Choices @($mailEnabledSecurityGroups.DisplayName)

    if ($membershipFailures -gt 0) {
        throw "SharePoint sync completed, but $membershipFailures group membership update(s) failed. See log above."
    }

    Write-RunbookLog 'Group choice sync complete.'
}
catch {
    Write-RunbookLog "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
