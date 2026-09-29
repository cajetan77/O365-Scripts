<#
.SYNOPSIS
    Syncs Entra groups and selected group member emails into SharePoint choice fields.

.DESCRIPTION
    Combines Get-Groups.ps1 + Get-UserEmailfromGroups.ps1:
      1. Mail-enabled + security group display names -> GroupsField (default: Group)
      2. Member emails from -GroupNames only -> UsersField (default: UserEmail)
    Choice values are updated by replacing the field SchemaXml CHOICES block.

.EXAMPLE
    .\Get-Groups.ps1

.EXAMPLE
    .\Get-Groups.ps1 -GroupNames 'Org Users','Allow Group Creators'
#>
[CmdletBinding()]
Param(
    [string[]]$GroupNames = @('Allow Group Creators1', 'Org Users'),
    [string]$SiteUrl = 'https://caje77sharepoint.sharepoint.com/sites/CajIntra/',
    [string]$ListTitle = 'Test1',
    [string]$GroupsField = 'Group',
    [string]$UsersField = 'OtherManager',
    [string]$ClientId = '66a1852a-1f21-46a2-ad58-35fc4c3f1530'
)

$ErrorActionPreference = 'Stop'

function Write-RunbookLog([string]$Message) {
    Write-Output ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Set-ChoiceFieldValues {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Values
    )

    $choices = @(
        $Values |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() } |
        Sort-Object -Unique
    )

    $choicesXml = (
        $choices | ForEach-Object {
            "<CHOICE>$([System.Security.SecurityElement]::Escape($_))</CHOICE>"
        }
    ) -join ''

    $field = Get-PnPField -List $ListTitle -Identity $FieldName -ErrorAction Stop
    if ($field.SchemaXml -notmatch '<CHOICES>') {
        throw "Field '$FieldName' has no CHOICES node (is it a Choice column?)."
    }

    $schema = $field.SchemaXml -replace '(?s)<CHOICES>.*?</CHOICES>', "<CHOICES>$choicesXml</CHOICES>"
    Set-PnPField -List $ListTitle -Identity $field.InternalName -Values @{ SchemaXml = $schema } | Out-Null
    Write-RunbookLog "Replaced '$FieldName' with $($choices.Count) choice(s)."
}

$siteUrl = Get-AzAutomationVariable -Name 'SiteUrl' -ErrorAction SilentlyContinue
if (-not $siteUrl) {
    $siteUrl = 'https://caje77sharepoint.sharepoint.com/sites/CajIntra/'
}
$listTitle = Get-AzAutomationVariable -Name 'ListTitle' -ErrorAction SilentlyContinue
if (-not $listTitle) {
    $listTitle = 'Test1'
}
$userManagedIdentity = Get-AzAutomationVariable -Name 'UserManagedIdentity' -ErrorAction SilentlyContinue
if (-not $userManagedIdentity) {
    $userManagedIdentity = '66a1852a-1f21-46a2-ad58-35fc4c3f1530'
}

Connect-MgGraph -Identity -ClientId $userManagedIdentity -ErrorAction SilentlyContinue | Out-Null

Connect-MgGraph -Scopes 'Group.Read.All', 'User.Read.All' -NoWelcome

# --- 1) Group display names (mail-enabled or security) — same as original Get-Groups ---
Write-RunbookLog 'Loading Entra groups...'
$mailEnabledGroups = @(Get-MgGroup -All | Where-Object { $_.MailEnabled -eq $true })
$securityGroups = @(Get-MgGroup -All | Where-Object { $_.SecurityEnabled -eq $true })
$syncedGroupNames = @(
    @($securityGroups.DisplayName) + @($mailEnabledGroups.DisplayName) |
    Where-Object { $_ } |
    Sort-Object -Unique
)
Write-RunbookLog "Collected $($syncedGroupNames.Count) group name(s)"

# --- 2) Member emails — first matching group in -GroupNames wins; skip the rest ---
$userEmails = [System.Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $GroupNames.Count; $i++) {
    $name = $GroupNames[$i]
    $group = Get-MgGroup -Filter "displayName eq '$name'" | Select-Object -First 1
    if (-not $group) {
        Write-RunbookLog "Group not found: '$name'"
        continue
    }

    Write-RunbookLog "Using group '$name' (ignoring remaining GroupNames)."
    $memberIds = @(Get-MgGroupMember -GroupId $group.Id -All -Property Id | ForEach-Object { $_.Id })
    if ($memberIds.Count -eq 0) {
        Write-RunbookLog "No members in group '$name'"
        break
    }

    foreach ($memberId in $memberIds) {
        try {
            $user = Get-MgUser -UserId $memberId -Property Id, UserPrincipalName, Mail -ErrorAction Stop
        }
        catch {
            Write-RunbookLog "Skipping member '$memberId' (not a user)"
            continue
        }

        Write-RunbookLog "Found user: '$($user.UserPrincipalName)' with email '$($user.Mail)'"
        $email = @($user.UserPrincipalName, $user.Mail) | Where-Object { $_ } | Select-Object -First 1
        if ($email) {
            $userEmails.Add([string]$email)
        }
    }
    break
}
Write-RunbookLog "Collected $($userEmails.Count) user email(s)"

# --- 3) Update SharePoint choice fields ---
Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $siteUrl -ManagedIdentity -UserAssignedManagedIdentityClientId $userManagedIdentity -ErrorAction SilentlyContinue | Out-Null
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

Set-ChoiceFieldValues -ListTitle $ListTitle -FieldName $GroupsField -Values $syncedGroupNames

if ($userEmails.Count -eq 0) {
    Write-RunbookLog 'No users found — skipping UsersField update.'
}
else {
    Set-ChoiceFieldValues -ListTitle $ListTitle -FieldName $UsersField -Values @($userEmails)
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-RunbookLog 'Done.'
