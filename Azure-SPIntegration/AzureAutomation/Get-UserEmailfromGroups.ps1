[CmdletBinding()]
Param(
    [string[]]$GroupNames = @('Allow Group Creators', 'Org Users1'),
    [string]$SiteUrl = 'https://caje77sharepoint.sharepoint.com/sites/CajIntra/',
    [string]$ListTitle = 'Test1',
    [string]$ChoiceField = 'Group',
    [string]$ClientId = '66a1852a-1f21-46a2-ad58-35fc4c3f1530'
)

$ErrorActionPreference = 'Stop'

Connect-MgGraph -Scopes 'Group.Read.All', 'User.Read.All' -NoWelcome

$userEmails = [System.Collections.Generic.List[string]]::new()

foreach ($GroupName in $GroupNames) {
    $group = Get-MgGroup -Filter "displayName eq '$GroupName'" | Select-Object -First 1
    if (-not $group) {
        Write-Output "Group not found: '$GroupName'"
        continue
    }

    $members = @(Get-MgGroupMember -GroupId $group.Id -All -Property "id").Id
    if ($members.Count -eq 0) {
        Write-Output "No members in group '$GroupName'"
        continue
    }

    foreach ($member in $members) {
        try {
            $user = Get-MgUser -UserId $member -Property Id, UserPrincipalName, Mail -ErrorAction Stop
            if ($user) {
                Write-Output "Found user: '$($user.UserPrincipalName)' with email '$($user.Mail)'"
            }
        }
        catch {
            Write-Output "Skipping member '$($member.Id)' (not a user)"
            continue
        }

        $email = @($user.UserPrincipalName, $user.Mail) | Where-Object { $_ } | Select-Object -First 1
        if ($email) {
            $userEmails.Add($email)
        }
    }
}

Write-Output "Collected $($userEmails.Count) user email(s)"

if ($userEmails.Count -eq 0) {
    Write-Output "No users found"
    exit
}

Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# Replace all Group choices (Choices.Clear/Add is a fixed array and fails)
$choicesXml = (
    $userEmails |
    Where-Object { $_ } |
    Sort-Object -Unique |
    ForEach-Object { "<CHOICE>$([System.Security.SecurityElement]::Escape($_))</CHOICE>" }
) -join ''

$field = Get-PnPField -List $ListTitle -Identity $ChoiceField
$schema = $field.SchemaXml -replace '(?s)<CHOICES>.*?</CHOICES>', "<CHOICES>$choicesXml</CHOICES>"
Set-PnPField -List $ListTitle -Identity $field.InternalName -Values @{ SchemaXml = $schema } | Out-Null
Write-Output "Replaced '$ChoiceField' choices with $($userEmails.Count) email(s)."

Disconnect-PnPOnline -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
