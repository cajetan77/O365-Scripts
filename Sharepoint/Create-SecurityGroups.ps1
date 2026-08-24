<#
.SYNOPSIS
    Creates Entra security groups for a Power BI workspace using the CAJ naming convention.

.DESCRIPTION
    Creates four security groups per workspace:

      CAJ_COMp_PowerBIWorkspace_{WorkspaceName}_{Role}
      CAJ_COMp_PowerBIWorkspace_{WorkspaceName}_Sensitive_{Role}   # when -Sensitive

    Roles: Admin, Members, Contributors, Viewers

    The -AdminUpn user is added as Owner of Members, Contributors, and Viewers.
    The Admin group does not get that user as owner.

.PARAMETER WorkspaceName
    Workspace name segment used in the group display name (no spaces preferred).

.PARAMETER AdminUpn
    User principal name added as Owner on all groups except Admin.

.PARAMETER Sensitive
    When set, inserts _Sensitive_ into the group name.

.EXAMPLE
    .\Create-SecurityGroups.ps1 -WorkspaceName "Finance" -AdminUpn "lester@contoso.com"

.EXAMPLE
    .\Create-SecurityGroups.ps1 -WorkspaceName "HR" -AdminUpn "lester@contoso.com" -Sensitive
#>

[CmdletBinding()]
param(
   
    [string]$WorkspaceName = "Test Workspace 7",

   
    [string]$AdminUpn = "caje77@keiratheapp.com",

    [switch]$Sensitive = $true,

    [string]$NamePrefix = "CAJ_COMp_PowerBIWorkspace",

    [string[]]$Roles = @("Admin", "Members", "Contributors", "Viewers")
)

$ErrorActionPreference = "Stop"

# Sanitize workspace segment for display name / mailNickname
$workspaceSegment = ($WorkspaceName -replace '[^a-zA-Z0-9\-]', '').Trim()
if ([string]::IsNullOrWhiteSpace($workspaceSegment)) {
    throw "WorkspaceName must contain at least one alphanumeric character."
}

Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All", "Directory.Read.All" -NoWelcome

$adminUser = Get-MgUser -UserId $AdminUpn -ErrorAction SilentlyContinue
if (-not $adminUser) {
    # Try filter by UPN if -UserId didn't resolve
    $adminUser = Get-MgUser -Filter "userPrincipalName eq '$AdminUpn'" -ErrorAction Stop
}
if (-not $adminUser) {
    throw "Admin user not found: $AdminUpn"
}
Write-Host "Admin owner user: $($adminUser.DisplayName) [$($adminUser.Id)]" -ForegroundColor Green

function Get-GroupDisplayName {
    param([string]$Role)

    $parts = @($NamePrefix, $workspaceSegment)
    if ($Sensitive) { $parts += "Sensitive" }
    $parts += $Role
    return ($parts -join "_")
}

function Get-MailNickname {
    param([string]$DisplayName)
    # MailNickname: letters, numbers, hyphens only; max 64
    $nick = ($DisplayName -replace '[^a-zA-Z0-9]', '')
    if ($nick.Length -gt 64) { $nick = $nick.Substring(0, 64) }
    return $nick
}

$created = @()

foreach ($role in $Roles) {
    $displayName = Get-GroupDisplayName -Role $role
    $mailNickname = Get-MailNickname -DisplayName $displayName
    $description = "Power BI workspace access group. Workspace=$WorkspaceName; Role=$role" + $(if ($Sensitive) { "; Sensitive=true" } else { "" })

    Write-Host ""
    Write-Host "Group: $displayName" -ForegroundColor Cyan

    $existing = Get-MgGroup -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Already exists [$($existing.Id)]" -ForegroundColor Yellow
        $group = $existing
    }
    else {
        try {
            $group = New-MgGroup `
                -DisplayName $displayName `
                -Description $description `
                -MailEnabled:$false `
                -SecurityEnabled:$true `
                -MailNickname $mailNickname `
                -GroupTypes @()
            Write-Host "  Created [$($group.Id)]" -ForegroundColor Green
        }
        catch {
            Write-Host "  Create failed: $_" -ForegroundColor Red
            continue
        }
    }

    # Owner: Admin UPN on all groups EXCEPT Admin
    if ($role -ne "Admin") {
        $owners = @(Get-MgGroupOwner -GroupId $group.Id -ErrorAction SilentlyContinue)
        $alreadyOwner = $owners | Where-Object { $_.Id -eq $adminUser.Id }
        if ($alreadyOwner) {
            Write-Host "  Owner already set: $AdminUpn" -ForegroundColor DarkYellow
        }
        else {
            try {
                $body = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($adminUser.Id)"
                }
                New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter $body
                Write-Host "  Owner added: $AdminUpn" -ForegroundColor Green
            }
            catch {
                Write-Host "  Owner add failed: $_" -ForegroundColor Red
            }
        }
    }
    else {
       
        try {
            $body = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($adminUser.Id)"
            }
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $body
            Write-Host "  Member added: $AdminUpn" -ForegroundColor Green
        }
        catch {
            Write-Host "  Member add failed: $_" -ForegroundColor Red
        }
    }
}