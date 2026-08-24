<#
.SYNOPSIS
    Creates a new Power BI workspace using interactive login.

.DESCRIPTION
    Same approach as Microsoft's sample:
      https://github.com/microsoft/PowerBI-Developer-Samples/blob/master/PowerShell%20Scripts/Create-Workspace.ps1

      Connect-PowerBIServiceAccount
      New-PowerBIGroup -Name $Name

.EXAMPLE
    .\New-BIWorkspace.ps1 -Name "Test Workspace 1"

.EXAMPLE
    .\New-BIWorkspace.ps1 -Name "Finance BI" -AdminUserUpns "user@contoso.com"

.EXAMPLE
    .\New-BIWorkspace.ps1 -Name "Finance"
    Resolves CAJ_COMp_PowerBIWorkspace_Finance[_Sensitive]_{Admin|Members|Contributors|Viewers}
    (Sensitive segment optional) and adds them as workspace Admin, Member, Contributor, Viewer.
#>

[CmdletBinding()]
param(
   
    [string]$Name = "Test Workspace 7",

    [string]$Description = "Test Description",

    [string[]]$AdminUserUpns = @(),
    [string[]]$AdminAppIds = @(),

    [ValidateSet("Admin", "Member", "Contributor", "Viewer")]
    [string]$DefaultAccessRight = "Admin",

    # Matches Create-SecurityGroups.ps1 naming → Power BI workspace roles.
    # Looks up both with and without _Sensitive_ in the display name.
    [string]$NamePrefix = "CAJ_COMp_PowerBIWorkspace",
    [string]$SecurityGroupWorkspaceName,
    [switch]$SkipSecurityGroups,

    [guid]$CapacityId,
    [switch]$SkipIfExists
)

$ErrorActionPreference = "Stop"

Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop

Write-Host "Connecting to Power BI (standard login)..." -ForegroundColor Cyan
Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null

try {
    # Prefer browser / interactive login (Microsoft sample)
    Connect-PowerBIServiceAccount -ErrorAction Stop | Out-Null
}
catch {
    # Common on machines with Microsoft.Graph / newer MSAL:
    # Method not found: ...WithBroker(... BrokerOptions)
    Write-Host "Interactive broker login failed (MSAL assembly conflict)." -ForegroundColor Yellow
}
Write-Host "Connected." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Create workspace
# ---------------------------------------------------------------------------
Write-Host "Checking for existing workspace '$Name'..." -ForegroundColor Cyan
$existing = @(Get-PowerBIWorkspace -Name $Name -ErrorAction SilentlyContinue)
if ($existing.Count -gt 0) {
    $workspace = $existing[0]
    Write-Host "Workspace already exists: $($workspace.Name) [$($workspace.Id)]" -ForegroundColor Yellow
    if ($SkipIfExists) {
        Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
        return $workspace
    }
}
else {
    Write-Host "Creating workspace '$Name'..." -ForegroundColor Yellow
    $workspace = New-PowerBIGroup -Name $Name
    Write-Host "Created: $($workspace.Name) [$($workspace.Id)]" -ForegroundColor Green
}

if ($Description) {
    try {
        Set-PowerBIWorkspace -Id $workspace.Id -Description $Description -Scope Organization
        Write-Host "Description set." -ForegroundColor Green
    }
    catch {
        Write-Host "Could not set description: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

foreach ($upn in $AdminUserUpns) {
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }
    Write-Host "Adding user $DefaultAccessRight : $upn" -ForegroundColor Cyan
    try {
        Add-PowerBIWorkspaceUser -Id $workspace.Id -UserPrincipalName $upn -AccessRight $DefaultAccessRight
        Write-Host "  Added." -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

foreach ($appId in $AdminAppIds) {
    if ([string]::IsNullOrWhiteSpace($appId)) { continue }
    Write-Host "Adding app $DefaultAccessRight : $appId" -ForegroundColor Cyan
    try {
        Add-PowerBIWorkspaceUser -Id $workspace.Id -AccessRight $DefaultAccessRight -PrincipalType App -Identifier $appId
        Write-Host "  Added." -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------------
# Add security groups (Create-SecurityGroups.ps1 naming) as workspace roles
# ---------------------------------------------------------------------------
if (-not $SkipSecurityGroups) {
    $segmentSource = if ($SecurityGroupWorkspaceName) { $SecurityGroupWorkspaceName } else { $Name }
    $workspaceSegment = ($segmentSource -replace '[^a-zA-Z0-9\-]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($workspaceSegment)) {
        Write-Host "Skipping security groups: no alphanumeric segment from workspace name." -ForegroundColor DarkYellow
    }
    else {
        # Group role suffix (Create-SecurityGroups) → Power BI AccessRight
        $roleMap = [ordered]@{
            Admin        = "Admin"
            Members      = "Member"
            Contributors = "Contributor"
            Viewers      = "Viewer"
        }

        try {
            Import-Module Microsoft.Graph.Groups -ErrorAction Stop
            Connect-MgGraph -Scopes "Group.Read.All" -NoWelcome | Out-Null
        }
        catch {
            Write-Host "Failed to connect to Graph for security groups: $($_.Exception.Message)" -ForegroundColor Red
            $roleMap = $null
        }

        if ($roleMap) {
            foreach ($groupRole in $roleMap.Keys) {
                $accessRight = $roleMap[$groupRole]
                # Sensitive may or may not be in the group name — try both
                $candidateNames = @(
                    ($NamePrefix, $workspaceSegment, "Sensitive", $groupRole) -join "_"
                    ($NamePrefix, $workspaceSegment, $groupRole) -join "_"
                )

                Write-Host "Resolving $groupRole security group (with/without Sensitive)..." -ForegroundColor Cyan
                try {
                    $group = $null
                    $matchedName = $null
                    foreach ($candidateName in $candidateNames) {
                        $escapedName = $candidateName.Replace("'", "''")
                        $group = Get-MgGroup -Filter "displayName eq '$escapedName'" -ErrorAction Stop |
                        Select-Object -First 1
                        if ($group) {
                            $matchedName = $candidateName
                            break
                        }
                    }

                    if (-not $group) {
                        Write-Host "  Not found as: $($candidateNames -join ' | '). Create with Create-SecurityGroups.ps1." -ForegroundColor Yellow
                        continue
                    }

                    Write-Host "  Found: $matchedName [$($group.Id)]" -ForegroundColor Green
                    Write-Host "Adding as workspace $accessRight..." -ForegroundColor Cyan
                    Add-PowerBIWorkspaceUser -Id $workspace.Id -AccessRight $accessRight -PrincipalType Group -Identifier $group.Id
                    Write-Host "  $groupRole group added as $accessRight." -ForegroundColor Green
                }
                catch {
                    Write-Host "  Failed to add $groupRole security group: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    }
}

Write-Host "Removing direct user access (PrincipalType User)..." -ForegroundColor Cyan
$wsWithUsers = Get-PowerBIWorkspace -Id $workspace.Id -Scope Organization -Include All
foreach ($u in @($wsWithUsers.Users)) {
    if ([string]$u.PrincipalType -ne "User") { continue }

    $upn = if ($u.UserPrincipalName) { $u.UserPrincipalName } else { $u.Identifier }
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    Write-Host "  Removing user: $upn" -ForegroundColor Yellow
    try {
        Remove-PowerBIWorkspaceUser -Id $workspace.Id -UserPrincipalName $upn -Scope Organization -ErrorAction Stop
        Write-Host "    Removed." -ForegroundColor Green
    }
    catch {
        Write-Host "    Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$workspace = Get-PowerBIWorkspace -Id $workspace.Id
Write-Host ""
Write-Host "Workspace: $($workspace.Name)" -ForegroundColor Green
Write-Host "Id      : $($workspace.Id)" -ForegroundColor White
Write-Host "URL     : https://app.powerbi.com/groups/$($workspace.Id)" -ForegroundColor White

Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
return $workspace
