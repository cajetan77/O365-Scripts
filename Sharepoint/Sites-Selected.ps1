<#
.SYNOPSIS
    Grants an Entra app (Sites.Selected) write access to a specific SharePoint site via Microsoft Graph.

.NOTES
    Sign-in account needs Sites.FullControl.All (delegated) or use an app with Sites.FullControl.All.
    Target app must have Sites.Selected (application) with admin consent.
#>

[CmdletBinding()]
param(
    [string]$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/M365Updates",
    [string]$AppId = "355afea0-ab69-4219-bb75-831e3fb42a6e",
    [string]$AppDisplayName = "M365 Reporting",
    [ValidateSet("read", "write")]
    [string]$Role = "write",
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"

# Load selectedAppID from config when present
if (-not $ConfigPath) {
    $ConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"
    if (-not (Test-Path $ConfigPath)) { $ConfigPath = Join-Path $PSScriptRoot "..\config.json" }
    if (-not (Test-Path $ConfigPath)) { $ConfigPath = ".\config.json" }
}
if ((Test-Path $ConfigPath) -and -not $PSBoundParameters.ContainsKey("AppId")) {
    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
    if ($config.selectedAppID) { $AppId = $config.selectedAppID }
}

Import-Module Microsoft.Graph.Sites -ErrorAction Stop

Write-Host "Connecting to Microsoft Graph (Sites.FullControl.All)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Sites.FullControl.All" -NoWelcome

# Resolve site by path (Get-MgSite -Search is deprecated and can return bad/multiple results)
# Format: hostname:/sites/sitename
$uri = [Uri]$SiteUrl
$sitePathId = "{0}:{1}" -f $uri.Host, $uri.AbsolutePath.TrimEnd("/")
Write-Host "Resolving site: $sitePathId" -ForegroundColor Gray

$site = Get-MgSite -SiteId $sitePathId
if (-not $site -or -not $site.Id) {
    throw "Site not found for '$SiteUrl' (tried SiteId '$sitePathId')."
}

Write-Host "Site: $($site.DisplayName) [$($site.Id)]" -ForegroundColor Green

# Correct Graph body: roles = string[], grantedToIdentities = array of identity objects
$body = @{
    roles               = @($Role)
    grantedToIdentities = @(
        @{
            application = @{
                id          = $AppId
                displayName = $AppDisplayName
            }
        }
    )
}

Write-Host "Granting '$Role' on site to app $AppId ($AppDisplayName)..." -ForegroundColor Yellow
try {
    $permission = New-MgSitePermission -SiteId $site.Id -BodyParameter $body
    Write-Host "Granted. Permission Id: $($permission.Id) Roles: $($permission.Roles -join ', ')" -ForegroundColor Green
}
catch {
    $msg = $_.ErrorDetails.Message
    if (-not $msg) { $msg = $_.Exception.Message }
    Write-Host "Failed: $msg" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common causes:" -ForegroundColor Yellow
    Write-Host "  - Signed-in user lacks Full Control on the site / Sites.FullControl.All consent" -ForegroundColor White
    Write-Host "  - App Id is wrong (use Application/client ID of the Sites.Selected app)" -ForegroundColor White
    Write-Host "  - Permission already exists (list with Get-MgSitePermission -SiteId '$($site.Id)')" -ForegroundColor White
    throw
}

Write-Host ""
Write-Host "Current site permissions:" -ForegroundColor Cyan
Get-MgSitePermission -SiteId $site.Id | ForEach-Object {
    $app = $_.GrantedToIdentities.Application
    if (-not $app) { $app = $_.GrantedToIdentitiesV2.Application }
    Write-Host ("  {0} -> {1} ({2})" -f ($_.Roles -join ","), $app.DisplayName, $app.Id) -ForegroundColor Gray
}




# Install module if needed
Install-Module Microsoft.Graph.Sites
# Connect with admin account
Connect-MgGraph -Scopes "Sites.FullControl.All"
# Grant the app read/write access to a specific SharePoint site
$appId = "0ba8f02c-9063-409b-a35d-92ce5dabf069"
$siteId = (Get-MgSite -Search "M365Updates").Id
$params = @{
    roles               = @(
        "read"
    )
    grantedToIdentities = @{
        application = @{
            id          = "0ba8f02c-9063-409b-a35d-92ce5dabf069"
            displayName = "SitesSelected"
        }
    }
}
New-MgSitePermission -SiteId $siteId -BodyParameter $params




Connect-PnPOnline -Url "https://caje77sharepoint.sharepoint.com/sites/App-3" -Interactive -ClientId "66a1852a-1f21-46a2-ad58-35fc4c3f1530"

Grant-PnPEntraIDAppSitePermission -AppId "c2636d73-2fbc-44df-a6fa-19a71f16afc8" -DisplayName "aa-auto" -Permissions FullControl

Get-PnpEntraIDAppSitePermission -AppId "c2636d73-2fbc-44df-a6fa-19a71f16afc8" -DisplayName "Test Workspace 2"

Revoke-PnPEntraIDAppSitePermission -PermissionId "aTowaS50fG1zLnNwLmV4dHwwZDcyY2YyNi0wZDYxLTRmMDctYWE1MC0yYWU1ZGZlZGE3YzlANzY0YjQ2ZTgtZDc5OC00ZWQzLTg3ZGItYWU1NWVkN2IwNDMy"



Connect-MgGraph

New-MgServicePrincipal 