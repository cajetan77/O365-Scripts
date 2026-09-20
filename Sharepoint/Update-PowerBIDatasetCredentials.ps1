<#
.SYNOPSIS
    Takes over a Power BI dataset and sets SharePoint datasource credentials to the calling service principal.

.DESCRIPTION
    Uses credentialDetails.useCallerAADIdentity = true so Power BI uses the same
    service principal that authenticated to the Power BI API (after TakeOver).

    Why MgGraph can work while a hand-built SharePoint accessToken fails:
      - Get-MgSite uses a Microsoft Graph token + Sites.Selected
      - Stuffing an SPO accessToken into gateway credentials is often rejected
      - Power BI's supported SP path is useCallerAADIdentity (caller = dataset owner SP)

.PARAMETER Credential
    PSCredential: UserName = Application (client) ID, Password = client secret

.EXAMPLE
    .\Update-PowerBIDatasetCredentials.ps1 -Credential $cred
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",
    [string]$WorkspaceId = "586bfe85-ac4b-4314-ac61-982723223f07",
    [string]$DatasetId = "f41ca084-844c-4483-934b-4d567d058256",
    [string]$SharePointSiteUrl = "https://caje77sharepoint.sharepoint.com/sites/M365Updates",
    [switch]$VerifyGraphAccess,
    [switch]$TakeOverOnly
)

$ErrorActionPreference = "Stop"

$ClientId = $Credential.UserName

function Test-GraphSiteAccess {
    param(
        [System.Management.Automation.PSCredential]$Credential,
        [string]$TenantId,
        [string]$SiteUrl
    )

    Import-Module Microsoft.Graph.Sites -ErrorAction Stop
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -NoWelcome
    $uri = [Uri]$SiteUrl
    $sitePathId = "{0}:{1}" -f $uri.Host, $uri.AbsolutePath.TrimEnd("/")
    Write-Host "Verifying Graph access: $sitePathId" -ForegroundColor Cyan
    try {
        $site = Get-MgSite -SiteId $sitePathId
        Write-Host "  Graph OK: $($site.DisplayName)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  Graph FAILED: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

# ---------------------------------------------------------------------------
# Optional: prove Sites.Selected works via Graph (same as your manual test)
# ---------------------------------------------------------------------------
if ($VerifyGraphAccess) {
    if (-not (Test-GraphSiteAccess -Credential $Credential -TenantId $TenantId -SiteUrl $SharePointSiteUrl)) {
        throw "Graph cannot read $SharePointSiteUrl. Grant Sites.Selected first (.\Sites-Selected.ps1)."
    }
}

# ---------------------------------------------------------------------------
# Connect Power BI as service principal
# ---------------------------------------------------------------------------
Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop

Write-Host "Connecting to Power BI as service principal..." -ForegroundColor Cyan
Write-Host "  TenantId : $TenantId" -ForegroundColor Gray
Write-Host "  ClientId : $ClientId" -ForegroundColor Gray

Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null
Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $TenantId -Credential $Credential | Out-Null
Write-Host "Connected." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Take over (required so this SP owns the cloud datasource)
# ---------------------------------------------------------------------------
$takeOverUrl = "groups/$WorkspaceId/datasets/$DatasetId/Default.TakeOver"
Write-Host "Taking over dataset $DatasetId..." -ForegroundColor Yellow
try {
    Invoke-PowerBIRestMethod -Url $takeOverUrl -Method Post -Body "{}" | Out-Null
    Write-Host "TakeOver succeeded." -ForegroundColor Green
}
catch {
    Write-Host "TakeOver note: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

if ($TakeOverOnly) {
    Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
    return
}

# ---------------------------------------------------------------------------
# List bound datasources and set credentials = caller SP identity
# ---------------------------------------------------------------------------
$boundSourcesUrl = "groups/$WorkspaceId/datasets/$DatasetId/Default.GetBoundGatewayDataSources"
Write-Host "Getting bound gateway datasources..." -ForegroundColor Cyan
$sources = Invoke-PowerBIRestMethod -Url $boundSourcesUrl -Method Get | ConvertFrom-Json

if (-not $sources.value -or $sources.value.Count -eq 0) {
    Write-Host "No bound gateway datasources found." -ForegroundColor Yellow
    Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
    return
}

# Supported Power BI path for SP: useCallerAADIdentity (no manual accessToken)
# See: https://learn.microsoft.com/en-us/rest/api/power-bi/gateways/update-datasource
$credDetails = @{
    credentialDetails = @{
        credentialType         = "OAuth2"
        useCallerAADIdentity   = $true
        encryptedConnection    = "Encrypted"
        encryptionAlgorithm    = "None"
        privacyLevel           = "Organizational"
    }
} | ConvertTo-Json -Depth 5

$anyOk = $false
foreach ($source in $sources.value) {
    $gatewayId = $source.gatewayId
    $datasourceId = $source.id
    Write-Host "Updating datasource $datasourceId (useCallerAADIdentity)..." -ForegroundColor Yellow
    Write-Host "  Gateway : $gatewayId" -ForegroundColor Gray
    if ($source.datasourceType) {
        Write-Host "  Type    : $($source.datasourceType)" -ForegroundColor Gray
    }
    if ($source.connectionDetails) {
        Write-Host "  Conn    : $($source.connectionDetails)" -ForegroundColor Gray
    }

    $patchUrl = "gateways/$gatewayId/datasources/$datasourceId"
    try {
        Invoke-PowerBIRestMethod -Url $patchUrl -Method Patch -Body $credDetails | Out-Null
        Write-Host "  Updated." -ForegroundColor Green
        $anyOk = $true
    }
    catch {
        $detail = $_.ErrorDetails.Message
        if (-not $detail) { $detail = $_.Exception.Message }
        Write-Host "  Failed: $detail" -ForegroundColor Red

        # SharePoint Folder / Web connectors often do NOT support SP — only SharePoint List does
        Write-Host "  Note: SharePoint List connector supports SP; SharePoint Folder/Web usually require user OAuth." -ForegroundColor DarkYellow
    }
}

Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue

if (-not $anyOk) {
    throw @"
Credential update failed.

Your MgGraph test proves Graph Sites.Selected works — that is expected.
Power BI must use useCallerAADIdentity (this script), not a pasted SharePoint accessToken.

If it still fails:
  1. Dataset must use SharePoint List (not Folder/Web/Excel-from-library via Web)
  2. SP must own the dataset (TakeOver — already attempted)
  3. App needs site access (Sites.Selected grant) — which you already verified via Graph
"@
}

Write-Host "Done. Try a dataset refresh next." -ForegroundColor Green
