<#
.SYNOPSIS
    Refresh a Power BI dataset using a service principal (MicrosoftPowerBIMgmt).

.NOTES
    Based on: https://www.youtube.com/watch?v=XnNppwGbZBE

    If POST refresh returns ItemNotFound while Get-PowerBIDataset succeeds:
      - Add the APP (AIALSharepoint) directly as Admin on the workspace
      - Ensure tenant setting allows group PowerBIApp for Power BI APIs
      - Remove Power BI *Application* permissions from the app registration
        (they conflict with service-principal auth; Delegated only is OK)
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$WorkspaceId = "586bfe85-ac4b-4314-ac61-982723223f07",
    [string]$DatasetId = "f41ca084-844c-4483-934b-4d567d058256",
    [string]$DatasetName = "GuestAccounts",
    [string]$ConfigPath,
    [switch]$DiagnoseOnly
)

$ErrorActionPreference = "Stop"

if (-not $ConfigPath) {
    $ConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"
    if (-not (Test-Path $ConfigPath)) {
        $ConfigPath = Join-Path $PSScriptRoot "..\config.json"
    }
}

if ((-not $TenantId -or -not $ClientId -or -not $ClientSecret) -and (Test-Path $ConfigPath)) {
    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
    if (-not $TenantId) { $TenantId = $config.TenantId }
    if (-not $ClientId) { $ClientId = $config.AppId }
    if (-not $ClientSecret) { $ClientSecret = $config.ClientSecret }
}

if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
    throw "Set TenantId, ClientId, and ClientSecret as parameters or in config.json."
}

# ---------------------------------------------------------------------------
# Module
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Write-Host "Installing MicrosoftPowerBIMgmt..." -ForegroundColor Yellow
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}
Import-Module MicrosoftPowerBIMgmt -Force

# ---------------------------------------------------------------------------
# Connect (same as video)
# ---------------------------------------------------------------------------
Write-Host "Connecting to Power BI as service principal..." -ForegroundColor Cyan
Write-Host "  TenantId : $TenantId" -ForegroundColor Gray
Write-Host "  ClientId : $ClientId" -ForegroundColor Gray

$secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ($ClientId, $secureSecret)

Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId $TenantId | Out-Null
Write-Host "Connected." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Resolve workspace + dataset from live inventory (do not trust hardcoded GUIDs alone)
# ---------------------------------------------------------------------------
Write-Host "Listing workspaces..." -ForegroundColor Cyan
$workspaces = @(Get-PowerBIWorkspace -ErrorAction SilentlyContinue)
if ($workspaces.Count -eq 0) {
    $workspaces = @(Get-PowerBIWorkspace -Scope Organization -All -ErrorAction SilentlyContinue)
}
$workspaces | ForEach-Object { Write-Host "  - $($_.Name) [$($_.Id)]" -ForegroundColor DarkGray }

$ws = $workspaces | Where-Object { $_.Id.Guid -eq $WorkspaceId -or $_.Id.ToString() -eq $WorkspaceId -or $_.Name -eq "M365 Reporting" } |
    Select-Object -First 1
if (-not $ws) {
    throw "Workspace not found / not visible to this service principal."
}

# Always use Id returned by the API
$WorkspaceId = $ws.Id.ToString()
Write-Host "Workspace OK: $($ws.Name) [$WorkspaceId]" -ForegroundColor Green

Write-Host "Listing datasets..." -ForegroundColor Cyan
$datasets = @(Get-PowerBIDataset -WorkspaceId $WorkspaceId)
$datasets | ForEach-Object {
    $refreshable = $null
    if ($_.PSObject.Properties.Name -contains "IsRefreshable") { $refreshable = $_.IsRefreshable }
    Write-Host "  - $($_.Name) [$($_.Id)] IsRefreshable=$refreshable" -ForegroundColor DarkGray
}

$ds = $datasets | Where-Object { $_.Name -eq $DatasetName } | Select-Object -First 1
if (-not $ds) {
    $ds = $datasets | Where-Object { $_.Id.ToString() -eq $DatasetId } | Select-Object -First 1
}
if (-not $ds) {
    throw "Dataset '$DatasetName' not found in workspace '$($ws.Name)'."
}

$DatasetId = $ds.Id.ToString()
Write-Host "Dataset OK: $($ds.Name) [$DatasetId]" -ForegroundColor Green

if ($DiagnoseOnly) {
    Write-Host "DiagnoseOnly: done." -ForegroundColor Cyan
    Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
    return
}

# ---------------------------------------------------------------------------
# Refresh helpers
# ---------------------------------------------------------------------------
function Get-ErrorBody {
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }
    return $ErrorRecord.Exception.Message
}

$relativeUrl = "groups/$WorkspaceId/datasets/$DatasetId/refreshes"
$absoluteUrl = "https://api.powerbi.com/v1.0/myorg/$relativeUrl"
$fabricUrl = "https://api.fabric.microsoft.com/v1/workspaces/$WorkspaceId/items/$DatasetId/jobs/instances?jobType=Refresh"

Write-Host "Triggering refresh..." -ForegroundColor Yellow

$attempts = @()

# 1) Module-native REST (preferred — uses the SP session from Connect-PowerBIServiceAccount)
$attempts += {
    Write-Host "  Attempt: Invoke-PowerBIRestMethod + NoNotification" -ForegroundColor Gray
    $body = @{ notifyOption = "NoNotification" } | ConvertTo-Json -Compress
    Invoke-PowerBIRestMethod -Url $relativeUrl -Method Post -Body $body | Out-Null
}

# 2) Module-native, no body
$attempts += {
    Write-Host "  Attempt: Invoke-PowerBIRestMethod (no body)" -ForegroundColor Gray
    Invoke-PowerBIRestMethod -Url $relativeUrl -Method Post | Out-Null
}

# 3) Raw REST with Get-PowerBIAccessToken (video style)
$attempts += {
    Write-Host "  Attempt: Invoke-RestMethod + Get-PowerBIAccessToken" -ForegroundColor Gray
    $headers = Get-PowerBIAccessToken
    $headers["Content-Type"] = "application/json"
    Invoke-RestMethod -Method Post -Uri $absoluteUrl -Headers $headers -Body '{"notifyOption":"NoNotification"}' | Out-Null
}

# 4) Fabric semantic-model job (works for some Fabric/trial workspaces)
$attempts += {
    Write-Host "  Attempt: Fabric jobType=Refresh" -ForegroundColor Gray
    $headers = Get-PowerBIAccessToken
    # Fabric needs a Fabric token audience; try with Power BI token first, then AAD Fabric scope
    try {
        Invoke-RestMethod -Method Post -Uri $fabricUrl -Headers $headers | Out-Null
    }
    catch {
        $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $tokenBody = @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://api.fabric.microsoft.com/.default"
        }
        $tok = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
        $fabricHeaders = @{
            Authorization  = "Bearer $($tok.access_token)"
            "Content-Type" = "application/json"
        }
        Invoke-RestMethod -Method Post -Uri $fabricUrl -Headers $fabricHeaders | Out-Null
    }
}

$succeeded = $false
$lastError = $null

foreach ($attempt in $attempts) {
    try {
        & $attempt
        Write-Host "Refresh request accepted." -ForegroundColor Green
        $succeeded = $true
        break
    }
    catch {
        $lastError = Get-ErrorBody -ErrorRecord $_
        Write-Host "    Failed: $lastError" -ForegroundColor Red
    }
}

if (-not $succeeded) {
    # Extra diagnostic: can we read refresh history?
    Write-Host ""
    Write-Host "Checking GET refresh history (permission probe)..." -ForegroundColor Cyan
    try {
        $hist = Invoke-PowerBIRestMethod -Url "$relativeUrl?`$top=1" -Method Get
        Write-Host "  GET history works. Response: $hist" -ForegroundColor Yellow
        Write-Host "  So the dataset is visible but POST refresh is blocked." -ForegroundColor Yellow
    }
    catch {
        Write-Host "  GET history also failed: $(Get-ErrorBody $_)" -ForegroundColor Red
        Write-Host "  Service principal likely lacks write/refresh rights on this dataset." -ForegroundColor Yellow
    }

    Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
    throw @"
All refresh attempts failed.

Most likely fix for your tenant (matches the video + your screenshots):
  1. Workspace 'M365 Reporting' → Manage access → add 'AIALSharepoint' (the app) as Admin DIRECTLY
  2. Fabric Admin → Tenant settings → Developer settings →
     Allow service principals to use Power BI APIs → Enabled for 'PowerBIApp'
  3. App registration → API permissions → remove Power BI *Application* permissions
     (keep only Delegated Dataset.ReadWrite.All if you follow the video)
  4. Wait 15–30 minutes, then rerun

Last error: $lastError
"@
}

# History
Start-Sleep -Seconds 2
try {
    $historyJson = Invoke-PowerBIRestMethod -Url "$relativeUrl?`$top=1" -Method Get
    $history = $historyJson | ConvertFrom-Json
    if ($history.value) {
        $latest = $history.value[0]
        Write-Host "Latest refresh: $($latest.status) (start: $($latest.startTime))" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Refresh was accepted; could not read history yet: $(Get-ErrorBody $_)" -ForegroundColor DarkYellow
}

Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
