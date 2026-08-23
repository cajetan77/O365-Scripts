[CmdletBinding()]
param(
    [string]$SourceSiteUrl = "https://caje77sharepoint.sharepoint.com/sites/CajIntra",
    [string]$ListTitle = "Test1",
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -Force

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
if (-not $ConfigPath) {
    $ConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"
    if (-not (Test-Path $ConfigPath)) {
        $ConfigPath = Join-Path $PSScriptRoot "..\config.json"
    }
    if (-not (Test-Path $ConfigPath)) {
        $ConfigPath = ".\config.json"
    }
}

if (-not (Test-Path $ConfigPath)) {
    throw "config.json not found. Pass -ConfigPath or place config.json next to / above this script."
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$Thumbprint = $config.ThumbPrint


function Connect_Graph {
    Write-Host "Connecting to Graph..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    }
    catch {
        Write-Host "Error connecting to Graph: $_" -ForegroundColor Red
        exit 1
    }
}

Connect_Graph

$groups = Get-MgGroup | Where-Object { $_.DisplayName -like "Caj*" } 

$groupsTest = Get-MgGroup | Where-Object { $_.DisplayName -like "Test*" } 

$allGroups = @( $groups + $groupsTest )
Write-Host "Found $($allGroups.Count) groups" -ForegroundColor Green
try {
    Connect-PnPOnline -Url $SourceSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId
}
catch {
    Write-Host "Error connecting to SharePoint: $_" -ForegroundColor Red
    exit 1
}
$list = Get-PnPList -Identity $ListTitle

$fields = Get-PnPField -List $list | Where-Object { $_.InternalName -eq "Group" }

$choices = [string[]]@(
    $allGroups |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.DisplayName) } |
    Select-Object -ExpandProperty DisplayName -Unique |
    Sort-Object
)

try {
    Set-PnPField -List $list -Identity $fields.InternalName -Values @{ Choices = $choices }
}
catch {
    Write-Host "Error updating group choices: $_" -ForegroundColor Red
    exit 1
}
