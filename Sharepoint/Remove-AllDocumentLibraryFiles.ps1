<#
.SYNOPSIS
    Deletes all files (and folders) from Document Libraries across SharePoint and/or OneDrive sites.

.DESCRIPTION
    Connects to the SharePoint Admin Center using credentials from config.json, enumerates
    tenant sites (SharePoint and optionally OneDrive), then empties each Document Library
    (BaseTemplate 101), excluding known system libraries.

    This is a destructive operation. Run with -WhatIf first to preview impact.
    Actual deletion requires -Force.

.PARAMETER IncludeSharePoint
    Process SharePoint team/communication/hub sites. Default: $true

.PARAMETER IncludeOneDrive
    Process OneDrive for Business personal sites. Default: $true

.PARAMETER EmptyRecycleBin
    Also empty each site's recycle bin after deletion. Default: $false

.PARAMETER LibraryTitle
    Optional. Limit deletion to a single library title (e.g. "Documents").
    When omitted, all non-system Document Libraries are processed.

.PARAMETER SiteUrl
    Optional. Process only this site URL instead of the whole tenant.

.PARAMETER WhatIf
    Preview which sites/libraries/items would be deleted without making changes. Default: $true

.PARAMETER Force
    Required to perform real deletions. Ignored when -WhatIf is also set.

.PARAMETER ConfigPath
    Path to config.json. Defaults to parent folder of this script.

.EXAMPLE
    # Preview only (safe default)
    .\Remove-AllDocumentLibraryFiles.ps1

.EXAMPLE
    # Preview OneDrive only
    .\Remove-AllDocumentLibraryFiles.ps1 -IncludeSharePoint:$false -IncludeOneDrive

.EXAMPLE
    # Actually delete files in all Document Libraries (SharePoint + OneDrive)
    .\Remove-AllDocumentLibraryFiles.ps1 -WhatIf:$false -Force

.EXAMPLE
    # Delete only the Documents library on one site, and empty recycle bin
    .\Remove-AllDocumentLibraryFiles.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
        -LibraryTitle "Documents" -WhatIf:$false -Force -EmptyRecycleBin
#>

[CmdletBinding()]
param(
    [switch]$IncludeSharePoint = $true,
    [switch]$IncludeOneDrive = $true,
    [switch]$EmptyRecycleBin,
    [string]$LibraryTitle,
    [string]$SiteUrl,
    # Custom WhatIf (not SupportsShouldProcess) so we can default to preview-only
    [switch]$WhatIf = $true,
    [switch]$Force,
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -Force

# ---------------------------------------------------------------------------
# Config / auth (matches other scripts in this repo)
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
$TenantName = $config.TenantName
$Thumbprint = $config.ThumbPrint
#$clientSecret = $config.ClientSecret
$adminUrl = "https://$TenantName-admin.sharepoint.com"

$doDelete = (-not $WhatIf) -and $Force
if (-not $WhatIf -and -not $Force) {
    throw "Refusing to delete without -Force. Preview with default -WhatIf, or run: -WhatIf:`$false -Force"
}

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = Join-Path $PSScriptRoot "Remove-DocumentLibraryFiles_$timestamp.csv"
$results = [System.Collections.Generic.List[object]]::new()

# System libraries that should never be wiped
$excludedLibraryTitles = @(
    "Form Templates",
    "Style Library",
    "Site Assets",
    "Site Pages",
    "Preservation Hold Library",
    "PreservationHoldLibrary",
    "Apps for SharePoint",
    "App Packages",
    "Content Organizer Rules",
    "Converted Forms",
    "Drop Off Library",
    "Master Page Gallery",
    "Solution Gallery",
    "Theme Gallery",
    "Web Part Gallery",
    "wfpub",
    "Pages"
)

function Connect-TenantAdmin {
    Write-Host "Connecting to Admin Center: $adminUrl" -ForegroundColor Yellow
    try {
        if ($clientSecret) {
            Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -ClientSecret $clientSecret
        }
        elseif ($Thumbprint) {
            Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
        }
        else {
            Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $ClientId
        }
        Write-Host "Connected to Admin Center." -ForegroundColor Green
    }
    catch {
        throw "Failed to connect to SharePoint Admin Center: $_"
    }
}

function Connect-Site {
    param([string]$Url)

    if ($clientSecret) {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $clientSecret
    }
    elseif ($Thumbprint) {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
    }
    else {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
    }
}

function Get-TargetSites {
    if ($SiteUrl) {
        Write-Host "Single-site mode: $SiteUrl" -ForegroundColor Cyan
        return @(Get-PnPTenantSite -Url $SiteUrl)
    }

    Write-Host "Enumerating tenant sites (IncludeOneDriveSites=$IncludeOneDrive)..." -ForegroundColor Cyan
    $allSites = Get-PnPTenantSite -IncludeOneDriveSites:$IncludeOneDrive -Detailed

    $targets = @()
    foreach ($site in $allSites) {
        $isOneDrive = ($site.Template -like "SPSPERS*") -or ($site.Url -match "-my\.sharepoint\.com/personal/")

        if ($isOneDrive -and -not $IncludeOneDrive) { continue }
        if (-not $isOneDrive -and -not $IncludeSharePoint) { continue }

        # Skip admin / app catalog / search / redirect sites
        if ($site.Url -match "-admin\.sharepoint\.com") { continue }
        if ($site.Template -in @("RedirectSite#0", "POINTPUBLISHINGTOPIC#0", "POINTPUBLISHINGHUB#0")) { continue }
        if ($site.LockState -and $site.LockState -ne "Unlock") {
            Write-Host "  Skipping locked site: $($site.Url) [$($site.LockState)]" -ForegroundColor DarkYellow
            continue
        }

        $targets += $site
    }

    return $targets
}

function Clear-DocumentLibrary {
    param(
        [string]$ListTitle,
        [string]$SiteType
    )

    $list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (-not $list) {
        return [PSCustomObject]@{
            SiteUrl      = (Get-PnPWeb).Url
            SiteType     = $SiteType
            Library      = $ListTitle
            ItemCount    = 0
            DeletedCount = 0
            Status       = "NotFound"
            Message      = "Library not found"
        }
    }

    $itemCount = $list.ItemCount
    Write-Host "    Library '$ListTitle' — $itemCount item(s)" -ForegroundColor White

    if ($itemCount -eq 0) {
        return [PSCustomObject]@{
            SiteUrl      = (Get-PnPWeb).Url
            SiteType     = $SiteType
            Library      = $ListTitle
            ItemCount    = 0
            DeletedCount = 0
            Status       = "Empty"
            Message      = "Already empty"
        }
    }

    if (-not $doDelete) {
        Write-Host "    [WhatIf] Would delete $itemCount item(s) from '$ListTitle'" -ForegroundColor DarkCyan
        return [PSCustomObject]@{
            SiteUrl      = (Get-PnPWeb).Url
            SiteType     = $SiteType
            Library      = $ListTitle
            ItemCount    = $itemCount
            DeletedCount = 0
            Status       = "WhatIf"
            Message      = "Would delete $itemCount items"
        }
    }

    $deleted = 0
    $errors = 0
    $pageSize = 100

    try {
        # --- Fast path: delete each top-level folder as one tree ---
        $rootFolder = Get-PnPFolder -ListRootFolder $ListTitle
        $topFolders = @(
            Get-PnPFolderItem -Identity $rootFolder -ItemType Folder |
                Where-Object { $_.Name -notin @("Forms", "Item", "Attachments") }
        )

        foreach ($folder in $topFolders) {
            Write-Host "      Removing folder tree '$($folder.Name)'..." -ForegroundColor Gray
            Remove-PnPFolder -Name $folder.Name -Folder $rootFolder -Recycle -Force
            $deleted++
        }

        # --- Sweep remaining items (root files + anything left) in true multi-item pages ---
        # Always re-query the first page (no continuation token) so deletes don't skip rows.
        do {
            $caml = @"
<View Scope='RecursiveAll'>
  <Query>
    <OrderBy>
      <FieldRef Name='FileRef' Ascending='FALSE' />
    </OrderBy>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' />
    <FieldRef Name='FileRef' />
  </ViewFields>
  <RowLimit>$pageSize</RowLimit>
</View>
"@
            # Pipe through ForEach-Object so a ListItemCollection is enumerated, not wrapped as 1 object
            $items = @(Get-PnPListItem -List $ListTitle -Query $caml | ForEach-Object { $_ })
            if ($items.Count -eq 0) { break }

            $batch = New-PnPBatch
            foreach ($item in $items) {
                # -Force cannot be combined with -Batch in this PnP version
                Remove-PnPListItem -List $ListTitle -Identity $item.Id -Recycle -Batch $batch
            }
            Invoke-PnPBatch -Batch $batch

            $deleted += $items.Count
            Write-Host "      Deleted $($items.Count) item(s) (running total: $deleted)..." -ForegroundColor Gray
        } while ($true)

        return [PSCustomObject]@{
            SiteUrl      = (Get-PnPWeb).Url
            SiteType     = $SiteType
            Library      = $ListTitle
            ItemCount    = $itemCount
            DeletedCount = $deleted
            Status       = "Deleted"
            Message      = "Deleted $deleted items/folders (sent to recycle bin)"
        }
    }
    catch {
        $errors++
        Write-Host "      ERROR clearing '$ListTitle': $_" -ForegroundColor Red
        return [PSCustomObject]@{
            SiteUrl      = (Get-PnPWeb).Url
            SiteType     = $SiteType
            Library      = $ListTitle
            ItemCount    = $itemCount
            DeletedCount = $deleted
            Status       = "Error"
            Message      = $_.Exception.Message
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host " Remove Document Library Files" -ForegroundColor Magenta
Write-Host " Mode: $(if ($doDelete) { 'DELETE' } else { 'WHAT-IF (no changes)' })" -ForegroundColor Magenta
Write-Host " SharePoint: $IncludeSharePoint | OneDrive: $IncludeOneDrive" -ForegroundColor Magenta
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host ""

Connect-TenantAdmin
$sites = @(Get-TargetSites)
Write-Host "Sites to process: $($sites.Count)" -ForegroundColor Green

foreach ($site in $sites) {
    $isOneDrive = ($site.Template -like "SPSPERS*") -or ($site.Url -match "-my\.sharepoint\.com/personal/")
    $siteType = if ($isOneDrive) { "OneDrive" } else { "SharePoint" }

    Write-Host ""
    Write-Host "[$siteType] $($site.Url)" -ForegroundColor Cyan

    try {
        Connect-Site -Url $site.Url

        # Document libraries only (BaseTemplate 101), exclude system libs
        $libraries = Get-PnPList | Where-Object {
            $_.BaseTemplate -eq 101 -and
            $_.Hidden -eq $false -and
            ($excludedLibraryTitles -notcontains $_.Title)
        }

        if ($LibraryTitle) {
            $libraries = $libraries | Where-Object { $_.Title -eq $LibraryTitle }
        }

        if (-not $libraries -or $libraries.Count -eq 0) {
            Write-Host "  No matching Document Libraries." -ForegroundColor DarkYellow
            $results.Add([PSCustomObject]@{
                    SiteUrl      = $site.Url
                    SiteType     = $siteType
                    Library      = ""
                    ItemCount    = 0
                    DeletedCount = 0
                    Status       = "Skipped"
                    Message      = "No matching document libraries"
                })
            continue
        }

        foreach ($lib in $libraries) {
            $result = Clear-DocumentLibrary -ListTitle $lib.Title -SiteType $siteType
            $results.Add($result)
        }

        if ($EmptyRecycleBin -and $doDelete) {
            Write-Host "  Emptying recycle bin..." -ForegroundColor Yellow
            try {
                Clear-PnPRecycleBinItem -Force -ErrorAction Stop
                Write-Host "  Recycle bin emptied." -ForegroundColor Green
            }
            catch {
                Write-Host "  Warning: could not empty recycle bin: $_" -ForegroundColor DarkYellow
            }
        }
        elseif ($EmptyRecycleBin -and -not $doDelete) {
            Write-Host "  [WhatIf] Would empty site recycle bin" -ForegroundColor DarkCyan
        }
    }
    catch {
        Write-Host "  ERROR on site: $_" -ForegroundColor Red
        $results.Add([PSCustomObject]@{
                SiteUrl      = $site.Url
                SiteType     = $siteType
                Library      = ""
                ItemCount    = 0
                DeletedCount = 0
                Status       = "SiteError"
                Message      = $_.Exception.Message
            })
    }
}

# ---------------------------------------------------------------------------
# Summary / log
# ---------------------------------------------------------------------------
$results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

$deletedLibs = @($results | Where-Object { $_.Status -eq "Deleted" }).Count
$whatIfLibs = @($results | Where-Object { $_.Status -eq "WhatIf" }).Count
$errorCount = @($results | Where-Object { $_.Status -in @("Error", "SiteError") }).Count
$totalItems = ($results | Measure-Object -Property ItemCount -Sum).Sum

Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host " Done" -ForegroundColor Magenta
Write-Host " Sites processed : $($sites.Count)" -ForegroundColor White
Write-Host " Libraries w/ items: $totalItems item(s) across results" -ForegroundColor White
Write-Host " Deleted libraries : $deletedLibs" -ForegroundColor $(if ($doDelete) { "Green" } else { "Gray" })
Write-Host " WhatIf libraries  : $whatIfLibs" -ForegroundColor Cyan
Write-Host " Errors            : $errorCount" -ForegroundColor $(if ($errorCount) { "Red" } else { "Gray" })
Write-Host " Log               : $logPath" -ForegroundColor Yellow
Write-Host "=============================================" -ForegroundColor Magenta

if (-not $doDelete) {
    Write-Host ""
    Write-Host "No files were deleted. To execute for real, run:" -ForegroundColor Yellow
    Write-Host "  .\Remove-AllDocumentLibraryFiles.ps1 -WhatIf:`$false -Force" -ForegroundColor White
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
