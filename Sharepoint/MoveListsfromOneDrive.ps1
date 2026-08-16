<#
.SYNOPSIS
    Moves (or copies) all items from a SharePoint/OneDrive list to a list on another site.

.DESCRIPTION
    Reads all items from -ListTitle on -SourceSiteUrl and creates matching items on
    -DestinationSiteUrl / -DestinationListTitle. User-editable field values are copied
    (attachments are not). System fields (ID, Created, Modified, Author, etc.) are skipped.

    If the destination list does not exist, it is created by copying the source list
    definition (Copy-PnPList), then items are verified / copied as needed.

.PARAMETER DeleteSource
    After a successful copy, delete the item from the source list (true move).

.PARAMETER WhatIf
    Preview only; do not create or delete items. Default: $false

.EXAMPLE
    .\MoveListsfromOneDrive.ps1

.EXAMPLE
    .\MoveListsfromOneDrive.ps1 -DeleteSource
#>

[CmdletBinding()]
param(
    [string]$SourceSiteUrl = "https://caje77sharepoint-my.sharepoint.com/personal/caje77_keiratheapp_com",
    [string]$DestinationSiteUrl = "https://caje77sharepoint.sharepoint.com/sites/CajIntra",
    [string]$ListTitle = "TestCajOneDrive",
    [string]$DestinationListTitle = "TestCajOneDrive",
    [switch]$DeleteSource,
    [switch]$WhatIf,
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

$skipFieldNames = @(
    "ID", "ContentType", "ContentTypeId", "UniqueId", "GUID", "Edit",
    "LinkTitle", "LinkTitleNoMenu", "DocIcon", "ServerUrl", "EncodedAbsUrl",
    "BaseName", "FileRef", "FileDirRef", "FileLeafRef", "FSObjType",
    "PermMask", "Attachments", "MetaInfo", "ScopeId", "SyncClientId",
    "ParentUniqueId", "Last_x0020_Modified", "Created_x0020_Date",
    "Author", "Editor", "Created", "Modified",
    "_UIVersionString", "_ModerationStatus", "_ModerationComments",
    "_HasCopyDestinations", "_CopySource", "_IsCurrentVersion",
    "owshiddenversion", "WorkflowVersion", "WorkflowInstanceID",
    "InstanceID", "Order", "Title" # Title added back below if present as writable
) | ForEach-Object { $_ }

# Title should be copied — remove from skip
$skipFieldNames = $skipFieldNames | Where-Object { $_ -ne "Title" }

function Connect-Site {
    param([string]$Url)
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
}

function Get-WritableListFields {
    param([string]$ListName)

    $fields = Get-PnPField -List $ListName
    $writable = foreach ($f in $fields) {
        if ($f.ReadOnlyField) { continue }
        if ($f.Hidden) { continue }
        if ($f.InternalName -like "_*") { continue }
        if ($skipFieldNames -contains $f.InternalName) { continue }
        if ($f.FieldTypeKind -eq "Computed") { continue }
        if ($f.FieldTypeKind -eq "File") { continue }
        $f
    }
    return @($writable)
}

function ConvertTo-PnPFieldValue {
    param(
        $Field,
        $RawValue
    )

    if ($null -eq $RawValue -or $RawValue -eq "") { return $null }

    switch ($Field.TypeAsString) {
        { $_ -in @("User", "UserMulti") } {
            $isEnumerable = ($RawValue -is [System.Collections.IEnumerable]) -and ($RawValue -isnot [string])
            if ($isEnumerable) {
                $emails = @()
                foreach ($u in @($RawValue)) {
                    if ($u.Email) { $emails += $u.Email }
                    elseif ($u.LookupValue) { $emails += $u.LookupValue }
                    elseif ($u -is [string]) { $emails += $u }
                }
                if ($emails.Count -eq 0) { return $null }
                return ($emails -join ";")
            }
            if ($RawValue.Email) { return $RawValue.Email }
            if ($RawValue.LookupValue) { return $RawValue.LookupValue }
            return [string]$RawValue
        }
        { $_ -in @("Lookup", "LookupMulti") } {
            Write-Verbose "Skipping lookup field '$($Field.InternalName)' (cross-site)."
            return $null
        }
        "URL" {
            if ($RawValue.Url) {
                $desc = if ($RawValue.Description) { $RawValue.Description } else { $RawValue.Url }
                return "{0}, {1}" -f $RawValue.Url, $desc
            }
            return [string]$RawValue
        }
        { $_ -in @("TaxonomyFieldType", "TaxonomyFieldTypeMulti") } {
            Write-Verbose "Skipping taxonomy field '$($Field.InternalName)'."
            return $null
        }
        "MultiChoice" {
            if ($RawValue -is [System.Array]) { return @($RawValue) }
            return $RawValue
        }
        "DateTime" {
            return [DateTime]$RawValue
        }
        "Boolean" {
            return [bool]$RawValue
        }
        "Number" {
            return [double]$RawValue
        }
        "Currency" {
            return [double]$RawValue
        }
        default {
            return $RawValue
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host " Move list items between sites" -ForegroundColor Magenta
Write-Host " Source : $SourceSiteUrl" -ForegroundColor White
Write-Host " Dest   : $DestinationSiteUrl" -ForegroundColor White
Write-Host " List   : $ListTitle -> $DestinationListTitle" -ForegroundColor White
Write-Host " Mode   : $(if ($WhatIf) { 'WHAT-IF' } elseif ($DeleteSource) { 'MOVE (delete source)' } else { 'COPY' })" -ForegroundColor White
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host ""

# --- Source: load list, fields, items ---
Write-Host "Connecting to source..." -ForegroundColor Yellow
Connect-Site -Url $SourceSiteUrl

$sourceList = Get-PnPList -Identity $ListTitle -ErrorAction Stop
Write-Host "Source list '$($sourceList.Title)' — $($sourceList.ItemCount) item(s), template $($sourceList.BaseTemplate)" -ForegroundColor Green

$writableFields = Get-WritableListFields -ListName $ListTitle
Write-Host "Writable fields to copy: $($writableFields.InternalName -join ', ')" -ForegroundColor Gray

$sourceItems = @(Get-PnPListItem -List $ListTitle -PageSize 500 | ForEach-Object { $_ })
Write-Host "Loaded $($sourceItems.Count) source item(s)." -ForegroundColor Green

if ($sourceItems.Count -eq 0) {
    Write-Host "Nothing to move." -ForegroundColor Yellow
    return
}

# Keep field values while still on source connection
$payload = foreach ($item in $sourceItems) {
    $values = @{}
    foreach ($field in $writableFields) {
        $name = $field.InternalName
        if (-not $item.FieldValues.ContainsKey($name)) { continue }
        $converted = ConvertTo-PnPFieldValue -Field $field -RawValue $item.FieldValues[$name]
        if ($null -ne $converted -and $converted -ne "") {
            $values[$name] = $converted
        }
    }

    [PSCustomObject]@{
        SourceId = $item.Id
        Values   = $values
    }
}

# --- Destination: ensure list exists ---
Write-Host "Connecting to destination..." -ForegroundColor Yellow
Connect-Site -Url $DestinationSiteUrl

$destList = Get-PnPList -Identity $DestinationListTitle -ErrorAction SilentlyContinue
if (-not $destList) {
    Write-Host "Destination list '$DestinationListTitle' not found. Creating via Copy-PnPList..." -ForegroundColor Yellow
    if ($WhatIf) {
        Write-Host "[WhatIf] Would copy list definition to destination." -ForegroundColor DarkCyan
    }
    else {
        # Copy-PnPList must run against the source connection
        Connect-Site -Url $SourceSiteUrl
        Copy-PnPList -Identity $ListTitle -DestinationWebUrl $DestinationSiteUrl -Title $DestinationListTitle
        Connect-Site -Url $DestinationSiteUrl
        $destList = Get-PnPList -Identity $DestinationListTitle -ErrorAction Stop
        Write-Host "Created destination list '$($destList.Title)'." -ForegroundColor Green

        # Copy-PnPList may already have copied items; if counts match, optional delete-only path
        if ($destList.ItemCount -ge $sourceItems.Count) {
            Write-Host "Destination already has $($destList.ItemCount) item(s) after list copy." -ForegroundColor Green
            if ($DeleteSource -and -not $WhatIf) {
                Connect-Site -Url $SourceSiteUrl
                Write-Host "Deleting source items..." -ForegroundColor Yellow
                $batch = New-PnPBatch
                foreach ($item in $sourceItems) {
                    Remove-PnPListItem -List $ListTitle -Identity $item.Id -Recycle -Batch $batch
                }
                Invoke-PnPBatch -Batch $batch
                Write-Host "Source items deleted (recycle bin)." -ForegroundColor Green
            }
            Write-Host "Done." -ForegroundColor Magenta
            return
        }
    }
}
else {
    Write-Host "Destination list '$($destList.Title)' exists — $($destList.ItemCount) item(s)." -ForegroundColor Green
}

# --- Copy items ---
Connect-Site -Url $DestinationSiteUrl
$created = 0
$failed = 0

foreach ($row in $payload) {
    $label = if ($row.Values.ContainsKey("Title") -and $row.Values["Title"]) { $row.Values["Title"] } else { "ID $($row.SourceId)" }
    Write-Host "  Item $label..." -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "    [WhatIf] Would create item with fields: $($row.Values.Keys -join ', ')" -ForegroundColor DarkCyan
        $created++
        continue
    }

    try {
        if ($row.Values.Count -eq 0) {
            Add-PnPListItem -List $DestinationListTitle | Out-Null
        }
        else {
            Add-PnPListItem -List $DestinationListTitle -Values $row.Values | Out-Null
        }
        $created++
    }
    catch {
        $failed++
        Write-Warning "  Failed to create item '$label': $_"
    }
}

# --- Delete source (move) ---
if ($DeleteSource -and -not $WhatIf -and $created -gt 0 -and $failed -eq 0) {
    Write-Host "Deleting $($payload.Count) source item(s)..." -ForegroundColor Yellow
    Connect-Site -Url $SourceSiteUrl
    $batch = New-PnPBatch
    foreach ($row in $payload) {
        Remove-PnPListItem -List $ListTitle -Identity $row.SourceId -Recycle -Batch $batch
    }
    Invoke-PnPBatch -Batch $batch
    Write-Host "Source items moved to recycle bin." -ForegroundColor Green
}
elseif ($DeleteSource -and $failed -gt 0) {
    Write-Warning "Skipped source delete because $failed item(s) failed to copy."
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host " Created : $created" -ForegroundColor Green
Write-Host " Failed  : $failed" -ForegroundColor $(if ($failed) { "Red" } else { "Gray" })
Write-Host "=============================================" -ForegroundColor Magenta

Disconnect-PnPOnline -ErrorAction SilentlyContinue
