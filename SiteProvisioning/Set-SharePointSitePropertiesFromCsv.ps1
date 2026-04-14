param(
    [string]$CsvPath = ".\Sites.csv",
    [string]$ConfigPath = ".\config.json",
    [string]$HubSiteUrl = "https://caje77sharepoint.sharepoint.com/sites/AIALIntranet"
)

Write-Host "Starting SharePoint site property update from CSV..." -ForegroundColor Green

if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: Config file not found at $ConfigPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at $CsvPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$tenantId = $config.TenantId
$clientId = if ($config.AppId) { $config.AppId } else { $config.AppId }
$tenantName = $config.TenantName
$thumbprint = $config.ThumbPrint

if ([string]::IsNullOrWhiteSpace($tenantId) -or
    [string]::IsNullOrWhiteSpace($clientId) -or
    [string]::IsNullOrWhiteSpace($tenantName) -or
    [string]::IsNullOrWhiteSpace($thumbprint)) {
    Write-Host "ERROR: Missing TenantId / AppId / TenantName / ThumbPrint in config.json" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($HubSiteUrl)) {
    Write-Host "ERROR: HubSiteUrl is required. Example: -HubSiteUrl 'https://contoso.sharepoint.com/sites/hub'" -ForegroundColor Red
    exit 1
}

$HubSiteUrl = $HubSiteUrl.ToString().Trim()



function Set-SiteRegionalSettings {
    param(
        [string]$SiteUrl
    )
 
    try {
        $web = Get-PnpWeb -Includes RegionalSettings.LocaleId, RegionalSettings.TimeZones -Connection $siteconnection
        $localeId = 5129 # New Zealand English
        $web.RegionalSettings.LocaleId = $localeId
        $web.Update()
        Invoke-PnPQuery
        Write-Host  "Updated Site Regional Settings to have NZ Time Zone and NZ Locale  $($web.Url)" "Info"
    }
    catch {
        Write-Host "Error connecting to site $SiteUrl :$($_.Exception.Message)"
    
    }
    

}  


function Set-SensitivityLabel {
    param (
        [string]$SiteUrl
    )
    try {
        Set-PnpTenantSite -Identity $SiteUrl -SensitivityLabel "Sensitive Site" -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to set sensitivity label on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red 
    }
}


function Add-GroupstoSharePointGroups {
    param (
        [string]$SiteUrl
    )

    try {
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $EntraGroupObjectId = "b4fa1e98-2893-4f26-bc64-f7a3e93b3753" # Replace with actual Entra Group Object ID
        $groupLoginName = "c:0t.c|tenant|$EntraGroupObjectId"
        Add-PnPGroupMember `
            -Group $ownersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-Host "Added Entra group with Object ID $EntraGroupObjectId to Owners group on $($SiteUrl)" -ForegroundColor Green    
    $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop
        Add-PnPGroupMember `
            -Group $membersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-Host "Added Entra group with Object ID $EntraGroupObjectId to Members group on $($SiteUrl)" -ForegroundColor Green
    
    }
    catch {
        Write-Host "ERROR: Failed to add group to SharePoint group on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red 
    }
    
}

Function Set-DocLibraryPermissions {
    param(
        [string]$SiteUrl
    )
    try {
        Write-Host "Setting DocLibraryPermissions on $SiteUrl" -ForegroundColor Yellow
        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-Host "ERROR: DocLibrary not found on $SiteUrl" -ForegroundColor Red
            throw "Document library 'Documents' was not found."
        }

        Set-PnPList -Identity $library -BreakRoleInheritance -CopyRoleAssignments -ErrorAction Stop

        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop

        if (-not $ownersGroup -or -not $membersGroup) {
            throw "Could not resolve Owners or Members group for $SiteUrl."
        }

        # Ensure both site groups have Contribute on Documents.
        Set-PnPListPermission `
            -Identity $library `
            -Group $ownersGroup.Title `
            -RemoveRole "Full Control" `
            -AddRole "Contribute" `
            
        Set-PnPListPermission `
            -Identity $library `
            -Group $membersGroup.Title `
            -RemoveRole "Edit"
        -AddRole "Contribute" `
            
        Write-Host "Permissions updated: Owners and Members have Contribute on Documents." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to set DocLibraryPermissions: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }

}


function Add-SiteToHubAssociation {
    param(
        [string]$SiteUrl,
        [string]$TargetHubSiteUrl
    )

    try {
       
        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $TargetHubSiteUrl -ErrorAction Stop
        Write-Host "Site $SiteUrl added to hub $TargetHubSiteUrl" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to add site to hub association: $($_.Exception.Message)" -ForegroundColor Red
    }
}

try {
    $adminUrl = "https://$tenantName-admin.sharepoint.com"
    Write-Host "Connecting to SharePoint Admin Center: $adminUrl" -ForegroundColor Yellow
    Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to connect to SharePoint Admin Center: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    $rows = Import-Csv -Path $CsvPath -Encoding UTF8
    if ($null -eq $rows -or $rows.Count -eq 0) {
        Write-Host "ERROR: CSV has no rows." -ForegroundColor Red
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        exit 1
    }

    if ($rows[0].PSObject.Properties.Name -notcontains "SiteUrl") {
        Write-Host "ERROR: CSV must contain the required 'SiteUrl' column header." -ForegroundColor Red
        Write-Host "Example header: SiteUrl" -ForegroundColor Yellow
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        exit 1
    }

    $total = $rows.Count
    $index = 0
    $success = 0
    $failed = 0

    foreach ($row in $rows) {
        $index++
        $siteUrl = if ($row.SiteUrl) { $row.SiteUrl.ToString().Trim() } else { "" }

        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            Write-Host "[$index/$total] Skipping row because required 'SiteUrl' is empty." -ForegroundColor Yellow
            continue
        }

        try {
            Write-Host "[$index/$total] Associating site to hub: $siteUrl" -ForegroundColor Cyan
            Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint -ErrorAction Stop
            # Add-SiteToHubAssociation -SiteUrl $siteUrl -TargetHubSiteUrl $HubSiteUrl
            #Set-SiteRegionalSettings -SiteUrl $siteUrl
            #Set-DocLibraryPermissions -SiteUrl $siteUrl
            #Set-SensitivityLabel -SiteUrl $siteUrl
            Add-GroupstoSharePointGroups -SiteUrl $siteUrl
            
            $success++
            Write-Host "[$index/$total] Associated successfully." -ForegroundColor Green
        }
        catch {
            $failed++
            Write-Host "[$index/$total] ERROR: Failed to associate '$siteUrl' to hub '$HubSiteUrl'. $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "Hub association completed." -ForegroundColor Green
    Write-Host "Success: $success" -ForegroundColor Green
    Write-Host "Failed:  $failed" -ForegroundColor Yellow
    Write-Host "Total:   $total" -ForegroundColor Cyan
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

