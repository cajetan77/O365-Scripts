param(
    [int]$DaysThreshold = 30
)

$ErrorActionPreference = "Stop"



# Clear any existing connections and modules to avoid assembly conflicts
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Remove-Module Microsoft.Graph.* -Force -ErrorAction SilentlyContinue
}
catch {
    # Ignore errors when cleaning up
}

# Load Microsoft Graph modules with error handling
try {
    Write-Host "Loading Microsoft Graph modules..." -ForegroundColor Yellow
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -Force -ErrorAction Stop
    Write-Host "Microsoft Graph modules loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error loading Microsoft Graph modules: $_" -ForegroundColor Red
    Write-Host "Trying to install Microsoft Graph PowerShell..." -ForegroundColor Yellow
    
    try {
        Install-Module Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
        Import-Module Microsoft.Graph.Authentication -Force
        Import-Module Microsoft.Graph.Applications -Force
        Write-Host "Microsoft Graph modules installed and loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install Microsoft Graph modules: $_" -ForegroundColor Red
        exit 1
    }
}

# Try different authentication methods to avoid assembly conflicts
$connected = $false

# Method 1: Try interactive authentication first (most reliable)


# Method 2: Try certificate authentication if interactive fails
if (-not $connected) {
    try {
        Write-Host "Attempting certificate authentication..." -ForegroundColor Yellow
        $config = Get-Content -Raw -Path ".\config.json" | ConvertFrom-Json
        $TenantId = $config.TenantId
        $ClientId = $config.SharePointReportingAppId
        $Thumbprint = $config.ThumbPrint
        
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome
        Write-Host "Connected to Microsoft Graph successfully via certificate" -ForegroundColor Green
        $connected = $true
    }
    catch {
        Write-Host "Certificate authentication failed: $_" -ForegroundColor Yellow
    }
}



$now = (Get-Date).ToUniversalTime()
$cutoff = $now.AddDays($DaysThreshold)

# Pull app registrations with credential properties
$apps = Get-MgApplication -All -Property "id,displayName,appId,passwordCredentials,keyCredentials"

$results = foreach ($app in $apps) {

    # Client secrets
    foreach ($pc in ($app.PasswordCredentials | Where-Object { $_ })) {
        $end = $pc.EndDateTime.ToUniversalTime()

        $status =
        if ($end -lt $now) { "Expired" }
        elseif ($end -le $cutoff) { "ExpiringSoon" }
        else { "Valid" }

        [pscustomobject]@{
            Status         = $status
            DaysRemaining  = [math]::Floor(($end - $now).TotalDays)
            DisplayName    = $app.DisplayName
            AppId          = $app.AppId
            ExpiredDate    = $pc.EndDateTime.ToUniversalTime().ToString("o")
            CredentialType = "Secret"
            KeyId          = $pc.KeyId
            Hint           = $pc.Hint
            EndDateTimeUtc = $pc.EndDateTime.ToUniversalTime().ToString("o")
        }
    }

    # Certificates
    foreach ($kc in ($app.KeyCredentials | Where-Object { $_ })) {
        $end = $kc.EndDateTime.ToUniversalTime()

        $status =
        if ($end -lt $now) { "Expired" }
        elseif ($end -le $cutoff) { "ExpiringSoon" }
        else { "Valid" }

        [pscustomobject]@{
            Status         = $status
            DaysRemaining  = [math]::Floor(($end - $now).TotalDays)
            DisplayName    = $app.DisplayName
            AppId          = $app.AppId
            ExpiredDate    = $pc.EndDateTime.ToUniversalTime().ToString("o")
            CredentialType = "Certificate"
            KeyId          = $kc.KeyId
            Hint           = $null
            EndDateTimeUtc = $kc.EndDateTime.ToUniversalTime().ToString("o")
        }
    }
}

# Show only expired / expiring soon
$report = $results |
Where-Object { $_.Status -in @("Expired", "ExpiringSoon") } |
Sort-Object Status, DaysRemaining, DisplayName

if (-not $report) {
    Write-Output "No App Registration secrets/certs are expired or expiring within $DaysThreshold days."
    # Create empty report for SharePoint upload
    $report = @([PSCustomObject]@{
            DisplayName   = "No expiring apps found"
            AppId         = ""
            Type          = ""
            KeyId         = ""
            StartDate     = ""
            EndDate       = ""
            DaysRemaining = ""
            Status        = "No Issues"
        })
}
else {
    Write-Host "Found $($report.Count) expiring/expired app registrations:" -ForegroundColor Yellow
    $report | Format-Table -AutoSize
}

# Export report to SharePoint AzureAppExpiry list using Microsoft Graph
Write-Host "`nExporting report to SharePoint list using Microsoft Graph..." -ForegroundColor Cyan

try {
    # SharePoint site details
    $siteHostname = "caje77sharepoint.sharepoint.com"
    $siteName = "M365Updates"
    $listName = "AzureAppExpiry"
    
    Write-Host "Getting SharePoint site information..." -ForegroundColor Yellow
    
    # Get the site using the correct format: hostname:/sites/sitename
    $siteId = "${siteHostname}:/sites/${siteName}"
    Write-Host "Looking for site with ID: $siteId" -ForegroundColor Gray
    
    try {
        $site = Get-MgSite -SiteId $siteId
        Write-Host "✓ Found SharePoint site: $($site.DisplayName)" -ForegroundColor Green
    }
    catch {
        # Try alternative method - search by name
        Write-Host "Direct site access failed, searching by name..." -ForegroundColor Yellow
        $sites = Get-MgSite -Search $siteName
        $site = $sites | Where-Object { $_.DisplayName -eq $siteName -or $_.Name -eq $siteName } | Select-Object -First 1
        
        if (-not $site) {
            Write-Host "Could not find SharePoint site: $siteName" -ForegroundColor Red
            Write-Host "Available sites:" -ForegroundColor Yellow
            $allSites = Get-MgSite -All | Select-Object -First 10
            foreach ($s in $allSites) {
                Write-Host "  - $($s.DisplayName) ($($s.WebUrl))" -ForegroundColor Gray
            }
            return
        }
        Write-Host "✓ Found SharePoint site via search: $($site.DisplayName)" -ForegroundColor Green
    }
    
    # Get the list
    $list = Get-MgSiteList -SiteId $site.Id | Where-Object { $_.DisplayName -eq $listName }
    if (-not $list) {
        Write-Host "List '$listName' not found. Please create the list with the following columns:" -ForegroundColor Red
        Write-Host "  - Title (Single line of text) - for App Name" -ForegroundColor Yellow
        Write-Host "  - AppId (Single line of text) - for App ID" -ForegroundColor Yellow
        Write-Host "  - ExpiryDate (Date and Time) - for Expiry Date" -ForegroundColor Yellow
        Write-Host "  - CredentialType (Choice: Secret, Certificate) - for Credential Type" -ForegroundColor Yellow
        return
    }
    Write-Host "✓ Found list: $($list.DisplayName)" -ForegroundColor Green
    
    # Get existing items to check for updates
    Write-Host "Getting existing items from list..." -ForegroundColor Yellow
    $existingItems = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ExpandProperty "fields" 
    
    # Create a hashtable for quick lookup by KeyId
    $existingItemsLookup = @{}
    foreach ($existingItem in $existingItems) {
        $keyId = $existingItem.Fields.AdditionalProperties["KeyId"]
        if ($keyId) {
            $existingItemsLookup[$keyId] = $existingItem
        }
    }
    Write-Host "Found $($existingItems.Count) existing items" -ForegroundColor Green
    
    # Process report items - update existing or create new
    Write-Host "Processing $($report.Count) items (update existing or create new)..." -ForegroundColor Yellow
    $successCount = 0
    $errorCount = 0
    
    foreach ($item in $report) {
        try {
            # Prepare the fields according to your specification
            $fields = @{
                "Title"          = $item.DisplayName          # App Name
                "AppId"          = $item.AppId                # App ID
                "ExpiryDate"     = $item.ExpiredDate        # Expiry Date
                "CredentialType" = $item.CredentialType
                "KeyId"          = $item.KeyId       
            }
            
            # Check if item with this KeyId already exists
            $existingItem = $existingItemsLookup[$item.KeyId]
            
            if ($existingItem) {
                # Update existing item
                $listItemBody = @{
                    fields = $fields
                }
                
                Update-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $existingItem.Id -BodyParameter $listItemBody | Out-Null
                $successCount++
                Write-Host "  ✓ Updated: $($item.DisplayName) (KeyId: $($item.KeyId))" -ForegroundColor Cyan
            }
            else {
                # Create new item
                $listItemBody = @{
                    fields = $fields
                }
                
                New-MgSiteListItem -SiteId $site.Id -ListId $list.Id -BodyParameter $listItemBody | Out-Null
                $successCount++
                Write-Host "  ✓ Created: $($item.DisplayName) (KeyId: $($item.KeyId))" -ForegroundColor Green
            }
        }
        catch {
            $errorCount++
            Write-Host "  ✗ Failed to process: $($item.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "✓ Successfully processed $successCount items (updated existing or created new)" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "✗ Failed to add $errorCount items" -ForegroundColor Red
    }
    
    Write-Host "List URL: https://$siteHostname$sitePath/Lists/$listName" -ForegroundColor Cyan
}
catch {
    Write-Host "Error exporting to SharePoint using Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Report data is still available in the `$report variable" -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null
Write-Host "`nScript completed!" -ForegroundColor Green