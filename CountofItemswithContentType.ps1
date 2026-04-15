

Write-Host "Starting Count of Items with Content Type Script..." -ForegroundColor Green
Write-Host ""

# Load configuration
Write-Host "Loading configuration from config.json..." -ForegroundColor Yellow
$configPath = ".\config.json"
if (-not (Test-Path $configPath)) {
    Write-Host "ERROR: config.json not found at $configPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$TenantName = $config.TenantName    
$Thumbprint = $config.ThumbPrint
$contentTypeName = $config.ContentTypeName
$ctReport = New-Object System.Collections.Generic.List[object]
$outCsv = ".\ContentTypesReport.csv"
$ExcludedLists = @("Site Assets", "Site Pages", "Form Templates", "Style Library")

Write-Host "Configuration loaded successfully." -ForegroundColor Green
Write-Host "Tenant: $TenantName" -ForegroundColor Cyan
Write-Host "Content Type Name: $contentTypeName" -ForegroundColor Cyan
Write-Host ""

# Connect to admin center
Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
    Write-Host "Connected to Admin Center successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to connect to Admin Center: $_" -ForegroundColor Red
    exit 1
}

# Get content type hub URL
Write-Host "Getting Content Type Hub URL..." -ForegroundColor Yellow
try {
    $cthubUrl = Get-PnPContentTypePublishingHubUrl -ErrorAction Stop
    Write-Host "Content Type Hub URL: $cthubUrl" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to get Content Type Hub URL: $_" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

# Connect to content type hub
Write-Host "Connecting to Content Type Hub..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $cthubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
    Write-Host "Connected to Content Type Hub successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to connect to Content Type Hub: $_" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

# Get hub content type
Write-Host "Searching for Content Type '$contentTypeName' in Hub..." -ForegroundColor Yellow
$ct = Get-PnPContentType | Where-Object { $_.Name -eq $contentTypeName }
if ($null -eq $ct) {
    Write-Host "ERROR: Content Type '$contentTypeName' not found in Content Type Hub!" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

Write-Host "Content Type found in Hub!" -ForegroundColor Green
Write-Host "Content Type ID: $($ct.Id.StringValue)" -ForegroundColor Cyan
$hubId = $ct.Id.StringValue
#$hubId = "0x01010008B553E15AC913458E6A45D46B4BC319"
Disconnect-PnPOnline

# Reconnect to admin center
Write-Host "Reconnecting to Admin Center..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to reconnect to Admin Center: $_" -ForegroundColor Red
    exit 1
}

# Get sites
Write-Host "Retrieving sites matching pattern..." -ForegroundColor Yellow
$sites = Get-PnPTenantSite | Where-Object { $_.Url -like "https://caje77sharepoint.sharepoint.com/sites/tesmpc906*" }
$siteCount = ($sites | Measure-Object).Count
Write-Host "Found $siteCount site(s) to process." -ForegroundColor Green
Write-Host ""
$siteIndex = 0
foreach ($site in $sites) {
    $siteIndex++
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Processing Site $siteIndex of $siteCount - $($site.Url)" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    
    try {
        Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to connect to site $($site.Url): $_" -ForegroundColor Red
    
        continue
    }

    $siteCT = Get-PnPContentType | Where-Object { $_.Name -eq $contentTypeName }
    if ($null -ne $siteCT) {
        Write-Host "Content Type '$contentTypeName' exists in the site $($site.Url)."
        Write-Host "Content Type ID: $($siteCT.Id.StringValue)"
        if ($siteCT.Id.StringValue -eq $hubId) {
            Write-Host "The Content Type ID matches the hub Content Type ID $($site.Url). "
            
            $libraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and -not ($ExcludedLists -contains $_.Title) }
            foreach ($library in $libraries) {
                Write-Host "Checking library: $($library.Title)"
                $count = Get-PnpListItem -List $library.Title | Where-Object {$_.FieldValues.ContentTypeId.StringValue.Contains($hubId)} | Measure-Object | Select-Object -ExpandProperty Count
                Write-Host "Count of items with Content Type '$contentTypeName' in library $($library.Title): $count"
                $libCT = Get-PnPContentType -List $library.Title | Where-Object { $_.Name -eq $contentTypeName }
                
                if ($null -ne $libCT) {
                    Write-Host "Content Type '$contentTypeName' exists in the library $($library.Title). in $($site.Url)"
                    Write-Host "Content Type ID in library: $($libCT.Id.StringValue)"
                   # Write-Host "Parent Content Type ID in library: $($libCT.Parent)"
                    if ($libCT.Id.StringValue.Contains($hubId)) {
                        Write-Host "The Content Type ID in library matches the hub Content Type ID."
                        $ctReport.Add([pscustomobject]@{
                                SiteUrl         = $site.Url
                                ContentTypeName = $contentTypeName
                                Status          = "Matched in Library"
                                library         = $library.Title
                             
                                CountofDocuments           = $count
                            })
                    }
                    else {
                        Write-Host "The Content Type ID in library does NOT CONTAIN the hub Content Type ID."
                        $ctReport.Add([pscustomobject]@{
                                SiteUrl         = $site.Url
                                ContentTypeName = $contentTypeName
                                Status          = "Library Content Type ID Not Matched"
                                library         = $library.Title
                               
                            })
                    }
                }
                else {
                    Write-Host "Content Type '$contentTypeName' does NOT exist in the library $($library.Title)."
                    $ctReport.Add([pscustomobject]@{
                            SiteUrl         = $site.Url
                            ContentTypeName = $contentTypeName
                            Status          = "Content Type Missing in Library"
                            library         = $library.Title
                            Parent          = ""
                        })
                   
                }
            } 
        }                    
        else {
            Write-Host "The Content Type ID does NOT match the hub Content Type ID $($site.Url)."

            $ctReport.Add([pscustomobject]@{
                    SiteUrl         = $site.Url
                    ContentTypeName = $contentTypeName
                    Status          = "Site Content Type ID Not Matched"
                    library         = ""
                    Parent          = ""
                })
         
            # Still check libraries even if site content type doesn't match
          
        }
    }

    Write-Host "Content Type '$contentTypeName' is missing in the site $($site.Url)."
    $ctReport.Add([pscustomobject]@{
            SiteUrl         = $site.Url
            ContentTypeName = $contentTypeName
            Status          = "Content Type Missing in Site"
            library         = ""
            Parent          = ""
        })
    
    Disconnect-PnPOnline

}

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Disconnecting from SharePoint..." -ForegroundColor Yellow
Disconnect-PnPOnline

Write-Host "Exporting results to CSV..." -ForegroundColor Yellow
$ctReport | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
$reportCount = ($ctReport | Measure-Object).Count
Write-Host "Export complete! $reportCount record(s) written to $outCsv" -ForegroundColor Green
Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green