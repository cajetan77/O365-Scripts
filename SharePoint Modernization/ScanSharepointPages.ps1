param(
    [string]$SiteUrl = ""
)


$configPath = ".\config.json"

$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json

$TenantId = $config.TenantId
$ClientId = $config.AppId
$Thumbprint = $config.ThumbPrint
$TenantName = $config.TenantName
$allClassic = New-Object System.Collections.Generic.List[object]


function Connect-Site {
    param(
        [string]$Url
    )
    try {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to connect to site $($Url): $_" -ForegroundColor Red
        exit 1
    }
}


function Get-SitePages {
    param(
        [string]$SiteUrl
    )
    $Lists = Get-PnPList  | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.RootFolder.ServerRelativeUrl -like "*/*Pages*" }
    $results = @()
 
    foreach ($List in $Lists) {
        $items = Get-PnPListItem -Connection $Connection -List $List.Title -IncludeContentType  -PageSize 500 `
        -Fields "FileRef","FileLeafRef","File_x0020_Type","ContentTypeId","CanvasContent1","ClientSideApplicationId","WikiField","PublishingPageLayout" `
        -ErrorAction SilentlyContinue
        foreach ($i in $items) {
            $fileType = $i.FieldValues["File_x0020_Type"]
            if ($fileType -ne "aspx") { continue }
        
            $fileRef  = $i.FieldValues["FileRef"]
            $leaf     = $i.FieldValues["FileLeafRef"]
            $ctid     = [string]$i.FieldValues["ContentTypeId"]
            $contentType = $i.ContentType.Name
            # Modern page indicators (most reliable in Site Pages)
            $canvas   = $i.FieldValues["CanvasContent1"]
            $csAppId  = $i.FieldValues["ClientSideApplicationId"]
        
            $isModern = ($null -ne $canvas -and "$canvas".Trim().Length -gt 0) -or ($null -ne $csAppId -and "$csAppId".Trim().Length -gt 0)
            if ($isModern) { continue }  # we only want classic pages
        
            # Classify classic page type
            $pageType = "Classic (Unknown ASPX)"
        
            # Classic Wiki Page content type commonly starts with 0x010108
            if ($ctid -like "0x010108*") {
                Write-Host "Classic (Wiki Page) found in $($SiteUrl) - $($List.Title) - $($leaf)" -ForegroundColor Yellow
              $pageType = "Classic (Wiki Page)"
            }
            # Classic Web Part Page content type commonly starts with 0x010109
            elseif ($ctid -like "0x010109*") {
                Write-Host "Classic (Web Part Page) found in $($SiteUrl) - $($List.Title) - $($leaf)" -ForegroundColor Yellow
              $pageType = "Classic (Web Part Page)"
            }
            # Publishing pages usually live in "Pages" and often have Publishing fields/layouts
            elseif ($null -ne $i.FieldValues["PublishingPageLayout"]) {
                Write-Host "Classic (Publishing Page) found in $($SiteUrl) - $($List.Title) - $($leaf)" -ForegroundColor Yellow
              $pageType = "Classic (Publishing Page)"
            }
            # Fallback: if WikiField exists/has value, treat as Wiki
            elseif ($null -ne $i.FieldValues["WikiField"] -and "$($i.FieldValues["WikiField"])" -ne "") {
                Write-Host "Classic (Wiki Page) found in $($SiteUrl) - $($List.Title) - $($leaf)" -ForegroundColor Yellow
              $pageType = "Classic (Wiki Page)"
            }
        
            $results += [pscustomobject]@{
              SiteUrl        = $SiteUrl
              Library        = $List.Title
              PageName       = $leaf
              ServerRelative = $fileRef
              FullUrl        = ($SiteUrl.TrimEnd('/') + $fileRef)
              PageType       = $pageType
              ContentTypeId  = $ctid
              ContentType    = $contentType
            }
          }
        }

    return $results
    }
  



$adminUrl = "https://$TenantName-admin.sharepoint.com"
Connect-Site -Url $adminUrl

$sites = Get-PnPTenantSite 
$siteCount = $sites.Count
$idx = 0

# Collect all sites including subsites
$allSitesToProcess = New-Object System.Collections.Generic.List[object]

foreach ($site in $sites) {
    $idx++
    Write-Host ("[{0}/{1}] Discovering sites and subsites for {2}" -f $idx, $siteCount, $site.Url) -ForegroundColor Cyan
    
    # Add the main site
    $allSitesToProcess.Add([pscustomobject]@{ Url = $site.Url; IsSubsite = $false })
    
    # Get all subsites
    try {
        Connect-Site -Url $site.Url
        $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
        foreach ($subsite in $subsites) {
            $allSitesToProcess.Add([pscustomobject]@{ Url = $subsite.Url; IsSubsite = $true })
        }
        Write-Host "  Found $($subsites.Count) subsite(s)" -ForegroundColor Gray
    }
    catch {
        Write-Host "  WARNING: Could not get subsites: $_" -ForegroundColor Yellow
    }
}

$totalSitesToProcess = $allSitesToProcess.Count
$idx = 0

Write-Host "`nProcessing $totalSitesToProcess site(s) (including subsites)...`n" -ForegroundColor Green

foreach ($siteInfo in $allSitesToProcess) {
    $idx++
    $siteType = if ($siteInfo.IsSubsite) { "Subsite" } else { "Site" }
    Write-Host ("[{0}/{1}] Scanning {2} ({3})" -f $idx, $totalSitesToProcess, $siteInfo.Url, $siteType) -ForegroundColor Green
    
    try {
        Connect-Site -Url $siteInfo.Url
        $siteResults = Get-SitePages -SiteUrl $siteInfo.Url
        if ($siteResults -and $siteResults.Count -gt 0) {
            foreach ($result in $siteResults) {
                $allClassic.Add($result)
            }
            Write-Host "  Found $($siteResults.Count) classic page(s)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  ERROR: Failed to process $($siteInfo.Url): $_" -ForegroundColor Red
    }
}

Disconnect-PnPOnline

$allClassic `
  | Sort-Object SiteUrl, Library, PageName `
  | Export-Csv -Path ".\allClassic.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Exported to: .\allClassic.csv" -ForegroundColor Cyan