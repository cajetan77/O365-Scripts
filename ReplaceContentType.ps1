<#
.SYNOPSIS
    Replaces a content type in a site with a new content type from the content type hub.
.DESCRIPTION
    This script connects to the SharePoint Admin Center, gets the content type hub URL, and then replaces a content type in a site with a new content type from the content type hub.

.PARAMETER ContentTypeName
    The name of the content type to be replaced.
.PARAMETER OldContentTypeName
    The name of the old content type to be replaced.
.PARAMETER NewContentTypeName
    The name of the new content type to be replaced.

.PARAMETER ReportPath
    The path to the report file.
.PARAMETER ReportWithItemsPath
    The path to the report file with items.
.PARAMETER reportOnly
    If true, only the report will be generated.

#>
param(
 
    [string]$OldContentTypeName = "",
    [string]$newContentTypeName = "",
    [string]$ReportPath = ".\Report.csv",
    [string]$ReportWithItemsPath = ".\ReportWithItems.csv",
    [string]$ExcludedLists = @("Site Assets", "Site Pages", "Form Templates", "Style Library"),
    [string]$Reprtwithcontenttype= ".\Reportwithcontenttype.csv",
    [bool]$reportOnly = $false
)
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
$ExcludedLists = @("Site Assets", "Site Pages", "Form Templates", "Style Library")
$outCsv = $ReportPath
$ctReport = New-Object System.Collections.Generic.List[object]
$ctReportwithItems = New-Object System.Collections.Generic.List[object]
$ctReportwithcontenttype = New-Object System.Collections.Generic.List[object]



Write-Host "Configuration loaded successfully." -ForegroundColor Green
Write-Host "Tenant: $TenantName" -ForegroundColor Cyan
Write-Host "Content Type Name: $contentTypeName" -ForegroundColor Cyan
Write-Host ""

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


function Get-ContentTypeHub {
    param(
        [string]$ContentTypeName
    )
    try {
        $cthubUrl = Get-PnPContentTypePublishingHubUrl -ErrorAction Stop
        Write-Host "Content Type Hub URL: $cthubUrl" -ForegroundColor Green
        Connect-Site -Url $cthubUrl
        $ct = Get-PnPContentType | Where-Object { $_.Name -eq $ContentTypeName }
        if ($null -eq $ct) {
            Write-Host "ERROR: Content Type '$ContentTypeName' not found in Content Type Hub!" -ForegroundColor Red
            Disconnect-PnPOnline
            exit 1
        }
        Write-Host "Content Type found in Hub!" -ForegroundColor Green
        Write-Host "Content Type Name: $($ct.Name)" -ForegroundColor Cyan
        Write-Host "Content Type ID: $($ct.Id.StringValue)" -ForegroundColor Cyan
        return $ct.Id.StringValue
        Disconnect-PnPOnline
    }
    catch {
        Write-Host "ERROR: Failed to get Content Type Hub URL: $_" -ForegroundColor Red
        exit 1
    }
}


# Connect to admin center
Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Yellow
Connect-Site -Url "https://$TenantName-admin.sharepoint.com"



# Get hub content type
Write-Host "Searching for Content Type '$contentTypeName' in Hub..." -ForegroundColor Yellow

$oldct = Get-ContentTypeHub -ContentTypeName $OldContentTypeName
Write-Host "Old Content Type ID: $oldct" -ForegroundColor Cyan
Write-Host "Old Content Type Name: $OldContentTypeName" -ForegroundColor Cyan

$newct = Get-ContentTypeHub -ContentTypeName $NewContentTypeName
Write-Host "New Content Type ID: $newct" -ForegroundColor Cyan
Write-Host "New Content Type Name: $NewContentTypeName" -ForegroundColor Cyan




# Reconnect to admin center
Write-Host "Reconnecting to Admin Center..." -ForegroundColor Yellow
Connect-Site -Url "https://$TenantName-admin.sharepoint.com"

# Get sites
Write-Host "Retrieving sites matching pattern..." -ForegroundColor Yellow

#$sites = Get-PnPTenantSite | Where-Object { $_.Url -like "https://caje77sharepoint.sharepoint.com/sites/testqaproject661223322222222222ddssssssssssssssaaWWWDWWW1SQAQ*" }
$sites = Import-Csv -Path ".\sites.csv"
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
        Connect-Site -Url $site.Url
        $ctx=Get-PnPContext
    }
    catch {
        Write-Host "ERROR: Failed to connect to site $($site.Url): $_" -ForegroundColor Red
    
        continue
    }

    #$oldsiteCT = Get-PnPContentType | Where-Object { $_.Name -eq $oldcontentTypeName }
    $newsiteCT = Get-PnPContentType  | Where-Object { $_.Id.StringValue -contains $newct }
    if ($null -eq $newsiteCT) {
        if (!$ReportOnly) {
        Add-PnPContentTypesFromContentTypeHub -ContentTypes $newct -Site $site.Url
        Write-Host "Content Type '$newcontentTypeName' added to the site $($site.Url)"
        }
    }
    $oldsiteCT = Get-PnPContentType  | Where-Object { $_.Id.StringValue -contains $oldct }
    if ($null -ne $oldsiteCT) {
        $libraries = Get-PnPList  | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and -not ($ExcludedLists -contains $_.Title) }
        foreach ($library in $libraries) {

            if (!$ReportOnly) {
                Add-PnPContentTypeToList -List $library.Title -ContentType $newcontentTypeName
                Write-Host "Content Type '$newcontentTypeName' added to the library $($library.Title). in $($site.Url)"
            }
            else {
                $libraryCT = Get-PnPContentType -List $library.Title 
                $ctReportwithcontenttype.Add([pscustomobject]@{
                    SiteUrl          = $site.Url
                    ContentTypeName  = $library.Title
                    ContentTypeId    = $libraryCT.Name -join ","
                })
            }
           
            $oldctItems = Get-PnpListItem -List $library.Title -IncludeContentType | Where-Object { $_.FieldValues.ContentTypeId.StringValue.Contains($oldContentTypeId) }

            Write-Host "Found $($oldctItems.Count) items with Content Type '$oldcontentTypeName' in library $($library.Title). in $($site.Url)"

            if ($oldctItems.Count -eq 0) {
                try {
                    if (!$reportOnly) {
                        Remove-PnPContentTypeFromList -List $library.Title -ContentType $oldcontentTypeName
                        Write-Host "Content Type '$oldcontentTypeName' removed from the library $($library.Title). in $($site.Url)"
                    }
                    else {
                        $ctReport.Add([pscustomobject]@{
                                SiteUrl          = $site.Url
                                ContentTypeName  = $oldcontentTypeName
                                library          = $library.Title
                                CountofDocuments = $oldctItems.Count
                            })
                    }

                    
                }
                catch {
                    Write-Host "ERROR: Failed to remove content type '$oldcontentTypeName' from the library $($library.Title). in $($site.Url): $_" -ForegroundColor Red
                    continue
                }
            }
            else {
                Write-Host "Content Type '$oldcontentTypeName' still exists in the library $($library.Title). in $($site.Url)"
                IF ($ReporOnly) {
                    foreach ($item in $oldctItems) {
                      $ctx.Load($item)
                        $ctx.ExecuteQuery()
                        $ctReportwithItems.Add([pscustomobject]@{
                                SiteUrl          = $site.Url
                                ContentTypeName  = $item.ContentType.Name
                                library          = $library.Title
                                Title            = $item.FieldValues.Title
                                ItemUrl          = $item.FieldValues.FileRef
                                ItemId           = $item.Id
                                ContentTypeId    = $item.ContentType.Id.StringValue
                                CreatedBy        = $item.FieldValues.Author.LookupValue
                                ModifiedBy       = $item.FieldValues.Editor.LookupValue
                                Modified         = $item.FieldValues.Modified.DateTime
                                Created          = $item.FieldValues.Created.DateTime
                    

                            })
                    }
                }
            }
        }
    }

    Disconnect-PnPOnline
}


$ctReport | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
$ctReportwithItems | Export-Csv -Path $ReportWithItemsPath -NoTypeInformation -Encoding UTF8
$ctReportwithcontenttype | Export-Csv -Path $Reprtwithcontenttype -NoTypeInformation -Encoding UTF8

Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host ""