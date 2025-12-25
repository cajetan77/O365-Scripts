<#

.SYNOPSIS
    Gets the web parts from all classic pages and exports the results to a CSV file.
.DESCRIPTION
    This script gets the web parts from all classic pages and exports the results to a CSV file.
.PARAMETER CsvPath
    The path to the CSV file containing the classic pages.
.PARAMETER OutputPath
    The path to the output CSV file.
#>
param(
    [string]$CsvPath = ".\allClassic.csv",
    [string]$OutputPath = ".\ClassicWebPartsReport.csv"
)

# Load configuration
$configPath = ".\config.json"
if (-not (Test-Path $configPath)) {
    Write-Host "ERROR: config.json not found at $configPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$Thumbprint = $config.ThumbPrint
$TenantName = $config.TenantName

# Check if CSV file exists
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at $CsvPath" -ForegroundColor Red
    exit 1
}

# Read the CSV file
Write-Host "Reading classic pages from $CsvPath..." -ForegroundColor Yellow
$classicPages = Import-Csv -Path $CsvPath -Encoding UTF8
$totalPages = $classicPages.Count
Write-Host "Found $totalPages classic page(s) to process.`n" -ForegroundColor Green

$webPartsReport = New-Object System.Collections.Generic.List[object]
$currentSiteUrl = ""
$connection = $null

function Connect-Site {
    param(
        [string]$Url
    )
    
    if ($null -ne $connection -and $currentSiteUrl -eq $Url) {
        return $connection
    }
    
    try {
        Write-Host "Connecting to site: $Url" -ForegroundColor Cyan
        $conn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ReturnConnection -ErrorAction Stop
        $script:currentSiteUrl = $Url
        $script:connection = $conn
        return $conn
    }
    catch {
        Write-Host "ERROR: Failed to connect to site $($Url): $_" -ForegroundColor Red
        return $null
    }
}

function Get-WebPartsFromPage {
    param(
        [string]$SiteUrl,
        [string]$ServerRelativeUrl,
        [string]$PageName,
        [string]$Library,
        [string]$PageType,
        [object]$Connection
    )
    
    $webParts = @()
    
    try {
        # Get the file
        $file = Get-PnPFile -Url $ServerRelativeUrl -AsListItem -Connection $Connection -ErrorAction Stop
        
        # For classic pages, web parts are stored in the page's web part manager
        # We need to get the page and extract web parts from it
        $pageUrl = $ServerRelativeUrl
        
        # Try to get web parts using Get-PnPWebPart (works for both classic and modern)
        try {
            $webPartXml = Get-PnPWebPart -ServerRelativePageUrl $pageUrl -Connection $Connection -ErrorAction SilentlyContinue #| Select Id, Title, Properties, WebPartType, ZoneId, ZoneIndex, IsClosed, Hidden
            
            if ($null -ne $webPartXml -and $webPartXml.Count -gt 0) {
                # Handle both array and object with WebPart property
                $webPartsToProcess = if ($webPartXml.WebPart) { $webPartXml.WebPart } else { $webPartXml }
                foreach ($wp in $webPartsToProcess) {
                    # Get web part title - Title property is directly available from Get-PnPWebPart
                    $wpTitle = if ($null -ne $wp.Title -and $wp.Title -ne "") { 
                        $wp.Title 
                    } 
                    else { 
                        "Untitled" 
                    }

                    # Extract all field values using Get-PnPWebPartProperty
                    $allFieldValues = @{}
                    
                    # Only try Get-PnPWebPartProperty if we have a valid Id
                    if ($null -ne $wp.Id -and $wp.Id -ne "") {
                        try {
                            $wpProperties = Get-PnPWebPartProperty -ServerRelativePageUrl $pageUrl -Identity $wp.Id -Connection $Connection -ErrorAction SilentlyContinue
                            if ($null -ne $wpProperties) {
                            # Get all properties from the returned object
                            $properties = $wpProperties | Get-Member -MemberType NoteProperty, Property
                            foreach ($prop in $properties) {
                                $propName = $prop.Name
                                $propValue = $wpProperties.$propName
                                
                                # Skip certain system properties
                                if ($propName -notin @('PSComputerName', 'PSShowComputerName', 'RunspaceId', 'PSRemotingBehavior')) {
                                    # Convert to string, handling null and complex objects
                                    if ($null -eq $propValue) {
                                        $allFieldValues[$propName] = $null
                                    }
                                    elseif ($propValue -is [string] -or $propValue -is [int] -or $propValue -is [bool] -or $propValue -is [DateTime] -or $propValue -is [System.Guid]) {
                                        $allFieldValues[$propName] = $propValue.ToString()
                                    }
                                    else {
                                        # For complex objects, convert to JSON
                                        try {
                                            $allFieldValues[$propName] = ($propValue | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
                                        }
                                        catch {
                                            $allFieldValues[$propName] = $propValue.ToString()
                                        }
                                    }
                                }
                            }
                            
                            # Also try to access FieldValues if it exists
                            if ($null -ne $wpProperties.FieldValues) {
                                foreach ($key in $wpProperties.FieldValues.Keys) {
                                    $value = $wpProperties.FieldValues[$key]
                                    if ($null -eq $value) {
                                        $allFieldValues["FV_$key"] = $null
                                    }
                                    elseif ($value -is [string] -or $value -is [int] -or $value -is [bool] -or $value -is [DateTime]) {
                                        $allFieldValues["FV_$key"] = $value.ToString()
                                    }
                                    else {
                                        try {
                                            $allFieldValues["FV_$key"] = ($value | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
                                        }
                                        catch {
                                            $allFieldValues["FV_$key"] = $value.ToString()
                                        }
                                    }
                                }
                            }
                            }
                        }
                        catch {
                            # If Get-PnPWebPartProperty fails, try accessing Properties directly
                            $errorMsg = $_.Exception.Message
                            Write-Host "  Warning: Could not get properties for web part $($wp.Id): $errorMsg" -ForegroundColor Yellow
                            
                            if ($null -ne $wp.Properties) {
                                try {
                                    $props = $wp.Properties | Get-Member -MemberType NoteProperty, Property -ErrorAction SilentlyContinue
                                    if ($null -ne $props) {
                                        foreach ($prop in $props) {
                                            $propName = $prop.Name
                                            $propValue = $wp.Properties.$propName
                                            if ($null -ne $propValue) {
                                                try {
                                                    $allFieldValues[$propName] = ($propValue | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
                                                }
                                                catch {
                                                    $allFieldValues[$propName] = $propValue.ToString()
                                                }
                                            }
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "  Warning: Could not access Properties directly: $_" -ForegroundColor Yellow
                                }
                            }
                        }
                    }
                    else {
                        # No Id available, try accessing Properties directly
                        if ($null -ne $wp.Properties) {
                            try {
                                $props = $wp.Properties | Get-Member -MemberType NoteProperty, Property -ErrorAction SilentlyContinue
                                if ($null -ne $props) {
                                    foreach ($prop in $props) {
                                        $propName = $prop.Name
                                        $propValue = $wp.Properties.$propName
                                        if ($null -ne $propValue) {
                                            try {
                                                $allFieldValues[$propName] = ($propValue | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
                                            }
                                            catch {
                                                $allFieldValues[$propName] = $propValue.ToString()
                                            }
                                        }
                                    }
                                }
                            }
                            catch {
                                Write-Host "  Warning: Could not access Properties directly (no Id): $_" -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    # Convert to JSON string for storage in CSV (always set a value, even if empty)
                    $fieldValuesJson = if ($allFieldValues.Count -gt 0) {
                        try {
                            ($allFieldValues | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
                        }
                        catch {
                            Write-Host "  Warning: Could not convert field values to JSON: $_" -ForegroundColor Yellow
                            "{}"
                        }
                    } else {
                        "{}"
                    }
                    
                    $webParts += [pscustomobject]@{
                        SiteUrl = $SiteUrl
                        Library = $Library
                        PageName = $PageName
                        ServerRelativeUrl = $ServerRelativeUrl
                        PageType = $PageType
                        WebPartId = $wp.Id
                        WebPartTitle = $wpTitle
                        WebPartType = $wp.WebPartType
                        ZoneId = $wp.ZoneId
                        ZoneIndex = $wp.ZoneIndex
                        IsClosed = $wp.IsClosed
                        Hidden = $wp.Hidden
                        AllFieldValues = $fieldValuesJson
                    }
                }
            }
        }
        catch {
            # If Get-PnPWebPart doesn't work, try alternative method for classic pages
            Write-Host "  Note: Could not retrieve web parts using standard method, trying alternative..." -ForegroundColor Gray
        }
        
        # Alternative method: Parse the page content for classic web part pages
        # For classic web part pages, web parts are stored in the page's HTML/XML
        if ($webParts.Count -eq 0) {
            try {
                # Get the page content
                $pageContent = Get-PnPFile -Url $ServerRelativeUrl -AsString -Connection $Connection -ErrorAction SilentlyContinue
                
                if ($null -ne $pageContent) {
                    # Look for web part definitions in the page content
                    # Classic web parts are typically defined in XML format within the page
                    # Find complete web part blocks (from <WebPart> to </WebPart>)
                    $webPartBlockPattern = '(?s)<WebPart[^>]*>.*?</WebPart>'
                    $webPartBlocks = [regex]::Matches($pageContent, $webPartBlockPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    
                    if ($webPartBlocks.Count -eq 0) {
                        # Try alternative pattern - sometimes web parts are stored differently
                        $webPartBlockPattern = '(?s)<WebPart[^>]*>.*?</asp:WebPart>'
                        $webPartBlocks = [regex]::Matches($pageContent, $webPartBlockPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    }
                    
                    if ($webPartBlocks.Count -gt 0) {
                        foreach ($blockMatch in $webPartBlocks) {
                            $wpBlockXml = $blockMatch.Value
                            
                            # Try to extract web part type/ID from the XML
                            $wpTypeMatch = [regex]::Match($wpBlockXml, 'TypeName="([^"]*)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                            if (-not $wpTypeMatch.Success) {
                                $wpTypeMatch = [regex]::Match($wpBlockXml, '<TypeName[^>]*>([^<]*)</TypeName>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                            }
                            $wpType = if ($wpTypeMatch.Success) { $wpTypeMatch.Groups[1].Value } else { "Unknown" }
                            
                            # Extract title from this web part block only
                            $wpTitle = "Untitled"
                            $wpTitleMatch = [regex]::Match($wpBlockXml, '<Title[^>]*>([^<]*)</Title>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                            if ($wpTitleMatch.Success) {
                                $wpTitle = $wpTitleMatch.Groups[1].Value.Trim()
                            }
                            else {
                                # Try alternative title attribute format
                                $wpTitleAttrMatch = [regex]::Match($wpBlockXml, 'Title="([^"]*)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                                if ($wpTitleAttrMatch.Success) {
                                    $wpTitle = $wpTitleAttrMatch.Groups[1].Value.Trim()
                                }
                            }
                            
                            # Extract WebPartId if available
                            $wpId = "N/A"
                            $wpIdMatch = [regex]::Match($wpBlockXml, 'ID="([^"]*)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                            if ($wpIdMatch.Success) {
                                $wpId = $wpIdMatch.Groups[1].Value
                            }
                            
                            # Extract properties from XML block
                            $xmlFieldValues = @{}
                            # Try to extract Property elements from XML
                            $propertyMatches = [regex]::Matches($wpBlockXml, '<Property[^>]*Name="([^"]*)"[^>]*>(.*?)</Property>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
                            foreach ($propMatch in $propertyMatches) {
                                if ($propMatch.Groups.Count -ge 3) {
                                    $propName = $propMatch.Groups[1].Value
                                    $propValue = $propMatch.Groups[2].Value.Trim()
                                    if (-not [string]::IsNullOrWhiteSpace($propName)) {
                                        $xmlFieldValues[$propName] = $propValue
                                    }
                                }
                            }
                            
                            # Convert to JSON
                            $xmlFieldValuesJson = if ($xmlFieldValues.Count -gt 0) {
                                try {
                                    ($xmlFieldValues | ConvertTo-Json -Compress -Depth 10 -ErrorAction Stop)
                                }
                                catch {
                                    "{}"
                                }
                            } else {
                                "{}"
                            }
                            
                            $webParts += [pscustomobject]@{
                                SiteUrl = $SiteUrl
                                Library = $Library
                                PageName = $PageName
                                ServerRelativeUrl = $ServerRelativeUrl
                                PageType = $PageType
                                WebPartId = $wpId
                                WebPartTitle = $wpTitle
                                WebPartType = $wpType
                                ZoneId = "N/A"
                                ZoneIndex = "N/A"
                                IsClosed = "N/A"
                                Hidden = "N/A"
                                AllFieldValues = $xmlFieldValuesJson
                            }
                        }
                    }
                }
            }
            catch {
                Write-Host "  Warning: Could not parse page content: $_" -ForegroundColor Yellow
            }
        }
        
        # If still no web parts found, check if it's a wiki page (which might have embedded content)
        if ($webParts.Count -eq 0 -and $PageType -like "*Wiki*") {
            # Wiki pages might not have traditional web parts, but could have embedded content
            $webParts += [pscustomobject]@{
                SiteUrl = $SiteUrl
                Library = $Library
                PageName = $PageName
                ServerRelativeUrl = $ServerRelativeUrl
                PageType = $PageType
                WebPartId = "N/A"
                WebPartTitle = "Wiki Page Content"
                WebPartType = "WikiField"
                ZoneId = "N/A"
                ZoneIndex = "N/A"
                IsClosed = "N/A"
                Hidden = "N/A"
                AllFieldValues = "N/A"
            }
        }
    }
    catch {
        Write-Host "  ERROR: Failed to process page $($PageName): $_" -ForegroundColor Red
        # Add error entry
        $webParts += [pscustomobject]@{
            SiteUrl = $SiteUrl
            Library = $Library
            PageName = $PageName
            ServerRelativeUrl = $ServerRelativeUrl
            PageType = $PageType
            WebPartId = "ERROR"
            WebPartTitle = "Error processing page"
            WebPartType = $_.Exception.Message
            ZoneId = "N/A"
            ZoneIndex = "N/A"
            IsClosed = "N/A"
            Hidden = "N/A"
            AllFieldValues = "N/A"
        }
    }
    
    return $webParts
}

# Process each page
$pageIndex = 0
foreach ($page in $classicPages) {
    $pageIndex++
    Write-Host "[$pageIndex/$totalPages] Processing: $($page.PageName) in $($page.SiteUrl)" -ForegroundColor Green
    
    # Connect to the site if needed
    $conn = Connect-Site -Url $page.SiteUrl
    if ($null -eq $conn) {
        Write-Host "  Skipping page due to connection failure" -ForegroundColor Yellow
        continue
    }
    
    # Get web parts from the page
    $webParts = Get-WebPartsFromPage -SiteUrl $page.SiteUrl `
                                      -ServerRelativeUrl $page.ServerRelative `
                                      -PageName $page.PageName `
                                      -Library $page.Library `
                                      -PageType $page.PageType `
                                      -Connection $conn
    
    if ($webParts.Count -eq 0) {
        Write-Host "  No web parts found on this page $($page.ServerRelative)" -ForegroundColor Gray
        # Still add an entry to show the page was processed
        $webPartsReport.Add([pscustomobject]@{
            SiteUrl = $page.SiteUrl
            Library = $page.Library
            PageName = $page.PageName
            ServerRelativeUrl = $page.ServerRelative
            PageType = $page.PageType
            WebPartId = "None"
            WebPartTitle = "No web parts found"
            WebPartType = "N/A"
            ZoneId = "N/A"
            ZoneIndex = "N/A"
            IsClosed = "N/A"
            Hidden = "N/A"
            AllFieldValues = "N/A"
        })
    }
    else {
        Write-Host "  Found $($webParts.Count) web part(s)" -ForegroundColor Cyan
        foreach ($wp in $webParts) {
            $webPartsReport.Add($wp)
        }
    }
    
    Write-Host ""
}

# Disconnect
if ($null -ne $connection) {
    Disconnect-PnPOnline 
}

# Export results
if ($webPartsReport.Count -gt 0) {
    $webPartsReport `
        | Sort-Object SiteUrl, Library, PageName, WebPartTitle `
        | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`nExport completed!" -ForegroundColor Green
    Write-Host "Total web parts found: $($webPartsReport.Count)" -ForegroundColor Cyan
    Write-Host "Results exported to: $OutputPath" -ForegroundColor Cyan
}
else {
    Write-Host "`nNo web parts found to export." -ForegroundColor Yellow
}

