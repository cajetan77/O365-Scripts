param(
    [string]$CsvPath = ".\Sites.csv",
    [string]$LogDirectory = ".\logs",
    [string]$ConfigPath = ".\config.json",
    [string]$SiteLogoPath = "D:\Downloads\pexels-padrinan-255379.jpg",
    [string]$SiteThumbnailPath = "D:\Downloads\pexels-padrinan-255379.jpg",
    [string]$backgroundImagePath = "D:\Downloads\pexels-padrinan-255379.jpg",
    [string]$HubSiteUrl = "https://caje77sharepoint.sharepoint.com/sites/AIALIntranet",
    $appIds = @("59903278-DD5D-4E9E-BEF6-562AAE716B8B", "00406271-0276-406F-9666-512623EB6709"),
    $pageTempaltes = @("Landing-Page", "Page-Template-1", "Page-Template-2"),
    $contentTypeName = "Doogle content category page",
    $contentTypeList = @("Site Pages"),
    $viewsList = @("Documents", "Site Pages")
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
    Write-Host "ERROR: HubSiteUrl is required. Example: -HubSiteUrl 'https://tenantname.sharepoint.com/sites/hub'" -ForegroundColor Red
    exit 1
}

$HubSiteUrl = $HubSiteUrl.ToString().Trim()



function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info",
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$LogFilePath = ".\logs",
        [string]$siteName

    )
    # Create the log directory if it doesn't exist
    if (!(Test-Path $LogFilePath)) {
        New-Item -Path $LogFilePath -ItemType Directory -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $siteName = $siteName.ToString().Trim()
    $logFile = Join-Path $LogFilePath "log-$($siteName).txt"
    $logLine = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    switch ($Level) {
        "Error" {
            Add-Content -Path $logFile -Value $logLine
            Write-Error $logLine
        }
        "Warning" {
            Add-Content -Path $logFile -Value $logLine
            Write-Warning $logLine
        }
        "Success" {
            Add-Content -Path $logFile -Value $logLine
            Write-Host $logLine -ForegroundColor Green
            
        }
        default {
            # Info: visible in console and captured by Start-Transcript
            Add-Content -Path $logFile -Value $logLine
            Write-Host $logLine -ForegroundColor Gray
        }
    }
}


function Set-SiteRegionalSettings {
    param(
        [string]$SiteUrl,
        [string]$siteName 
    )
 
    try {
        $web = Get-PnpWeb -Includes RegionalSettings.LocaleId, RegionalSettings.TimeZones -Connection $siteconnection
        $localeId = 5129 # New Zealand English
        $web.RegionalSettings.LocaleId = $localeId
        $web.Update()
        Invoke-PnPQuery
        Write-LogMessage -Message "Updated Site Regional Settings to have NZ Time Zone and NZ Locale  $($web.Url)" -Level Info -siteName $siteName
    }
    catch {
        Write-LogMessage -Message "Error connecting to site $SiteUrl :$($_.Exception.Message)" -Level Error -siteName  $siteName 
    
    }
    

}  




function Add-GroupstoSharePointGroups {
    param (
        [string]$SiteUrl,
        [string]$siteName
    )

    try {
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $EntraGroupObjectId = "b4fa1e98-2893-4f26-bc64-f7a3e93b3753" # Replace with actual Entra Group Object ID
        $groupLoginName = "c:0t.c|tenant|$EntraGroupObjectId"
        Add-PnPGroupMember `
            -Group $ownersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-LogMessage -Message "Added Entra group with Object ID $EntraGroupObjectId to Owners group on $($SiteUrl)" -Level Success -siteName $siteName
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop
        Add-PnPGroupMember `
            -Group $membersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-LogMessage -Message "Added Entra group with Object ID $EntraGroupObjectId to Members group on $($SiteUrl)" -Level Success -siteName $siteName
    
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to add group to SharePoint group on $($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
    }
    
}


function Set-Branding {
    param (
        [string]$SiteUrl,
        [string]$siteName
    )
    try {
        # 1. Make sure layout is Extended
        Set-PnPWebHeader -HeaderLayout Extended

        # 2. Upload image
        $file = Add-PnPFile -Path $backgroundImagePath -Folder "SiteAssets"

        # 3. Use URL encoded server-relative URL
        $url = [System.Uri]::EscapeUriString($file.ServerRelativeUrl)

        # 4. Apply in one call
        Set-PnPWebHeader `
            -HeaderLayout Extended `
            -HeaderBackgroundImageUrl $url `
            -HeaderBackgroundImageFocalX 0.5 `
            -HeaderBackgroundImageFocalY 0.5
        #$bgUrl = "/sites/WRK-Test4/SiteAssets/pexels-padrinan-255379.jpg"

        $a = Set-PnPWebHeader -HeaderBackgroundImageUrl $bgUrl -HeaderLayout Extended  -ErrorAction Stop
        Set-PnPWebHeader -HeaderLayout "Extended"
        Set-PnPFooter -Layout "Extended"
        Set-PnPWeb -MegaMenuEnabled:$false
        $file = Add-PnpFile -Path $SiteLogoPath -Folder "SiteAssets"
        Set-PnPWeb -SiteLogoUrl $file.ServerRelativeUrl
        $file = Add-PnpFile -Path $SiteThumbnailPath -Folder "SiteAssets"
        Set-PnPWebHeader -SiteThumbnailUrl $file.ServerRelativeUrl
       
      
        Write-LogMessage -Message "Header and Footer Extended  on $($SiteUrl)" -Level Success -siteName $siteName
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to set branding on $($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
        
    }
}


function Set-SearchSettings {
    param (
        [string]$SiteUrl,
        [string]$siteName
    )
    try {

        $list = Get-PnPList -Identity "Site Assets" -ErrorAction SilentlyContinue
        if ($list) {
            $list.NoCrawl = $true
            $list.Update()
            Invoke-PnPQuery
            Write-LogMessage -Message "Site Assets list no crawled on $($SiteUrl)" -Level Success -siteName $siteName
        }
        else {
            $web = Get-PnPWeb 
            $web.Lists.EnsureSiteAssetsLibrary()
            Invoke-PnPQuery
            $list = Get-PnPList -Identity "Site Assets" 
            Set-PnPList -Identity $list -NoCrawl:$true
            Invoke-PnPQuery
            Write-LogMessage -Message "Site Assets list no crawled on $($SiteUrl)" -Level Success -siteName $siteName

        }

    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to set search settings on $($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
       
    }
}


Function Set-DocLibraryPermissions {
    param(
        [string]$SiteUrl,
        [string]$siteName
    )
    try {
        Write-Host "Setting DocLibraryPermissions on $SiteUrl" -ForegroundColor Yellow
        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-LogMessage -Message "ERROR: $($library.Title) not found on $SiteUrl" -Level Error -siteName $siteName
           
        }

        Set-PnPList -Identity $library -BreakRoleInheritance -CopyRoleAssignments -ErrorAction Stop

        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop

        if (-not $ownersGroup -or -not $membersGroup) {
            Write-LogMessage -Message "ERROR: Could not resolve Owners or Members group for $SiteUrl." -Level Error -siteName $siteName
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
            -RemoveRole "Edit" `
            -AddRole "Contribute" `
            
        Write-LogMessage -Message "Permissions updated: Owners and Members have Contribute on Documents." -Level Success -siteName $siteName
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to set DocLibraryPermissions: $($_.Exception.Message)" -Level Error -siteName $siteName
       
    }

}


function Get-ContentTypeHub {
    param(
        [string]$ct,
        [string]$siteName
    )
    Write-LogMessage -Message "Adding Content Type from the Content Type Hub" -Level Info -siteName $siteName
    $contentTypesArray = $ct.Split(", ") | ForEach-Object { $_.Trim() }  
    $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
    Write-LogMessage -Message "Content Type Hub URL: $contentTypeHubUrl" -Level Info -siteName $siteName
    try {
        $ctconnection = Connect-PnPOnline -Url $contentTypeHubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
     
        $ctHub = Get-PnPContentType -Connection $ctconnection
        Disconnect-PnPOnline
        Write-LogMessage -Message "Disconnected from content type hub" -Level Success -siteName $siteName
    }
    
    catch {
        Write-LogMessage -Message "Error connecting to Content Type Hub: $($_.Exception.Message)" -Level Error -siteName $siteName
    }
 
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
       
        
        foreach ($cts in $ctHub) {
            if ($contentTypesArray -contains $cts.Name) {
                Add-PnPContentTypesFromContentTypeHub -ContentTypes $cts.Id -Site $SiteUrl -Connection $siteconnection
                Write-LogMessage -Message "Added content type '$($cts.Name)' to site: $SiteUrl" -Level Success -siteName $siteName
            }
        }
     
    }
    catch {
        Write-LogMessage -Message "Error adding Content Types from Hub: $($_.Exception.Message)" -Level Error -siteName $siteName
    }
    
}


function Add-ContentTypes {
    param(
        [string]$SiteUrl,
        [string]$siteName
    )
    try {

        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-LogMessage -Message "ERROR: DocLibrary not found on $SiteUrl" -Level Error -siteName $siteName
            
            
        }
        else {
            $library.ContentTypesEnabled = $true
            $library.Update()
            Invoke-PnPQuery
            Write-LogMessage -Message "Content types enabled on $($library.Title)" -Level Success -siteName $siteName
            Get-ContentTypeHub -ct $contentTypeName  -siteName $siteName
            #Get the content type
            $ContentType = Get-PnPContentType -Identity $contentTypeName
            If ($ContentType) {
                #Add Content Type to Library
                foreach ($listName in $contentTypeList) {  
                    Add-PnPContentTypeToList -List $listName -ContentType $ContentType
                   
                    Write-LogMessage -Message "Added content type '$($ContentType.Name)' to list '$($listName)'" -Level Success -siteName $siteName
                    Set-DefaultContentType -SiteUrl $SiteUrl -ListName $listName -ContentTypeName $contentTypeName -siteName $siteName
                }
            }
        }
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to add content type: $($_.Exception.Message)" -Level Error -siteName $siteName
        
    }
}

function Set-DefaultContentType {
    param(
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ContentTypeName,
        [string]$siteName
    )

    try {
        Set-PnPDefaultContentTypeToList -List $ListName -ContentType $ContentTypeName -ErrorAction Stop
        Write-LogMessage -Message "Default content type set to '$ContentTypeName' on list '$ListName' ($SiteUrl)" -Level Success -siteName $siteName
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to set default content type '$ContentTypeName' on list '$ListName' ($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
        
    }
}

function Add-SiteToHubAssociation {
    param(
        [string]$SiteUrl,
        [string]$TargetHubSiteUrl,
        [string]$SiteName
    )

    try {
       
        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $TargetHubSiteUrl -ErrorAction Stop
        Write-LogMessage -Message "Site $SiteUrl added to hub $TargetHubSiteUrl" -Level Success -siteName $SiteName
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to add site to hub association: $($_.Exception.Message)" -Level Error -siteName $SiteName
    }
}



function Install-App {
    param(
        [string]$SiteUrl,
        [string]$siteName
    )
    try {
      
        foreach ($appId in $appIds) {
            $app = Get-PnPApp -Identity $appId  -ErrorAction Stop
            if ($null -eq $app.InstalledVersion) {
                Install-PnPApp -Identity $app  -ErrorAction Stop
                Write-LogMessage -Message "App installed: $($app.Title)" -Level Success  -siteName $siteName    
            }
            else {
                Write-LogMessage -Message "App already installed: $($app.Title)" -Level Warning -siteName $siteName
            }
           
        }
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to install app on $($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
    }
}



function SavePageTemplate {
    param (
        [string] $siteUrl,
        [string] $siteName
    )
    try {
        $templatesiteCollection = "https://caje77sharepoint.sharepoint.com/sites/WRK-AboutUs"
        $saveTemplateLocation = $PSScriptRoot
        $templateName = "template"
        
        Connect-PnPOnline -Url $templatesiteCollection -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint
        $siteTemplate = Get-PnPSiteTemplate -IncludeAllClientSidePages -Handlers Pages, PageContents -OutputInstance 
        
         $pagesTemplate = New-PnPSiteTemplate
        foreach ($page in $siteTemplate.ClientSidePages) {
            if ($page.PromoteAsTemplate -eq $true) {
                $pagesTemplate.ClientSidePages.Add($page)
            }
        }
        
        Save-PnPSiteTemplate -Template $pagesTemplate -Out ("{0}{1}.xml" -f $saveTemplateLocation, $templateName)
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to save page template on $($siteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
    }

    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint
    Invoke-PnPSiteTemplate -Path ("{0}{1}.xml" -f $saveTemplateLocation, $templateName)
    Write-LogMessage -Message "Page template added: $($pageTempalte)" -Level Success  -siteName $siteName  
}



function Add-PageTemplates {
    param (
        [string] $siteUrl,
        [string] $siteName
    )
    
    foreach ($pageTempalte in $pageTempaltes) {
        try {
        
              


            if ($pageTempalte -eq "Landing-Page") {

                $pageFileName = "$pageTempalte.aspx"
                $pageItems = Get-PnPListItem -List "Site Pages" -Fields "FileLeafRef", "FileDirRef" -PageSize 2000 -ErrorAction Stop |
                Where-Object { $_["FileLeafRef"] -eq $pageFileName }
                foreach ($item in $pageItems) {
                    Remove-PnPListItem -List "Site Pages" -Identity $item.Id -Force -ErrorAction Stop
                }

                $page = Add-PnpPage -Name $pageTempalte -ErrorAction Stop
                $component = Get-PnPPageComponent -Page $page.Name -ListAvailable | Where-Object { $_.Name -eq "PnP - Search Results" }

                Add-PnPPageSection `
                    -Page $page.Name `
                    -SectionTemplate TwoColumnRight `
                    -Order 1 `
                    -ErrorAction Stop

                # Keep the left side empty on the template. This avoids text controls
                # being added to full-width sections on previously customized pages.

                if (-not $component) {
                    throw "PnP - Search Results web part is not available on page '$($page.Name)'."
                }

                $webPartConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "webpartproperties.json"
                $config = Get-Content -Raw -Path $webPartConfigPath
                Add-PnPPageTextPart `
                    -Page $page.Name `
                    -Section 1 `
                    -Column 1 `
                    -Text "<p>Hello from PowerShell.</p>" `
                    -ErrorAction Stop
                Add-PnPPageWebPart -Page $page.Name -Component $component -Section 1 -Column 2 -WebPartProperties $config -ErrorAction Stop
                
                Add-PnPPageSection `
                    -Page $page.Name `
                    -SectionTemplate TwoColumnRight `
                    -Order 2 `
                    -ErrorAction Stop  
                $component = Get-PnPPageComponent -Page $page.Name -ListAvailable | Where-Object { $_.Name -eq "PnP - Search Results" }

                Add-PnPPageSection `
                    -Page $page.Name `
                    -SectionTemplate TwoColumnRight `
                    -Order 1 `
                    -ErrorAction Stop

                # Keep the left side empty on the template. This avoids text controls
                # being added to full-width sections on previously customized pages.

                if (-not $component) {
                    throw "PnP - Search Results web part is not available on page '$($page.Name)'."
                }

                $webPartConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "webpartproperties.json"
                $config = Get-Content -Raw -Path $webPartConfigPath
                Add-PnPPageTextPart `
                    -Page $page.Name `
                    -Section 2 `
                    -Column 1 `
                    -Text "<p>Hello from PowerShell2.</p>" `
                    -ErrorAction Stop
                Add-PnPPageWebPart -Page $page.Name -Component $component -Section 2 -Column 2 -WebPartProperties $config -ErrorAction Stop



                Set-PnPPage -Identity $page.Name -PromoteAs Template -Publish -ErrorAction Stop

                # SharePoint keeps both /SitePages/<name>.aspx and /SitePages/Templates/<name>.aspx
                # after template promotion. Remove the normal page so only the template remains.
                $promotedItems = Get-PnPListItem -List "Site Pages" -Fields "FileLeafRef", "FileDirRef" -PageSize 2000 -ErrorAction Stop |
                Where-Object { $_["FileLeafRef"] -eq $pageFileName }
                $normalPageItems = @(
                    $promotedItems | Where-Object { $_["FileDirRef"] -notmatch "/SitePages/Templates$" }
                )
                foreach ($normalItem in $normalPageItems) {
                    Remove-PnPListItem -List "Site Pages" -Identity $normalItem.Id -Force -ErrorAction Stop
                }
            }

            else {
                $page = Add-PnpPage -Name $pageTempalte -PromoteAs Template -Publish
            }
            Write-LogMessage -Message "Page template added: $($pageTempalte)" -Level Success  -siteName $siteName  
            
            
        }
        catch {
            Write-LogMessage -Message "ERROR: Failed to add page template: $($pageTempalte): $($_.Exception.Message)" -Level Error -siteName $siteName
        }

    }
}


function Add-SiteColumns {
    param (
        [string] $siteUrl
    )
    try {
        $list = Get-PnPList -Identity "Documents"
        $columnNames = @("Main Category", "Review Date", "Notification Sent", "Sub Category", "Restricted Approval")
        foreach ($ColumnName in $columnNames) {
            $existingColumn = Get-PnPField -Identity $ColumnName -ErrorAction SilentlyContinue
            if ($existingColumn) {

                Write-LogMessage -Message "Site column '$ColumnName'  exists on $($siteUrl). Skipping creation." -Level Warning -siteName $siteName
                switch ($columnName) {
                    "Main Category" { 
                        #Write-Host "Site column '$ColumnName' already exists on $($siteUrl). Skipping adding to list." -ForegroundColor Yellow
                        #Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                        #Write-Host "Added site column '$columnName' to '$listTitle'"
                    }
                    "Review Date" {
                        $existingColumn.DefaultFormula = "=TODAY()+365"
                        $existingColumn.UpdateAndPushChanges($true)
                        Invoke-PnPQuery
                    }
                }
                $fieldInList = Get-PnPField -List $list -Identity $ColumnName -ErrorAction SilentlyContinue
                if ($fieldInList) {
                    Write-LogMessage -Message "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." -Level Warning -siteName $siteName
                }
                else {
                    switch ($columnName) {
                        "Main Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-LogMessage -Message "Added site column '$columnName' to '$list'" -Level Success -siteName $siteName
                        }
                        "Sub Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-LogMessage -Message "Added site column '$columnName' to '$list'" -Level Success -siteName $siteName
                        }
                        Default {
                            Add-PnPField -List $list -Field $existingColumn
                            Write-LogMessage -Message "Added site column '$columnName' to '$list'" -Level Success -siteName $siteName
                        }
                    }
                    
                    Write-LogMessage -Message "Added existing site column '$ColumnName' to Documents library on $($siteUrl)." -Level Success -siteName $siteName
                }
            } 
            else {
                Write-LogMessage -Message "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." -Level Warning -siteName $siteName
            }

        }

    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to add site column on $($siteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
    }
    
}



function Set-Views {
    param(
        [string]$SiteUrl
    )
    try {
        
        foreach ($list in $viewsList) {
        
            $library = Get-PnPList -Identity $list -ErrorAction Stop
            if (-not $library) {
                Write-LogMessage -Message "ERROR: DocLibrary not found on $SiteUrl" -Level Error -siteName $siteName
           
                
            }
            else {
            
                $view = Get-PnPView -List $library  | Where-Object { $_.DefaultView -eq $true }
                Set-PnPView -List $library -Identity $view.Id -Fields "DocIcon", "Title", "Modified", "Editor", "ReviewDate1", "DoogleWFMainCategory", "MSDNotificationSent", "DoogleWFRestrictedApproval", "DoogleWFSubCategory" -ErrorAction Stop
                Write-LogMessage -Message "Custom view created and set as default on $($library.Title)" -Level Success -siteName $siteName
            }
        }
    }
    catch {
        Write-LogMessage -Message "ERROR: Failed to set views on $($SiteUrl): $($_.Exception.Message)" -Level Error -siteName $siteName
        
    }

    
    
}


try {
    $rows = Import-Csv -Path $CsvPath -Encoding UTF8
 

    
    $index = 0
    $total = $rows.Count
    $failed = 0

    foreach ($row in $rows) {
        $index++
        $siteUrl = if ($row.SiteUrl) { $row.SiteUrl.ToString().Trim() } else { "" }
        $siteName = if ($row.Title) { $row.Title.ToString().Trim() } else { "" }
        

     
        try {

            #Write-Host "[$index / $total] Site '$siteNameForLog' — transcript: $logPath" -ForegroundColor Cyan

            Write-LogMessage -Message "[$index/$total] Site '$siteUrl'" -Level Info -siteName $siteName
            Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint -ErrorAction Stop

            # Add-SiteToHubAssociation -SiteUrl $siteUrl -TargetHubSiteUrl $HubSiteUrl -SiteName $siteName
            #Set-SiteRegionalSettings -SiteUrl $siteUrl -siteName $siteName
            #Set-SearchSettings -SiteUrl $siteUrl -siteName $siteName
            #Set-DocLibraryPermissions -SiteUrl $siteUrl -siteName $siteName
            #Add-GroupstoSharePointGroups -SiteUrl $siteUrl -siteName $siteName
            #Set-Branding -SiteUrl $siteUrl  -siteName $siteName
            #Install-App -SiteUrl $SiteUrl -siteName $siteName
            #Add-ContentTypes -SiteUrl $SiteUrl  -siteName $siteName
            #Add-SiteColumns -SiteUrl $siteUrl   -siteName $siteName
            #Add-PageTemplates -SiteUrl $SiteUrl -siteName $siteName
            #Set-Views -SiteUrl $siteUrl     -siteName $siteName
            SavePageTemplate -SiteUrl $SiteUrl -siteName $siteName
        }
        catch {
            $failed++
            Write-LogMessage -Message "[$index / $total] ERROR: Failed   $($_.Exception.Message)" -Level Error -siteName $siteName
        }
        finally {
            Write-LogMessage -Message "[$index / $total] Site '$siteUrl' processed successfully" -Level Success -siteName $siteName
        }
    }

 
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

