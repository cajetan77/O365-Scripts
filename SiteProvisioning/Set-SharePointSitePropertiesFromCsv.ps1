param(
    [string]$CsvPath = ".\Sites.csv",
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


function Set-Branding {
    param (
        [string]$SiteUrl
    )
    try {
        Set-PnPWebHeader -HeaderLayout "Extended"
        Set-PnPFooter -Layout "Extended"
        Set-PnPWeb -MegaMenuEnabled:$false
        $file = Add-PnpFile -Path $SiteLogoPath -Folder "SiteAssets"
        Set-PnPWeb -SiteLogoUrl $file.ServerRelativeUrl
        $file = Add-PnpFile -Path $SiteThumbnailPath -Folder "SiteAssets"
        Set-PnPWebHeader -SiteThumbnailUrl $file.ServerRelativeUrl
        $file = Add-PnpFile -Path $backgroundImagePath -Folder "SiteAssets"
        $bgUrl = "https://$((Get-PnPWeb).Url.Split('/')[2])$($file.ServerRelativeUrl)"
        Set-PnPWebHeader -HeaderLayout "Extended" -HeaderBackgroundImageUrl $bgUrl -ErrorAction Stop
        Write-Host "Header and Footer Extended  on $($SiteUrl)" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to set branding on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red
        
    }
}


function Set-SearchSettings {
    param (
        [string]$SiteUrl
    )
    try {

        $list = Get-PnPList -Identity "Site Assets" -ErrorAction SilentlyContinue
        if ($list) {
            $list.NoCrawl = $true
            $list.Update()
            Invoke-PnPQuery
            Write-Host "Site Assets list no crawled on $($SiteUrl)" -ForegroundColor Green
        }
        else {
            $web = Get-PnPWeb 
            $web.Lists.EnsureSiteAssetsLibrary()
            Invoke-PnPQuery
            $list = Get-PnPList -Identity "Site Assets" 
            Set-PnPList -Identity $list -NoCrawl:$true
            Invoke-PnPQuery
            Write-Host "Site Assets list no crawled on $($SiteUrl)" -ForegroundColor Green
     
        }

    }
    catch {
        Write-Host "ERROR: Failed to set search settings on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red
       
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
            Write-Host "ERROR: $($library.Title) not found on $SiteUrl" -ForegroundColor Red
           
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
            -RemoveRole "Edit" `
            -AddRole "Contribute" `
            
        Write-Host "Permissions updated: Owners and Members have Contribute on Documents." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to set DocLibraryPermissions: $($_.Exception.Message)" -ForegroundColor Red
       
    }

}


function Get-ContentTypeHub {
    param(
        [string]$ct
    )
    Write-Host "Adding Content Type from the Content Type Hub" -ForegroundColor Green
    $contentTypesArray = $ct.Split(",") | ForEach-Object { $_.Trim() }  
    $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
    Write-Host "Content Type Hub URL: $contentTypeHubUrl" -ForegroundColor Green
    try {
        $ctconnection = Connect-PnPOnline -Url $contentTypeHubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
     
        $ctHub = Get-PnPContentType -Connection $ctconnection
        Disconnect-PnPOnline
        Write-Host "Disconnected from content type hub" -ForegroundColor Green
    }
    
    catch {
        Write-Host "Error connecting to Content Type Hub: $($_.Exception.Message)" -ForegroundColor Red
    }
 
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
       
        
        foreach ($cts in $ctHub) {
            if ($contentTypesArray -contains $cts.Name) {
                Add-PnPContentTypesFromContentTypeHub -ContentTypes $cts.Id -Site $SiteUrl -Connection $siteconnection
                Write-Host "Added content type '$($cts.Name)' to site: $SiteUrl" -ForegroundColor Green
            }
        }
     
    }
    catch {
        Write-Host "Error adding Content Types from Hub: $($_.Exception.Message)" -ForegroundColor Red
    }
    
}


function Add-ContentTypes {
    param(
        [string]$SiteUrl
    )
    try {

        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-Host "ERROR: DocLibrary not found on $SiteUrl" -ForegroundColor Red
            throw "Document library 'Documents' was not found."
            
        }
        else {
            $library.ContentTypesEnabled = $true
            $library.Update()
            Invoke-PnPQuery
            Write-Host "Content types enabled on $($library.Title)" -ForegroundColor Green
            Get-ContentTypeHub -ct $contentTypeName
            #Get the content type
            $ContentType = Get-PnPContentType -Identity $contentTypeName
            If ($ContentType) {
                #Add Content Type to Library
                foreach ($listName in $contentTypeList) {  
                    Add-PnPContentTypeToList -List $listName -ContentType $ContentType
                   
                    Write-Host "Added content type '$($ContentType.Name)' to list '$($listName)'" -ForegroundColor Green
                    Set-DefaultContentType -SiteUrl $SiteUrl -ListName $listName -ContentTypeName $contentTypeName
                }
            }
        }
    }
    catch {
        Write-Host "ERROR: Failed to add content type: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Set-DefaultContentType {
    param(
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ContentTypeName
    )

    try {
        Set-PnPDefaultContentTypeToList -List $ListName -ContentType $ContentTypeName -ErrorAction Stop
        Write-Host "Default content type set to '$ContentTypeName' on list '$ListName' ($SiteUrl)" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to set default content type '$ContentTypeName' on list '$ListName' ($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red
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



function Install-App {
    param(
        [string]$SiteUrl
    )
    try {
      
        foreach ($appId in $appIds) {
            $app = Get-PnPApp -Identity $appId  -ErrorAction Stop
            if ($null -eq $app.InstalledVersion) {
                Install-PnPApp -Identity $app  -ErrorAction Stop
                Write-Host "App installed: $($app.Title)" -ForegroundColor Green
            }
            else {
                Write-Host "App already installed: $($app.Title)" -ForegroundColor Yellow
            }
           
        }
    }
    catch {
        Write-Host "ERROR: Failed to install app on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red
    }
}


function Add-PageTemplates {
    param (
        [string] $siteUrl
    )
    
    foreach ($pageTempalte in $pageTempaltes) {
        try {
           

            if (Get-PnPPage -Identity $pageTempalte -ErrorAction SilentlyContinue) {
                Write-Host "Page template already exists: $($pageTempalte). Skipping creation." -ForegroundColor Yellow
            }
            else {
              


                if ($pageTempalte -eq "Landing-Page") {

                    $page = Add-PnpPage -Name $pageTempalte 

                    $component = Get-PnPPageComponent -Page $page.Name -ListAvailable | Where-Object { $_.Name -eq "PnP - Search Results" }
                   
                    $sections = Get-PnPPage -Identity $page.Name | Select-Object -ExpandProperty Sections -ErrorAction SilentlyContinue
                    if (-not $sections -or $sections.Count -eq 0) {
                        Add-PnPPageSection -Page $page.Name -SectionTemplate OneColumn -ErrorAction Stop
                    }
                    $config = Get-Content -Raw -Path "webpartproperties.json"
                    Add-PnPPageWebPart -Page $page.Name  -Component $component -Section 1 -Column 1 -WebPartProperties $config -ErrorAction Stop
                    Set-PnPPage -Identity $page.Name -PromoteAs Template -Publish -ErrorAction Stop
                }

                else {
                    $page = Add-PnpPage -Name $pageTempalte -PromoteAs Template -Publish
                }
                Write-Host "Page template added: $($pageTempalte)" -ForegroundColor Green
            }
            
        }
        catch {
            Write-Host "ERROR: Failed to add page template: $($pageTempalte): $($_.Exception.Message)" -ForegroundColor Red
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

                Write-Host "Site column '$ColumnName'  exists on $($siteUrl). Skipping creation." -ForegroundColor Yellow        
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
                    Write-Host "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." -ForegroundColor Yellow
                }
                else {
                    switch ($columnName) {
                        "Main Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-Host "Added site column '$columnName' to '$list'"
                        }
                        "Sub Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-Host "Added site column '$columnName' to '$list'"
                        }
                        Default {
                            Add-PnPField -List $list -Field $existingColumn
                            Write-Host "Added site column '$columnName' to '$list'"
                        }
                    }
                    
                    Write-Host "Added existing site column '$ColumnName' to Documents library on $($siteUrl)." -ForegroundColor Green
                }
            } 
            else {
                Write-Host "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." -ForegroundColor Yellow
            }

        }

    }
    catch {
        Write-Host "ERROR: Failed to add site column on $($siteUrl): $($_.Exception.Message)" -ForegroundColor Red
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
                Write-Host "ERROR: DocLibrary not found on $SiteUrl" -ForegroundColor Red
           
                
            }
            else {
            
                $view = Get-PnPView -List $library  | Where-Object { $_.DefaultView -eq $true }
                Set-PnPView -List $library -Identity $view.Id -Fields "DocIcon", "Title", "Modified", "Editor", "ReviewDate1", "DoogleWFMainCategory", "MSDNotificationSent", "DoogleWFRestrictedApproval", "DoogleWFSubCategory" -ErrorAction Stop
                Write-Host "Custom view created and set as default on $($library.Title)" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "ERROR: Failed to set views on $($SiteUrl): $($_.Exception.Message)" -ForegroundColor Red
        
    }

    
    
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



    foreach ($row in $rows) {
        $index++
        $siteUrl = if ($row.SiteUrl) { $row.SiteUrl.ToString().Trim() } else { "" }

        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            Write-Host "[$index/$total] Skipping row because required 'SiteUrl' is empty." -ForegroundColor Yellow
            continue
        }

        try {
        
            Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint -ErrorAction Stop    

            Add-PageTemplates -SiteUrl $SiteUrl  
            Add-SiteToHubAssociation -SiteUrl $siteUrl -TargetHubSiteUrl $HubSiteUrl
            Set-SiteRegionalSettings -SiteUrl $siteUrl
            Set-SearchSettings -SiteUrl $siteUrl
            Set-DocLibraryPermissions -SiteUrl $siteUrl
            Add-GroupstoSharePointGroups -SiteUrl $siteUrl
            Set-Branding -SiteUrl $siteUrl
            Install-App -SiteUrl $SiteUrl
            Add-ContentTypes -SiteUrl $SiteUrl
            Add-SiteColumns -SiteUrl $siteUrl
           
            Set-Views -SiteUrl $siteUrl    
            
        }
        catch {
            $failed++
            Write-Host "[$index/$total] ERROR: Failed to associate '$siteUrl' to hub '$HubSiteUrl'. $($_.Exception.Message)" -ForegroundColor Red
        }
    }

 
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

