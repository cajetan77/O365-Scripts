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
$MenuItems = @()

$ExcludedLists = @("Site Assets", "Site Pages", "Form Templates", "Style Library")

$contentTypesArray = $config.ContentTypeName
$ExcludedLists = @("Site Assets", "Site Pages", "Form Templates", "Style Library")

function Get-DefaultMenuItems() {
   
    # Returns a list of default menu items with their visibility settings.
    $DefaultMenuItems = @()
   
    $DefaultMenuItems += Add-MenuItem -title:"Word document" -templateId:"NewDOC" -visible:$false
    $DefaultMenuItems += Add-MenuItem -title:"Excel workbook" -templateId:"NewXSL" -visible:$false
    $DefaultMenuItems += Add-MenuItem -title:"PowerPoint presentation" -templateId:"NewPPT" -visible:$false
    $DefaultMenuItems += Add-MenuItem -title:"OneNote notebook" -templateId:"NewONE" -visible:$true
    $DefaultMenuItems += Add-MenuItem -title:"Visio drawing" -templateId:"NewVSDX" -visible:$true
    $DefaultMenuItems += Add-MenuItem -title:"Forms for Excel" -templateId:"NewXSLForm" -visible:$false
    $DefaultMenuItems += Add-MenuItem -title:"Link" -templateId:"Link" -visible:$false
 
    return $DefaultMenuItems
}
 


 
function Add-MenuItem() {
    param(
        [Parameter(Mandatory)]
        $title,
        [Parameter(Mandatory)]
        $visible,
        [Parameter(Mandatory)]
        $templateId,
        [Parameter()]
        $contentTypeId
    )
 
    $newChildNode = New-Object System.Object
    $newChildNode | Add-Member -type NoteProperty -name title -value:$title
    $newChildNode | Add-Member -type NoteProperty -name visible -value:$visible
    $newChildNode | Add-Member -type NoteProperty -name templateId -value:$templateId
    if ($null -ne $contentTypeId) {
        $newChildNode | Add-Member -type NoteProperty -name contentTypeId -value:$contentTypeId
        $newChildNode | Add-Member -type NoteProperty -name isContentType -value:$true
    }
 
    return $newChildNode
}
 
 
function Set-MenuOptions {
    param(
        [Parameter(Mandatory = $false)]
        $library )
    
  
    $listContentTypes = Get-PnPContentType -List $Library
    $defaultView = Get-PnpView -List $library | Where-Object { $_.DefaultView -eq $true }
    $listContentTypes | ForEach-Object {
        $ct = $PSItem
        if ($ct.Name -eq "Folder") {
            $MenuItems += Add-MenuItem -title:"Folder" -templateId:"NewFolder" -visible:$true
            Write-Output ("Added menu item for content type '$($ct.Name)' in library '$($library.Title)' in site '$($siteUrl)'")
            return
        }
          
        if ($libcontentTypesArray -contains $ct.Name ) {
            $MenuItems += Add-MenuItem -title: $ct.Name -visible:$true -templateId:$ct.StringId -contentTypeId:$ct.StringId
            Write-Output ("Added menu item for content type '$($ct.Name)' in library '$($library.Title)' in site '$($siteUrl)'")
            return
        }
           
        $MenuItems += Add-MenuItem -title:$ct.Name -visible:$false -templateId:$ct.StringId -contentTypeId:$ct.StringId
           
    }
    $MenuItems += Get-DefaultMenuItems
    $defaultView.NewDocumentTemplates = $menuItems | ConvertTo-Json
    $defaultView.Update()
    Invoke-PnPQuery
    Write-Output ("Updated menu items for library '$($library.Title)' in site '$($siteUrl)'")
    $MenuItems = @()
        
}
 

    


#Connect to SharePoint Online site

$sites = Import-Csv -Path ".\sites.csv"
foreach ($site in $sites) {
    $siteUrl = $site.URL
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId
        Write-Output ("Connected to site:$($siteUrl)") 
        $libcontentTypesArray = $contentTypesArray.Split(",")
        $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and ($ExcludedLists -notcontains $_.Title) }
        foreach ($library in $libraries) {
            foreach ($contentType in $libcontentTypesArray) {
           
                $ctInList = Get-PnPContentType -List $library.Id | Where-Object { $_.Name -eq $contentType } 
                if ($null -ne $ctInList) {
                    #Write-Output "Content type '$contentType' already exists in library: $($library.Title). Skipping addition."
                    continue
                }
                else {
                    
                    # Write-Output "Content type '$contentType' does not exist in library: $($library.Title). Adding it now."
                    #Add-PnPContentTypeToList -List $library.Id -ContentType $contentType
                    #Write-Output "Added content type '$contentType' to library: $($library.Title)" 
                }  
            
            }
            Set-MenuOptions -library $library
    

        }
    }
    catch {
        Write-Error "Failed to connect to SharePoint  using PnP: $_"
        continue
    }
}
       
   


# Get-ContentTypeHub -ct $contentTypesArray


Disconnect-PnPOnline