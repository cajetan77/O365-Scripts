#Define Parameters
$AdminSiteUrl = "https://caje77sharepoint-admin.sharepoint.com"
$SiteURL = "https://caje77sharepoint.sharepoint.com/sites/CDocs"

$SiteDesignName = "App4"
$removeSiteDesign = $false
$addSiteDesign=$true
$siteTemplate="Team Site Without Group"
  
if ($removeSiteDesign) {
    $crescentSiteDesign = Get-SPOSiteDesign | Where-Object { $_.Title -eq $SiteDesignName }
    if ($crescentSiteDesign) {
        Remove-SPOSiteDesign $crescentSiteDesign.Id
    }
}
#Connect to SharePoint Admin Center
try {
    Connect-SPOService -Url $AdminSiteUrl
}
catch {
    Write-Host "Error connecting to SharePoint Admin Center: $_"
    exit
}

if($addSiteDesign)
{  
#Get the site schema to a variable
try {
    $SiteSchema = Get-SPOSiteScriptFromWeb -WebURL $SiteURL -IncludeBranding -IncludeTheme -IncludeRegionalSettings -IncludeSiteExternalSharingCapability 
}
catch {
    Write-Host "Error getting site schema: $_"
    exit
}

  
#Add site schema as Site Script 
try {

    $SiteScript = Add-SPOSiteScript -Title $SiteDesignName -Content $SiteSchema
}
catch {
    Write-Host "Error creating site script: $_"
    exit
}
 
try {
    #https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
    switch($siteTemplate)
    {
    
    "Team Site Without Group"{$SiteDesign = Add-SPOSiteDesign -Title $SiteDesignName -WebTemplate 1 -SiteScripts $SiteScript.Id}
     "Communication"{$SiteDesign = Add-SPOSiteDesign -Title $SiteDesignName -WebTemplate 68 -SiteScripts $SiteScript.Id}

}
}
catch {
    Write-Host "Error creating site design: $_"
    exit
}
}

