<#
.SYNOPSIS
    Grants an Entra app Write permission to a SharePoint site.
#>
Param(
    [string]$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/WRK-AboutUs",
    [string]$AppId = "519ad27a-1af4-41e7-87e0-825fb8e368d6",
    [string]$AppDisplayName = "TestSiteSelected",
    [ValidateSet('Read', 'Write')]
    [string]$Permissions = "Write"
)

Import-Module PnP.PowerShell

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId 66a1852a-1f21-46a2-ad58-35fc4c3f1530

Grant-PnPAzureADAppSitePermission `
    -AppId $AppId `
    -DisplayName $AppDisplayName `
    -Permissions $Permissions `
    -Site $SiteUrl

Write-Host "Granted $Permissions to '$AppDisplayName' on $SiteUrl"




Disconnect-PnPOnline

$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432"
$clientsecret = ""
$secureSecret = ConvertTo-SecureString -String $clientsecret -AsPlainText -Force
$credential = [PSCredential]::new($appid, $secureSecret)
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome

#Connect-MgSite -SiteId $SiteUrl

Get-MgContext

$TargetSite = Get-MgSite -SiteId "caje77sharepoint.sharepoint.com:/sites/WRK-AboutUs"

Disconnect-MgGraph