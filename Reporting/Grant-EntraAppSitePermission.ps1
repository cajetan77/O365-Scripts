<#
.SYNOPSIS
    Grants an Entra app Write permission to a SharePoint site.
#>
Param(
    [string]$SiteUrl = "https://sharepoint.com",
    [string]$AppId = "YOUR-APP-CLIENT-ID",
    [string]$AppDisplayName = "Your App Name",
    [ValidateSet('Read', 'Write')]
    [string]$Permissions = "Write"
)

Import-Module PnP.PowerShell

Connect-PnPOnline -Url $SiteUrl -Interactive

Grant-PnPAzureADAppSitePermission `
    -AppId $AppId `
    -DisplayName $AppDisplayName `
    -Permissions $Permissions `
    -Site $SiteUrl

Write-Host "Granted $Permissions to '$AppDisplayName' on $SiteUrl"

Disconnect-PnPOnline
