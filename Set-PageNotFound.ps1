$clientId = "66a1852a-1f21-46a2-ad58-35fc4c3f1530" #Pnp Management Shell
$tenant = "caje77sharepoint"
$site = "WRK-AboutUs"

$siteUrl = "https://$($tenant).sharepoint.com/sites/$($site)"

Write-Host "Setting 404 page at $($siteUrl)..."
#$siteUrl = "https://caje77sharepoint.sharepoint.com/sites/WRK-AboutUs/"
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive

Write-Host "  Disabling NoScript" -ForegroundColor Cyan
Set-PnPTenantSite -Url $siteUrl -NoScriptSite:$false

Set-PnPPropertyBagValue -Key "vti_filenotfoundpage" `
    -Value "/sites/$($site)/SitePages/Page-B.aspx"


# Enable NoScript
Write-Host "  Enabling NoScript" -ForegroundColor Cyan
Set-PnPTenantSite -Url $siteUrl -NoScriptSite

Write-Host "Script Complete! :)" -ForegroundColor Green    