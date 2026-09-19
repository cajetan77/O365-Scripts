Connect-PnPOnline -Url "https://caje77sharepoint.sharepoint.com/sites/CajT21" -Interactive -ClientId "66a1852a-1f21-46a2-ad58-35fc4c3f1530"

$web = Get-PnPWeb -Includes Navigation


$env:NAV_NODES = '[{"Title":"Pages","Url":"{siteUrl}/SitePages"},{"Title":"Home","Url":"#"}]'
#[{"Title":"Pages","Url":"{siteUrl}/SitePages"},{"Title":"Home","Url":"#"}]


$rawNodes = $env:NAV_NODES

# Swap the placeholder with the dynamic variable
$dynamicJson = $rawNodes -replace '\{siteUrl\}', $web.Url

# Rehydrate into PowerShell objects
$Nodes = ConvertFrom-Json -InputObject $dynamicJson


#$Nodes = @(@{"Title" = "Pages"; "Url" = "{SiteUrl}/SitePages" }, @{"Title" = "Home"; "Url" = "#" })
$Nodes | ForEach-Object { $_.Url = $_.Url -replace "{SiteUrl}", $web.Url }
#Remove-PnPNavigationNode -Location QuickLaunch -Force
Get-PnPNavigationNode -Location QuickLaunch | Remove-PnPNavigationNode -Force
$Nodes | ForEach-Object { Add-PnPNavigationNode -Location QuickLaunch -Title $_.Title -Url $_.Url }

#How do I Put this is in enviornment variable so that I can run it from the command line?
$Env:RemoveSHarePointNav = $true

function Remove-SHarePointNav {
    $Env:RemoveSHarePointNav = $true
    Connect-PnPOnline -Url "https://caje77sharepoint.sharepoint.com/sites/BoardPortal" -Interactive -ClientId "66a1852a-1f21-46a2-ad58-35fc4c3f1530"
    $web = Get-PnPWeb -Includes Navigation
    $Nodes = @()
    $Nodes += @{
        Title = 'Pages'
        Url   = $web.Url + '/SitePages'
    }
}

Remove-SHarePointNav

#What would be the best way to define environment variables in a function app Nodes array?


