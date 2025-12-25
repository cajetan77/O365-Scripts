<#
.SYNOPSIS
    Gets the folder structure of all libraries in all sites and exports the results to a CSV file.
.DESCRIPTION
    This script gets the folder structure of all libraries in all sites and exports the results to a CSV file.
.PARAMETER OutputCsv
    The path to the output CSV file.
.PARAMETER ExcludedLists
    The lists to exclude from the folder structure.
#>


param(
    [string]$OutputCsv = ".\FolderStructure.csv",
    [string[]]$ExcludedLists = @("Site Assets", "Site Pages", "SitePages", "Form Templates", "Style Library")
)


$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$clientId = $config.SharePointReportingAppId
$Thumbprint = $config.ThumbPrint

$sites=Import-Csv -Path ".\sites.csv"
foreach ($site in $sites) {
   
    # Connect (interactive)
try {
    Connect-PnPOnline -Url $Site.Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
}
catch {
    Write-Host "Error connecting to SharePoint: $_"
    continue
}

# Get site's server-relative URL to remove from paths
$web = Get-PnPWeb
$siteServerRelativeUrl = $web.ServerRelativeUrl

# Get library root folder server-relative URL
$Lists   = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $ExcludedLists -notcontains $_.Title }
foreach ($List in $Lists) {
$rootFolderUrl = $List.RootFolder.ServerRelativeUrl

Write-Host "Library root: $($List.RootFolder.ServerRelativeUrl)"

# Get ONLY folders in the library (FSObjType = 1)
# Using list items is typically faster/reliable for large libraries
$caml = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='FSObjType' />
        <Value Type='Integer'>1</Value>
      </Eq>
    </Where>
  </Query>
  <ViewFields>
    <FieldRef Name='FileRef' />
    <FieldRef Name='FileLeafRef' />
    <FieldRef Name='FileDirRef' />
  </ViewFields>
  <RowLimit>5000</RowLimit>
</View>
"@

$folderItems = Get-PnPListItem -List $List.Title -Query $caml -PageSize 5000

# Build output
$results = foreach ($item in $folderItems) {
    $path = [string]$item["FileRef"]       # server-relative full folder path
    $name = [string]$item["FileLeafRef"]   # folder name
    $parent = [string]$item["FileDirRef"]  # parent folder path

    # Remove site server-relative URL from paths
    $pathRelative = $path
    $parentRelative = $parent
    if ($pathRelative.StartsWith($siteServerRelativeUrl)) {
        $pathRelative = $pathRelative.Substring($siteServerRelativeUrl.Length).TrimStart("/")
    }
    if ($parentRelative.StartsWith($siteServerRelativeUrl)) {
        $parentRelative = $parentRelative.Substring($siteServerRelativeUrl.Length).TrimStart("/")
    }

    # Compute "level" relative to library root
    $relative = $path.Replace($rootFolderUrl, "").Trim("/")
    $level = if ([string]::IsNullOrWhiteSpace($relative)) { 0 } else { ($relative -split "/").Count }

    $obj = [pscustomobject]@{
      SiteUrl      = $Site.Url
        Library      = $List.Title
        FolderName   = $name
        FolderPath   = $pathRelative
        ParentPath   = $parentRelative
        Level        = $level
    }

    if ($IncludeItemCount) {
        try {
            $folder = Get-PnPFolder -Url $path -Includes ItemCount
            $obj | Add-Member -NotePropertyName ItemCount -NotePropertyValue $folder.ItemCount
        } catch {
            $obj | Add-Member -NotePropertyName ItemCount -NotePropertyValue $null
        }
    }

    $obj
}
}

    # Sort by path for a clean hierarchy-like listing
    $resultsSorted = $results | Sort-Object FolderPath

    # Export
    $resultsSorted | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputCsv

    Write-Host "Done. Exported folder structure to: $OutputCsv"
}