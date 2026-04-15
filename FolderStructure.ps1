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

# Import PnP PowerShell module
Import-Module PnP.PowerShell -Force

try {
    $configPath = ".\config.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $TenantId = $config.TenantId
    $ClientId = $config.SharePointReportingAppId
    $Thumbprint = $config.ThumbPrint
    Write-Host "Configuration loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error loading configuration from config.json: $_" -ForegroundColor Red
    exit 1
}

try {
    $sites = Import-Csv -Path ".\sites.csv"
    Write-Host "Found $($sites.Count) site(s) to process" -ForegroundColor Cyan
}
catch {
    Write-Host "Error loading sites from sites.csv: $_" -ForegroundColor Red
    exit 1
}

$allResults = @()

foreach ($site in $sites) {
    Write-Host "`nProcessing site: $($site.Url)" -ForegroundColor Green
    
    # Connect to site
    try {
        Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
        Write-Host "Connected successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error connecting to SharePoint: $_" -ForegroundColor Red
        continue
    }

    # Get site's server-relative URL to remove from paths
    try {
        $web = Get-PnPWeb
        $siteServerRelativeUrl = $web.ServerRelativeUrl
        Write-Host "Retrieved site information successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error getting site information: $_" -ForegroundColor Red
        continue
    }

    # Get library root folder server-relative URL
    try {
        $Lists = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $ExcludedLists -notcontains $_.Title }
        Write-Host "Found $($Lists.Count) document libraries" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error getting document libraries: $_" -ForegroundColor Red
        continue
    }
    
    foreach ($List in $Lists) {
        Write-Host "  Processing library: $($List.Title)" -ForegroundColor Yellow
        
        try {
            $rootFolderUrl = $List.RootFolder.ServerRelativeUrl
            Write-Host "Library root: $($List.RootFolder.ServerRelativeUrl)"

            # Get all list items and then filter for folders only
            Write-Host "Getting all list items..."
            $allItems = Get-PnPListItem -List $List.Title -PageSize 5000
            
            # Filter to get only folders (FSObjType = 1)
            $folderItems = $allItems | Where-Object { $_["FSObjType"] -eq 1 }
            
            Write-Host "Total items: $($allItems.Count), Folders: $($folderItems.Count)"
        }
        catch {
            Write-Host "  Error processing library '$($List.Title)': $_" -ForegroundColor Red
            Write-Host "  Skipping this library and continuing..." -ForegroundColor Yellow
            continue
        }

        # Build output for this library
        $results = @()
        foreach ($item in $folderItems) {
            try {
                $path = [string]$item["FileRef"]       # server-relative full folder path
                $name = [string]$item["FileLeafRef"]   # folder name
                $parent = [string]$item["FileDirRef"]  # parent folder path

                # Remove library root path to get folder path relative to library
                $folderPathRelativeToLibrary = $path.Replace($rootFolderUrl, "").TrimStart("/")
                $parentPathRelativeToLibrary = $parent.Replace($rootFolderUrl, "").TrimStart("/")
        
                # If path is empty, it means this is the library root
                if ([string]::IsNullOrWhiteSpace($folderPathRelativeToLibrary)) {
                    $folderPathRelativeToLibrary = "/"
                }
                if ([string]::IsNullOrWhiteSpace($parentPathRelativeToLibrary)) {
                    $parentPathRelativeToLibrary = "/"
                }

                # Compute "level" relative to library root
                $level = if ($folderPathRelativeToLibrary -eq "/") { 0 } else { ($folderPathRelativeToLibrary -split "/").Count }

                $obj = [pscustomobject]@{
                    SiteUrl    = $site.Url
                    Library    = $List.Title
                    FolderName = $name
                    FolderPath = $folderPathRelativeToLibrary
                    #ParentPath = $parentPathRelativeToLibrary
                    #Level      = $level
                }

                if ($IncludeItemCount) {
                    try {
                        $folder = Get-PnPFolder -Url $path -Includes ItemCount
                        $obj | Add-Member -NotePropertyName ItemCount -NotePropertyValue $folder.ItemCount
                    }
                    catch {
                        Write-Host "    Warning: Could not get item count for folder '$name': $_" -ForegroundColor Yellow
                        $obj | Add-Member -NotePropertyName ItemCount -NotePropertyValue $null
                    }
                }

                $results += $obj
            }
            catch {
                Write-Host "    Error processing folder item: $_" -ForegroundColor Red
                continue
            }
        }
        
        # Add results from this library to the overall collection
        $allResults += $results
        Write-Host "    Added $($results.Count) folder(s) to results" -ForegroundColor Gray
    }
    
    try {
        Write-Host "Disconnecting from site" -ForegroundColor Green
        Disconnect-PnPOnline
    }
    catch {
        Write-Host "Warning: Error disconnecting from site: $_" -ForegroundColor Yellow
    }
}

# Sort all results by site and path for a clean hierarchy-like listing
try {
    Write-Host "`nTotal folders found: $($allResults.Count)" -ForegroundColor Cyan
    $resultsSorted = $allResults | Sort-Object SiteUrl, Library, FolderPath

    # Export all results
    $resultsSorted | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputCsv
    Write-Host "Done. Exported folder structure to: $OutputCsv" -ForegroundColor Green
}
catch {
    Write-Host "Error exporting results to CSV: $_" -ForegroundColor Red
    Write-Host "Results count: $($allResults.Count)" -ForegroundColor Yellow
}