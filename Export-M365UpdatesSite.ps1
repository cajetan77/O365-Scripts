# Export data from M365Updates SharePoint site to CSV in Documents folder
Import-Module PnP.PowerShell -Force

# Load configuration
$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$Thumbprint = $config.ThumbPrint

# Site URL
$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/M365Updates"

# Output path - Documents folder
$documentsPath = [Environment]::GetFolderPath("MyDocuments")
$outputPath = Join-Path $documentsPath "M365Updates-Export.csv"

Write-Host "Connecting to M365Updates site..." -ForegroundColor Yellow
Write-Host "Site URL: $SiteUrl" -ForegroundColor Gray

try {
    # Connect to SharePoint site
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
    Write-Host "✓ Connected successfully!" -ForegroundColor Green
    
    # Get site information
    $web = Get-PnPWeb
    Write-Host "`nSite Information:" -ForegroundColor Cyan
    Write-Host "  Title: $($web.Title)" -ForegroundColor White
    Write-Host "  Description: $($web.Description)" -ForegroundColor White
    Write-Host "  URL: $($web.Url)" -ForegroundColor White
    
    # Get all lists in the site
    Write-Host "`nGetting all lists..." -ForegroundColor Yellow
    $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.Title -notlike "Style Library" -and $_.Title -notlike "Form Templates" }
    
    Write-Host "Found $($lists.Count) lists:" -ForegroundColor Green
    foreach ($list in $lists) {
        Write-Host "  - $($list.Title) ($($list.BaseType)) - $($list.ItemCount) items" -ForegroundColor White
    }
    
    # Export data from all lists
    $allData = @()
    
    foreach ($list in $lists) {
        Write-Host "`nProcessing list: $($list.Title)..." -ForegroundColor Cyan
        
        try {
            # Get list items
            $items = Get-PnPListItem -List $list.Title -PageSize 1000
            Write-Host "  Retrieved $($items.Count) items" -ForegroundColor Green
            
            # Process each item
            foreach ($item in $items) {
                $itemData = [PSCustomObject]@{
                    ListName = $list.Title
                    ListType = $list.BaseType
                    ItemID = $item.Id
                    Title = $item.FieldValues.Title
                    Created = $item.FieldValues.Created
                    Modified = $item.FieldValues.Modified
                    CreatedBy = $item.FieldValues.Author.LookupValue
                    ModifiedBy = $item.FieldValues.Editor.LookupValue
                    FileRef = $item.FieldValues.FileRef
                    ContentType = $item.FieldValues.ContentType
                }
                
                # Add custom fields if they exist
                if ($item.FieldValues.ContainsKey("Description")) {
                    $itemData | Add-Member -MemberType NoteProperty -Name "Description" -Value $item.FieldValues.Description
                }
                if ($item.FieldValues.ContainsKey("Category")) {
                    $itemData | Add-Member -MemberType NoteProperty -Name "Category" -Value $item.FieldValues.Category
                }
                if ($item.FieldValues.ContainsKey("Status")) {
                    $itemData | Add-Member -MemberType NoteProperty -Name "Status" -Value $item.FieldValues.Status
                }
                
                $allData += $itemData
            }
        }
        catch {
            Write-Host "  Error processing list $($list.Title): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Export to CSV
    if ($allData.Count -gt 0) {
        Write-Host "`nExporting $($allData.Count) items to CSV..." -ForegroundColor Yellow
        $allData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Export completed!" -ForegroundColor Green
        Write-Host "File saved to: $outputPath" -ForegroundColor Cyan
        
        # Show sample data
        Write-Host "`nSample exported data:" -ForegroundColor Cyan
        $allData | Select-Object -First 3 | Format-Table -AutoSize
    } else {
        Write-Host "No data found to export" -ForegroundColor Yellow
    }
    
}
catch {
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Disconnect
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "`nDisconnected from SharePoint" -ForegroundColor Gray
    }
    catch {
        # Ignore disconnect errors
    }
}

Write-Host "`nScript completed!" -ForegroundColor Green