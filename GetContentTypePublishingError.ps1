# Load configuration from JSON file
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

# Initialize array to collect errors
$errors = New-Object System.Collections.Generic.List[object]

function Connect_Site
{
    param(
        [string]$SiteUrl
    )
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
    }
    catch {
        Write-Host "Error connecting to site $($site.Url): $_" -ForegroundColor Red
        exit
    }
}


$adminUrl = "https://$TenantName-admin.sharepoint.com"
Connect_Site -SiteUrl $adminUrl

$sites = Get-PnPTenantSite | Where-Object { $_.Url -like "*/sites*" }

foreach ($site in $sites) {
    Connect_Site -SiteUrl $site.Url
 
 $list= Get-PnPList | Where-Object {  $_.Hidden -eq $true -and $_.Title -eq "Content type publishing error log" }
if($list.Count -gt 0) {
    Write-Host "List: $($list.Title) found in $($site.Url)"
    $items = Get-PnPListItem -List $list.Title
    foreach ($item in $items) {
       Write-Host "Item: $($item.FieldValues.Title)" 
       Write-Host "Content Type: $($item.FieldValues.PublishedObjectName)"
       Write-Host "Failure Message: $($item.FieldValues.Failure_x0020_Message )"
       Write-Host "Failure Time: $($item.FieldValues.Failure_x0020_Time)"
      
       
       # Add error to collection
       $errorObject = [PSCustomObject]@{
           SiteUrl = $site.Url
           Title = $item.FieldValues.Title
           PublishedObjectName = $item.FieldValues.PublishedObjectName
           FailureMessage = $item.FieldValues.Failure_x0020_Message
           FailureTime = $item.FieldValues.Failure_x0020_Time
       }
       $errors.Add($errorObject)
    }
}
else {
   # Write-Host "List: $($list.Title) not found in $($site.Url)"
}
}

# Export errors to CSV
if ($errors.Count -gt 0) {
    $csvPath = ".\ContentTypePublishingErrors.csv"
    $errors | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported $($errors.Count) error(s) to $csvPath" -ForegroundColor Green
}
else {
    Write-Host "`nNo errors found to export." -ForegroundColor Yellow
}
