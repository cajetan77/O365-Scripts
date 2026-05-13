$config = Get-Content -Raw -Path "..\config.json" | ConvertFrom-Json
$clientId = $config.AppId
$thumbprint = $config.ThumbPrint
$tenantId = $config.TenantId
$listName = "Site Pages"
Connect-PnPOnline -Url "https://caje77sharepoint.sharepoint.com/sites/App-1" -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId

$siteUrl = (Get-PnPConnection).Url.TrimEnd('/')
$list = Get-PnPList -Identity $listName
$apiUrls = @(
    "$siteUrl/_api/web/lists(guid'$($list.Id)')/SyncFlowInstances",
    "$siteUrl/_api/web/lists/GetByTitle('$listName')/SyncFlowInstances"
)

$responseObject = $null
$lastError = $null
foreach ($apiUrl in $apiUrls) {
    try {
        Write-Host "Trying: $apiUrl" -ForegroundColor Cyan
        $response = Invoke-PnPSPRestMethod -Method Post -Url $apiUrl -Content "{}" -ContentType "application/json;odata=verbose"
        $responseObject = if ($response -is [string]) { $response | ConvertFrom-Json } else { $response }
        break
    }
    catch {
        $lastError = $_
        Write-Warning "SyncFlowInstances failed for $apiUrl : $($_.Exception.Message)"
    }
}

if (-not $responseObject) {
    throw "SyncFlowInstances failed for all endpoints. This is commonly caused by endpoint limitations with app-only auth on Site Pages. Last error: $($lastError.Exception.Message)"
}

$syncDataRaw = $responseObject.FlowSynchronizationResult.SynchronizationData
if ([string]::IsNullOrWhiteSpace($syncDataRaw)) {
    Write-Warning "No flow synchronization data returned."
    $flowData = @()
}
else {
    $flowData = $syncDataRaw | ConvertFrom-Json
}

$flowData
