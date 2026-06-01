# 1. Configuration - Set your live endpoint URL
$url = ""
$secretToken = ""

# 2. Define the exact JSON payload matching your C# structure
$payload = @{
    projectName = "SharePoint Migration Project"
    targetSite  = "https://sharepoint.com"
    action      = "ProvisionDocumentLibraries"
} | ConvertTo-Json

$headers = @{
    "Authorization" = "Bearer $secretToken"
}
# 3. Transmit the JSON payload to your Azure App Service API
Write-Host "Sending JSON test payload to Azure..." -ForegroundColor Cyan
$response = Invoke-RestMethod -Uri $url -Method Post -Body $payload -ContentType "application/json" -Headers $headers

# 4. Print out the response returned from your .NET app
Write-Host "Response received from Azure:" -ForegroundColor Green
$response | Format-List
