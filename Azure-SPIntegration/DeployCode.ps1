$RESOURCE_GROUP = "SPO-Automation"
$LOCATION = "australiasoutheast" 
$STORAGE_ACCOUNT_NAME = "spostoragecaj134" # Your target storage account name
$FUNCTION_APP_NAME = "func-secure-processor02" # Must be globally unique across Azure
$RUNTIME = "powershell" 
$RUNTIME_VERSION = "7.4"
$subscription = "Azure subscription 1"

# App Service settings (configure in Azure; secrets must differ):
#   CLOUD_GOVERNANCE_TOKEN  -> X-Cloud-Governance-Token from Cloud Governance
#   FUNCTION_HEADER_VALUE   -> X-INTERNAL-KEY (same value on App Service and Function App)
#   FUNCTION_URL            -> https://.../api/ProvisionSite (without ?code=)
#   FUNCTION_KEY            -> Function host key, sent as x-functions-key header


az login 
az account set --subscription $subscription

#Save-Module -Name Az.Accounts -RequiredVersion "3.0.1" -Path .\ExternalModules -Repository PSGallery -Force
#Save-Module -Name PnP.PowerShell -RequiredVersion "3.2.0" -Path .\ExternalModules -Repository PSGallery -Force

Remove-Item .\function.zip -Force -ErrorAction SilentlyContinue

Compress-Archive `
    -Path .\host.json, .\requirements.psd1, .\Modules, .\ProvisionSite, .\ExternalModules `
    -DestinationPath .\function.zip `
    -Force    

az functionapp deployment source config-zip `
    --resource-group $RESOURCE_GROUP `
    --name $FUNCTION_APP_NAME `
    --src ".\function.zip"

$cloudGovernanceToken = "Psalm87&6"   

$uri = "https://app-intra-poc-linux1.azurewebsites.net/caj/webhook"

$body = @{
    objectUrl   = "https://caje77sharepoint.sharepoint.com/sites/CajIntra"
    action      = "ProvisionDocumentLibraries"
    projectName = "Test"
} | ConvertTo-Json

Invoke-RestMethod `
    -Uri $uri `
    -Method Post `
    -Headers @{
    "X-Cloud-Governance-Token" = $cloudGovernanceToken
} `
    -Body $body `
    -ContentType "application/json"  





# Run this from your local workspace root directory

