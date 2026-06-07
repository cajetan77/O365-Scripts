$RESOURCE_GROUP = "SPO-Automation"
$LOCATION = "australiasoutheast" 
$STORAGE_ACCOUNT_NAME = "spostoragecaj134" # Your target storage account name
$FUNCTION_APP_NAME = "func-secure-processor02" # Must be globally unique across Azure
$RUNTIME = "powershell" 
$RUNTIME_VERSION = "7.4"
$subscription = "Azure subscription 1"

Write-Host "Deploying Function App via Azure CLI..."
az login 
az account set --subscription $subscription

# 1. Create a serverless Consumption Function App and turn on System-Assigned Identity
Write-Host "Creating Function App and activating Managed Identity..."
az functionapp create --name "$FUNCTION_APP_NAME" `
  --resource-group "$RESOURCE_GROUP" `
  --flexconsumption-location "$LOCATION" `
  --storage-account "$STORAGE_ACCOUNT_NAME" `
  --runtime "$RUNTIME" `
  --runtime-version "$RUNTIME_VERSION" `
  --functions-version 4 `
  --assign-identity '[system]'

# 2. Extract the auto-generated Managed Identity Principal ID
Write-Host "Retrieving Managed Identity details..."
$PRINCIPAL_ID = $(az functionapp `identity show `
    --name "$FUNCTION_APP_NAME" `
    --resource-group "$RESOURCE_GROUP" `
    --query principalId `
    --output tsv)

Write-Host "Managed Identity Principal ID: $PRINCIPAL_ID"

# 3. Retrieve the resource ID of your target Storage Account
Write-Host "Fetching Storage Account ID..."
$STORAGE_ID = $(az storage account show `
    --name "$STORAGE_ACCOUNT_NAME" `
    --resource-group "$RESOURCE_GROUP" `
    --query id `
    --output tsv)

# 4. Grant the Function App permission to read/write data inside the Storage Account
Write-Host "Assigning 'Storage Blob `Data Contributor' role to the Function App..."
az role assignment create `
  --assignee "$PRINCIPAL_ID" `
  --role "Storage Blob Data Contributor" `
  --scope "$STORAGE_ID"

Write-Host "==================================``======================="
Write-Host "SUCCESS: Function App '$FUNCTION_APP_NAME' is deployed!"
Write-Host "Authentication: Managed Identity is active and role assigned."
Write-Host "========================================================="
