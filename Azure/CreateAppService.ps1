$subscription = "Azure subscription 1"
$resourceGroupName = "SPO-Automation"
$appServicePlanName = "plan-spo-automation-consumption-linux"
$appServiceName = "app-intra-poc-linux"
$runbookName = "RunbookAppService"
$automationAccount = "aa-automation"


az login
az account set --subscription $subscription

az appservice plan create --name $appServicePlanName --resource-group $resourceGroupName --sku B1  --is-linux


az  webapp create --name $appServiceName `
    --resource-group $resourceGroupName `
    --plan $appServicePlanName `
    --runtime "DOTNETCORE:10.0"

dotnet new webapi -n $appServiceName 
cd $appServiceName


dotnet add package Azure.Identity --version 1.21.0
dotnet add package Azure.ResourceManager.Automation --version 1.1.2


dotnet build
# Compile

dotnet publish -c Release -o ../AppOutputClean --clean


Remove-Item -Recurse -Force ./publish -ErrorAction SilentlyContinue
dotnet publish -c Release -o ./publish



# Zip
Compress-Archive -Path ./publish/* -DestinationPath ./deploy.zip -Force

# Assign identity to the app service
az webapp identity assign --name $appServiceName --resource-group $resourceGroupName


# Deploy
az webapp deployment source config-zip `
    --name $appServiceName `
    --resource-group $resourceGroupName `
    --src "./deploy.zip" 

az webapp deploy --resource-group $resourceGroupName --name $appServiceName --src-path "./deploy.zip"



az webapp restart --name app-intra-poc --resource-group SPO-Automation


