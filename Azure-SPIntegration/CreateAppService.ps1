$subscription = "Azure subscription 1"
$resourceGroupName = "SPO-Automation"
$appServicePlanName = "plan-spo-automation-consumption-linux"
$appServiceName = "app-intra-poc-linux1"



az login
az account set --subscription $subscription

az appservice plan show `
    --resource-group $resourceGroupName `
    --name $appServicePlanName `


az appservice plan create --name $appServicePlanName --resource-group $resourceGroupName --sku B1  --is-linux


az  webapp create --name $appServiceName `
    --resource-group $resourceGroupName `
    --plan $appServicePlanName `
    --runtime "DOTNETCORE:8.0"

dotnet new webapi -n $appServiceName 
cd $appServiceName



dotnet build
# Compile




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

