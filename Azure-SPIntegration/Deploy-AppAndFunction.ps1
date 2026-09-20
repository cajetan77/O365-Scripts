#Requires -Version 7.0
<#
.SYNOPSIS
    Creates the App Service + Function App and deploys the project assets.

.DESCRIPTION
    Idempotent Azure CLI script that:
      1. Ensures resource group, storage account, App Service plan, Web App, Function App
      2. Enables Function App system-assigned managed identity + Storage Blob Data Contributor
      3. Publishes and deploys AppService/app-intra-poc-linux1
      4. Zips and deploys FunctionApp (ProvisionSite, modules, Assets) — excludes cert.pfx
      5. Optionally uploads FunctionApp/Assets to a blob container
      6. Wires CLOUD_GOVERNANCE_TOKEN / FUNCTION_* app settings between the two apps

.EXAMPLE
    .\Deploy-AppAndFunction.ps1 `
        -CloudGovernanceToken '<caller-secret>' `
        -FunctionHeaderValue '<internal-secret>'

.EXAMPLE
    .\Deploy-AppAndFunction.ps1 -SkipLogin -SkipCreate -UploadAssets
#>
[CmdletBinding()]
param(
    [string]$Subscription = 'Azure subscription 1',
    [string]$ResourceGroup = 'SPO-Automation',
    [string]$Location = 'australiasoutheast',

    [string]$StorageAccountName = 'spostoragecaj134',
    [string]$BlobContainerName = 'pnp-assets',

    [string]$AppServicePlanName = 'plan-spo-automation-consumption-linux',
    [string]$AppServiceName = 'app-intra-poc-linux1',
    [string]$AppServiceSku = 'B1',

    [string]$FunctionAppName = 'func-secure-processor02',
    [string]$FunctionRuntime = 'powershell',
    # Azure warns 7.4 EOL; Flex Consumption matches CreateInfra.ps1
    [string]$FunctionRuntimeVersion = '7.4',
    [ValidateSet('Flex', 'WindowsConsumption')]
    [string]$FunctionHosting = 'Flex',

    [string]$CloudGovernanceToken,
    [string]$FunctionHeaderValue,

    [string]$AppServiceProjectPath = (Join-Path $PSScriptRoot 'AppService\app-intra-poc-linux1'),
    [string]$FunctionAppPath = (Join-Path $PSScriptRoot 'FunctionApp'),

    [switch]$SkipLogin,
    [switch]$SkipCreate,
    [switch]$SkipDeployAppService,
    [switch]$SkipDeployFunction,
    [switch]$UploadAssets,
    [switch]$SkipAppSettings
)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Assert-Command([string]$Name) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "Required command not found: $Name"
    }
}

function Test-AzResourceExists {
    param([scriptblock]$ShowCommand)
    $null = & $ShowCommand 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Invoke-Az {
    param([Parameter(ValueFromRemainingArguments = $true)][object[]]$AzArgs)
    Write-Host ("az " + ($AzArgs -join ' ')) -ForegroundColor DarkGray
    & az @AzArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI failed: az $($AzArgs -join ' ')"
    }
}

Assert-Command az
Assert-Command dotnet

if (-not $SkipLogin) {
    Write-Step 'Azure login / subscription'
    Invoke-Az login
    Invoke-Az account set --subscription $Subscription
}
else {
    Write-Step "Using current Azure CLI session (subscription: $Subscription)"
    Invoke-Az account set --subscription $Subscription
}

# -----------------------------------------------------------------------------
# Create infrastructure
# -----------------------------------------------------------------------------
if (-not $SkipCreate) {
    Write-Step "Ensure resource group '$ResourceGroup'"
    if (-not (Test-AzResourceExists { az group show --name $ResourceGroup })) {
        Invoke-Az group create --name $ResourceGroup --location $Location
    }
    else {
        Write-Host "Resource group already exists."
    }

    Write-Step "Ensure storage account '$StorageAccountName'"
    if (-not (Test-AzResourceExists {
                az storage account show --name $StorageAccountName --resource-group $ResourceGroup
            })) {
        Invoke-Az storage account create `
            --name $StorageAccountName `
            --resource-group $ResourceGroup `
            --location $Location `
            --sku Standard_LRS `
            --kind StorageV2
    }
    else {
        Write-Host "Storage account already exists."
    }

    Write-Step "Ensure App Service plan '$AppServicePlanName'"
    if (-not (Test-AzResourceExists {
                az appservice plan show --name $AppServicePlanName --resource-group $ResourceGroup
            })) {
        Invoke-Az appservice plan create `
            --name $AppServicePlanName `
            --resource-group $ResourceGroup `
            --sku $AppServiceSku `
            --is-linux
    }
    else {
        Write-Host "App Service plan already exists."
    }

    Write-Step "Ensure Web App '$AppServiceName'"
    if (-not (Test-AzResourceExists {
                az webapp show --name $AppServiceName --resource-group $ResourceGroup
            })) {
        Invoke-Az webapp create `
            --name $AppServiceName `
            --resource-group $ResourceGroup `
            --plan $AppServicePlanName `
            --runtime 'DOTNETCORE:8.0'
    }
    else {
        Write-Host "Web App already exists."
    }

    Invoke-Az webapp identity assign --name $AppServiceName --resource-group $ResourceGroup | Out-Null

    Write-Step "Ensure Function App '$FunctionAppName' ($FunctionHosting)"
    if (-not (Test-AzResourceExists {
                az functionapp show --name $FunctionAppName --resource-group $ResourceGroup
            })) {
        # SPO-Automation often already has Windows consumption history, so Linux
        # Consumption fails with: "Linux dynamic workers are not available in resource group".
        # Flex Consumption (CreateInfra.ps1) avoids that. WindowsConsumption is the fallback.
        if ($FunctionHosting -eq 'Flex') {
            Invoke-Az functionapp create `
                --name $FunctionAppName `
                --resource-group $ResourceGroup `
                --flexconsumption-location $Location `
                --storage-account $StorageAccountName `
                --runtime $FunctionRuntime `
                --runtime-version $FunctionRuntimeVersion `
                --functions-version 4 `
                --assign-identity '[system]'
        }
        else {
            Invoke-Az functionapp create `
                --name $FunctionAppName `
                --resource-group $ResourceGroup `
                --storage-account $StorageAccountName `
                --consumption-plan-location $Location `
                --runtime $FunctionRuntime `
                --runtime-version $FunctionRuntimeVersion `
                --functions-version 4 `
                --os-type Windows `
                --assign-identity '[system]'
        }
    }
    else {
        Write-Host "Function App already exists."
        Invoke-Az functionapp identity assign --name $FunctionAppName --resource-group $ResourceGroup | Out-Null
    }

    Write-Step 'Grant Function App Storage Blob Data Contributor'
    $principalId = az functionapp identity show `
        --name $FunctionAppName `
        --resource-group $ResourceGroup `
        --query principalId `
        --output tsv
    if ([string]::IsNullOrWhiteSpace($principalId)) {
        throw "Could not read Function App managed identity principalId."
    }

    $storageId = az storage account show `
        --name $StorageAccountName `
        --resource-group $ResourceGroup `
        --query id `
        --output tsv

    $existingRole = az role assignment list `
        --assignee $principalId `
        --role 'Storage Blob Data Contributor' `
        --scope $storageId `
        --query '[0].id' `
        --output tsv 2>$null

    if ([string]::IsNullOrWhiteSpace($existingRole)) {
        Invoke-Az role assignment create `
            --assignee $principalId `
            --role 'Storage Blob Data Contributor' `
            --scope $storageId
    }
    else {
        Write-Host "Role assignment already present."
    }
}

# -----------------------------------------------------------------------------
# Deploy App Service
# -----------------------------------------------------------------------------
if (-not $SkipDeployAppService) {
    Write-Step "Publish + deploy App Service from '$AppServiceProjectPath'"
    $csproj = Get-ChildItem -LiteralPath $AppServiceProjectPath -Filter '*.csproj' -File -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $csproj) {
        throw "App Service project not found under: $AppServiceProjectPath"
    }

    Push-Location $AppServiceProjectPath
    try {
        $publishDir = Join-Path $AppServiceProjectPath 'publish'
        $zipPath = Join-Path $AppServiceProjectPath 'deploy.zip'

        Remove-Item -Recurse -Force $publishDir -ErrorAction SilentlyContinue
        Remove-Item -Force $zipPath -ErrorAction SilentlyContinue

        dotnet publish -c Release -o $publishDir
        if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed.' }

        Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $zipPath -Force

        Invoke-Az webapp deploy `
            --resource-group $ResourceGroup `
            --name $AppServiceName `
            --src-path $zipPath `
            --type zip `
            --clean true

        Invoke-Az webapp restart --name $AppServiceName --resource-group $ResourceGroup
    }
    finally {
        Pop-Location
    }
}

# -----------------------------------------------------------------------------
# Deploy Function App (+ Assets in package)
# -----------------------------------------------------------------------------
if (-not $SkipDeployFunction) {
    Write-Step "Zip + deploy Function App from '$FunctionAppPath'"
    if (-not (Test-Path -LiteralPath (Join-Path $FunctionAppPath 'host.json'))) {
        throw "Function App host.json not found under: $FunctionAppPath"
    }

    Push-Location $FunctionAppPath
    try {
        $zipPath = Join-Path $FunctionAppPath 'function.zip'
        Remove-Item -Force $zipPath -ErrorAction SilentlyContinue

        $include = @(
            'host.json'
            'requirements.psd1'
            'FunctionModules'
            'ProvisionSite'
            'ExternalModules'
            'Assets'
        ) | Where-Object { Test-Path -LiteralPath (Join-Path $FunctionAppPath $_) }

        if ($include.Count -eq 0) {
            throw 'No Function App assets found to zip.'
        }

        Compress-Archive -Path $include -DestinationPath $zipPath -Force

        Invoke-Az functionapp deployment source config-zip `
            --resource-group $ResourceGroup `
            --name $FunctionAppName `
            --src $zipPath

        Invoke-Az functionapp restart --name $FunctionAppName --resource-group $ResourceGroup
    }
    finally {
        Pop-Location
    }
}

# -----------------------------------------------------------------------------
# Optional: upload Assets to blob (folder already used by Function for PnP URLs)
# -----------------------------------------------------------------------------
if ($UploadAssets) {
    Write-Step "Upload FunctionApp/Assets to blob container '$BlobContainerName'"
    $assetsPath = Join-Path $FunctionAppPath 'Assets'
    if (-not (Test-Path -LiteralPath $assetsPath)) {
        throw "Assets folder not found: $assetsPath"
    }

    $exists = az storage container exists `
        --account-name $StorageAccountName `
        --name $BlobContainerName `
        --auth-mode login `
        --query exists `
        --output tsv

    if ($exists -ne 'true') {
        Invoke-Az storage container create `
            --account-name $StorageAccountName `
            --name $BlobContainerName `
            --auth-mode login
    }

    Get-ChildItem -LiteralPath $assetsPath -File | ForEach-Object {
        Invoke-Az storage blob upload `
            --account-name $StorageAccountName `
            --container-name $BlobContainerName `
            --name $_.Name `
            --file $_.FullName `
            --auth-mode login `
            --overwrite true
        Write-Host "Uploaded blob: $($_.Name)"
    }
}

# -----------------------------------------------------------------------------
# Wire App Service ↔ Function App settings
# -----------------------------------------------------------------------------
if (-not $SkipAppSettings) {
    Write-Step 'Configure App Service / Function App settings'

    if ([string]::IsNullOrWhiteSpace($CloudGovernanceToken)) {
        $CloudGovernanceToken = $env:CLOUD_GOVERNANCE_TOKEN
    }
    if ([string]::IsNullOrWhiteSpace($FunctionHeaderValue)) {
        $FunctionHeaderValue = $env:FUNCTION_HEADER_VALUE
    }

    if ([string]::IsNullOrWhiteSpace($CloudGovernanceToken) -or [string]::IsNullOrWhiteSpace($FunctionHeaderValue)) {
        Write-Warning @"
Skipping secret app settings — pass -CloudGovernanceToken and -FunctionHeaderValue
(or set `$env:CLOUD_GOVERNANCE_TOKEN` / `$env:FUNCTION_HEADER_VALUE`).
They must be different values.
"@
    }
    elseif ($CloudGovernanceToken -eq $FunctionHeaderValue) {
        throw 'CLOUD_GOVERNANCE_TOKEN and FUNCTION_HEADER_VALUE must be different.'
    }
    else {
        $functionKey = az functionapp keys list `
            --resource-group $ResourceGroup `
            --name $FunctionAppName `
            --query 'functionKeys.default' `
            --output tsv

        if ([string]::IsNullOrWhiteSpace($functionKey)) {
            $functionKey = az functionapp keys list `
                --resource-group $ResourceGroup `
                --name $FunctionAppName `
                --query 'masterKey' `
                --output tsv
        }

        if ([string]::IsNullOrWhiteSpace($functionKey)) {
            throw 'Could not retrieve a Function App key. Deploy the function first, then re-run with -SkipCreate -SkipDeployAppService -SkipDeployFunction.'
        }

        $functionUrl = "https://$FunctionAppName.azurewebsites.net/api/ProvisionSite"

        # Web App: @json array works. Function App Flex: @json is misparsed into
        # settings literally named "name"/"value" — use KEY=VALUE instead.
        $appSettingsFile = Join-Path $env:TEMP 'appservice-settings-deploy.json'
        @(
            @{ name = 'CLOUD_GOVERNANCE_TOKEN'; value = $CloudGovernanceToken }
            @{ name = 'FUNCTION_URL'; value = $functionUrl }
            @{ name = 'FUNCTION_KEY'; value = $functionKey }
            @{ name = 'FUNCTION_HEADER_VALUE'; value = $FunctionHeaderValue }
        ) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $appSettingsFile -Encoding utf8

        try {
            Invoke-Az webapp config appsettings set `
                --resource-group $ResourceGroup `
                --name $AppServiceName `
                --settings "@$appSettingsFile"

            Invoke-Az functionapp config appsettings set `
                --resource-group $ResourceGroup `
                --name $FunctionAppName `
                --settings "FUNCTION_HEADER_VALUE=$FunctionHeaderValue"
        }
        finally {
            Remove-Item -Force $appSettingsFile -ErrorAction SilentlyContinue
        }

        Write-Host "App settings applied. SPO_* / PNP_* settings still need to be set separately (cert + blob URLs)."
    }
}

$webhookUrl = "https://$AppServiceName.azurewebsites.net/caj/webhook"
$functionUrlOut = "https://$FunctionAppName.azurewebsites.net/api/ProvisionSite"

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Green
Write-Host "Done."
Write-Host "  App Service webhook : $webhookUrl"
Write-Host "  Function endpoint   : $functionUrlOut"
Write-Host "  Storage account     : $StorageAccountName"
if ($UploadAssets) {
    Write-Host "  Assets container    : $BlobContainerName"
}
Write-Host "=========================================================" -ForegroundColor Green
Write-Host @"

Next (SharePoint side — not set by this script):
  - Upload .pfx to Function App Certificates + Entra app registration
  - Set Function App: WEBSITE_LOAD_CERTIFICATES, SPO_CERT_THUMBPRINT, SPO_CLIENT_ID, SPO_TENANT_ID
  - Set Function App: PNP_TEMPLATE_BLOB_URL, PNP_BRANDING_BLOB_URL
  - Test: .\AppService\TriggerAppService.ps1 -ObjectUrl '<site-url>' -CloudGovernanceToken '<token>'
"@
