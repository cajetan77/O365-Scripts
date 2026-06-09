#Requires -Version 7.0

<#
.SYNOPSIS
    Triggers the App Service webhook to start SharePoint provisioning.

.DESCRIPTION
    Sends a POST request to the App Service /caj/webhook endpoint. The App Service
    validates X-Cloud-Governance-Token and forwards the JSON body to the Function App.

.PARAMETER JsonPayload
    JSON string sent as the request body. Use this to pass the full webhook payload.

.PARAMETER JsonFile
    Path to a JSON file sent as the request body.

.PARAMETER ObjectUrl
    SharePoint site URL to provision. Used when -JsonPayload / -JsonFile are not supplied.

.PARAMETER Action
    Provisioning action name passed through to the Function App.

.PARAMETER ProjectName
    Optional project name included in the request body.

.PARAMETER WebhookUri
    Full webhook URL. Defaults to the app-intra-poc-linux1 endpoint.

.PARAMETER CloudGovernanceToken
    Value for the X-Cloud-Governance-Token header. If omitted, the script uses
    $env:CLOUD_GOVERNANCE_TOKEN or reads CLOUD_GOVERNANCE_TOKEN from the App Service.

.PARAMETER ResourceGroup
    Resource group used when fetching the token from Azure app settings.

.PARAMETER AppServiceName
    App Service name used when fetching the token from Azure app settings.

.PARAMETER Subscription
    Azure subscription name or ID when fetching the token from Azure app settings.

.EXAMPLE
    .\TriggerAppService.ps1 -JsonPayload '{"objectUrl":"https://contoso.sharepoint.com/sites/MySite","action":"ProvisionDocumentLibraries","projectName":"Test"}'

.EXAMPLE
    .\TriggerAppService.ps1 -JsonFile .\payload.json -CloudGovernanceToken "<token>"

.EXAMPLE
    .\TriggerAppService.ps1 `
        -ObjectUrl "https://contoso.sharepoint.com/sites/MySite" `
        -CloudGovernanceToken "<token>"
#>

[CmdletBinding(DefaultParameterSetName = 'Fields')]
param(
    [Parameter(Mandatory, ParameterSetName = 'JsonPayload')]
    [string]$JsonPayload,

    [Parameter(Mandatory, ParameterSetName = 'JsonFile')]
    [string]$JsonFile,

    [Parameter(Mandatory, ParameterSetName = 'Fields')]
    [string]$ObjectUrl,

    [Parameter(ParameterSetName = 'Fields')]
    [string]$Action = "ProvisionDocumentLibraries",

    [Parameter(ParameterSetName = 'Fields')]
    [string]$ProjectName = "Test",

    [string]$WebhookUri = "https://app-intra-poc-linux1.azurewebsites.net/caj/webhook",

    [string]$CloudGovernanceToken,

    [string]$ResourceGroup = "SPO-Automation",

    [string]$AppServiceName = "app-intra-poc-linux1",

    [string]$Subscription = "Azure subscription 1"
)

function Get-CloudGovernanceToken {
    if (-not [string]::IsNullOrWhiteSpace($CloudGovernanceToken)) {
        return $CloudGovernanceToken
    }

    if (-not [string]::IsNullOrWhiteSpace($env:CLOUD_GOVERNANCE_TOKEN)) {
        return $env:CLOUD_GOVERNANCE_TOKEN
    }

    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        throw "CloudGovernanceToken was not provided and Azure CLI is not available to read app settings."
    }

    Write-Verbose "Reading CLOUD_GOVERNANCE_TOKEN from App Service app settings..."
    az account set --subscription $Subscription | Out-Null

    $token = az webapp config appsettings list `
        --resource-group $ResourceGroup `
        --name $AppServiceName `
        --query "[?name=='CLOUD_GOVERNANCE_TOKEN'].value | [0]" `
        --output tsv

    if ([string]::IsNullOrWhiteSpace($token)) {
        throw "CLOUD_GOVERNANCE_TOKEN was not found. Pass -CloudGovernanceToken or set `$env:CLOUD_GOVERNANCE_TOKEN."
    }

    return $token
}

function ConvertTo-RequestBody {
    param([string]$RawJson)

    if ([string]::IsNullOrWhiteSpace($RawJson)) {
        throw "JSON payload is empty."
    }

    try {
        $parsed = $RawJson | ConvertFrom-Json
    }
    catch {
        throw "Invalid JSON payload: $($_.Exception.Message)"
    }

    return ($parsed | ConvertTo-Json -Compress -Depth 20)
}

$token = Get-CloudGovernanceToken

$body = switch ($PSCmdlet.ParameterSetName) {
    'JsonPayload' { ConvertTo-RequestBody -RawJson $JsonPayload }
    'JsonFile' {
        if (-not (Test-Path -LiteralPath $JsonFile)) {
            throw "JSON file not found: $JsonFile"
        }

        ConvertTo-RequestBody -RawJson (Get-Content -LiteralPath $JsonFile -Raw)
    }
    'Fields' {
        @{
            objectUrl   = $ObjectUrl
            action      = $Action
            projectName = $ProjectName
        } | ConvertTo-Json -Compress
    }
}

Write-Host "Triggering App Service webhook..."
Write-Host "  URI:     $WebhookUri"
Write-Host "  Payload: $body"

try {
    $response = Invoke-WebRequest `
        -Uri $WebhookUri `
        -Method Post `
        -Headers @{
        "X-Cloud-Governance-Token" = $token
    } `
        -Body $body `
        -ContentType "application/json" `
        -UseBasicParsing -Verbose

    Write-Host "Response status: $($response.StatusCode)" -ForegroundColor Green

    if (-not [string]::IsNullOrWhiteSpace($response.Content)) {
        $parsed = $response.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($null -ne $parsed) {
            $parsed | ConvertTo-Json -Depth 10
        }
        else {
            $response.Content
        }
    }
}
catch {
    $statusCode = $null
    $errorBody = $_.ErrorDetails.Message

    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        if ([string]::IsNullOrWhiteSpace($errorBody)) {
            $errorBody = $_.Exception.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        }
    }

    if ($statusCode) {
        Write-Error "Webhook call failed with status ${statusCode}: $errorBody"
    }
    else {
        $message = if (-not [string]::IsNullOrWhiteSpace($errorBody)) { $errorBody } else { $_.Exception.Message }
        if ($_.Exception.InnerException) {
            $message = "$message ($($_.Exception.InnerException.Message))"
        }

        Write-Error $message
    }

    exit 1
}