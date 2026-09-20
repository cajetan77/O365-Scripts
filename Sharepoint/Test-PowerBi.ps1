$ErrorActionPreference = "Stop"

$ConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"
if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $PSScriptRoot "..\config.json"
}
if (-not (Test-Path $ConfigPath)) {
    throw "config.json not found. Place it next to / above this script."
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$ClientSecret = $config.ClientSecret
if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
    throw "config.json is missing ClientSecret."
}

$SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force

$Credential = New-Object System.Management.Automation.PSCredential(
    $ClientId,
    $SecureSecret
)

Connect-PowerBIServiceAccount `
    -ServicePrincipal `
    -Tenant $TenantId `
    -Credential $Credential

Get-PowerBIWorkspace -Scope Organization


$WorkspaceId = "586bfe85-ac4b-4314-ac61-982723223f07"
$DatasetId = "f41ca084-844c-4483-934b-4d567d058256"

$response = Invoke-PowerBIRestMethod `
    -Url "groups/$WorkspaceId/datasets/$DatasetId/datasources" `
    -Method Get

$response | ConvertFrom-Json | ConvertTo-Json -Depth 10


$GatewayId = "846ea36b-3edc-4c3c-9b40-8b7f4ef21c17"
$DatasourceId = "631586a4-bd7f-4081-b7c8-bef3a711090f"

$ds = Invoke-PowerBIRestMethod `
    -Url "gateways/$GatewayId/datasources/$DatasourceId" `
    -Method Get

$ds | ConvertFrom-Json | ConvertTo-Json -Depth 10


$TokenBody = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://caje77sharepoint.sharepoint.com/.default"
    grant_type    = "client_credentials"
}

$TokenResponse = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $TokenBody

$SharePointAccessToken = $TokenResponse.access_token


$Credentials = @{
    credentialData = @(
        @{
            name  = "accessToken"
            value = $SharePointAccessToken
        }
    )
} | ConvertTo-Json -Compress -Depth 5

$Body = @{
    credentialDetails = @{
        credentialType              = "OAuth2"
        credentials                 = $Credentials
        encryptedConnection         = "Encrypted"
        encryptionAlgorithm         = "None"
        privacyLevel                = "Organizational"
        useEndUserOAuth2Credentials = $false
    }
} | ConvertTo-Json -Depth 10


Invoke-PowerBIRestMethod `
    -Url "gateways/$GatewayId/datasources/$DatasourceId" `
    -Method Patch `
    -Body $Body