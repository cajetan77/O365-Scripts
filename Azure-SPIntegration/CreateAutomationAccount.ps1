#Requires -Version 7.0
<#
.SYNOPSIS
    Creates an Azure Automation Account for the O365 / SharePoint runbooks.

.DESCRIPTION
    Idempotent Azure CLI script that:
      1. Ensures resource group exists
      2. Creates Automation Account (Basic SKU)
      3. Enables system-assigned managed identity
      4. Creates common Automation Variables (empty placeholders)

    After this script:
      - Run Set-SystemManagedId.ps1 with the printed PrincipalId
      - Import modules: AzureAutomation\Update-GraphModules.ps1
      - Import PnP.PowerShell + ExchangeOnlineManagement via portal or az automation module import
      - Publish runbooks from AzureAutomation\

.EXAMPLE
    .\CreateAutomationAccount.ps1

.EXAMPLE
    .\CreateAutomationAccount.ps1 -SkipLogin -AutomationAccountName 'aa-spo-automation'
#>
[CmdletBinding()]
param(
    [string]$Subscription = 'Azure subscription 1',
    [string]$ResourceGroup = 'SPO-Automation',
    [string]$Location = 'australiasoutheast',
    [string]$AutomationAccountName = 'aa-spo-automation12',
    [ValidateSet('Free', 'Basic')]
    [string]$Sku = 'Basic',
    [switch]$SkipLogin,
    [switch]$SkipVariables
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

if (-not $SkipLogin) {
    Write-Step 'Azure login / subscription'
    Invoke-Az login
}
else {
    Write-Step "Using current Azure CLI session (subscription: $Subscription)"
}

Invoke-Az account set --subscription $Subscription

Write-Step "Ensure resource group '$ResourceGroup'"
if (-not (Test-AzResourceExists { az group show --name $ResourceGroup })) {
    Invoke-Az group create --name $ResourceGroup --location $Location
}
else {
    Write-Host 'Resource group already exists.'
}

Write-Step "Ensure Automation Account '$AutomationAccountName'"
# az automation account show is slow (automation CLI extension). ARM resource show returns immediately.
$accountExists = Test-AzResourceExists {
    az resource show `
        --resource-group $ResourceGroup `
        --name $AutomationAccountName `
        --resource-type 'Microsoft.Automation/automationAccounts' `
        --query id `
        --output tsv
}

if (-not $accountExists) {
    # The automation CLI extension does not accept --assign-identity.
    Invoke-Az automation account create `
        --name $AutomationAccountName `
        --resource-group $ResourceGroup `
        --location $Location `
        --sku $Sku
}
else {
    Write-Host 'Automation Account already exists.'
}

Write-Step 'Ensure system-assigned managed identity'
Invoke-Az resource update `
    --resource-group $ResourceGroup `
    --name $AutomationAccountName `
    --resource-type 'Microsoft.Automation/automationAccounts' `
    --set identity.type=SystemAssigned

$principalId = az resource show `
    --resource-group $ResourceGroup `
    --name $AutomationAccountName `
    --resource-type 'Microsoft.Automation/automationAccounts' `
    --query 'identity.principalId' `
    --output tsv

if ([string]::IsNullOrWhiteSpace($principalId)) {
    throw 'Could not read Automation Account managed identity principalId.'
}

if (-not $SkipVariables) {
    Write-Step 'Ensure Automation Variables (placeholders)'

    $variables = @(
        @{ Name = 'SHAREPOINT_SITE_URL'; Value = 'https://contoso.sharepoint.com/sites/IT' }
        @{ Name = 'EXCHANGE_ORGANIZATION'; Value = 'contoso.onmicrosoft.com' }
        @{ Name = 'SHAREPOINT_FOLDER_PATH'; Value = 'Shared Documents/LicensingReports' }
        @{ Name = 'SHAREPOINT_FOLDER_PATH_GROUPACCESS'; Value = 'Shared Documents/GroupAccessLogs' }
        @{ Name = 'SHAREPOINT_FOLDER_PATH_COPILOT'; Value = 'Shared Documents/CopilotReports' }
        @{ Name = 'SHAREPOINT_FOLDER_PATH_INTUNE'; Value = 'Shared Documents/IntuneReports' }
    )

    foreach ($var in $variables) {
        $exists = Test-AzResourceExists {
            az automation variable show `
                --automation-account-name $AutomationAccountName `
                --resource-group $ResourceGroup `
                --name $var.Name
        }

        if ($exists) {
            Write-Host "  Variable exists: $($var.Name)"
            continue
        }

        Invoke-Az automation variable create `
            --automation-account-name $AutomationAccountName `
            --resource-group $ResourceGroup `
            --name $var.Name `
            --value $var.Value `
            --encrypted false `
            --description 'Created by CreateAutomationAccount.ps1 — update in portal'

        Write-Host "  Created: $($var.Name)"
    }
}

Write-Host ""
Write-Host '=========================================================' -ForegroundColor Green
Write-Host 'Automation Account ready.'
Write-Host "  Name         : $AutomationAccountName"
Write-Host "  Resource group: $ResourceGroup"
Write-Host "  Location     : $Location"
Write-Host "  PrincipalId  : $principalId"
Write-Host '=========================================================' -ForegroundColor Green
Write-Host @"

Next steps:
  1. Set `$PrincipalId in Set-SystemManagedId.ps1 to '$principalId' and run it
  2. Portal > Automation Account > Identity > confirm System assigned = On
  3. Import modules (Graph): .\AzureAutomation\Update-GraphModules.ps1
  4. Import PnP.PowerShell + ExchangeOnlineManagement (portal or az automation module import)
  5. Update Automation Variables with real tenant / site values
  6. Publish runbooks from AzureAutomation\ and test on PowerShell 7.2 runtime
"@
