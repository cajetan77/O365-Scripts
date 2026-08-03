# Aligns Microsoft Graph modules in Azure Automation to the same version.
# Run this once, wait until all modules show Available, then re-run Licensing.ps1.

$ResourceGroup = 'SPO-Automation'
$AutomationAccount = 'aa-spo-automation'   # <-- set your Automation Account name
$TargetVersion = '2.38.1'

$modules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
)

foreach ($moduleName in $modules) {
    $uri = "https://www.powershellgallery.com/api/v2/package/$moduleName/$TargetVersion"
    Write-Host "Importing $moduleName $TargetVersion..."

    az automation module import `
        --resource-group $ResourceGroup `
        --automation-account-name $AutomationAccount `
        --name $moduleName `
        --module-uri $uri `
        --output none

    if ($LASTEXITCODE -ne 0) {
        throw "Failed to import $moduleName $TargetVersion"
    }
}

Write-Host ''
Write-Host "Queued imports for version $TargetVersion."
Write-Host 'In the portal: Automation Account > Modules > wait until all four show Status = Available.'
Write-Host 'Then re-run the Licensing runbook.'
