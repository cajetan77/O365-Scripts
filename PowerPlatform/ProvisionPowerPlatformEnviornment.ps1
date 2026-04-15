# ================================
# CONFIGURATION
# ================================
$EnvironmentName = "Caj Prod Environment"
$EnvironmentSku = "Production"     # Trial | Sandbox | Production
$EnvironmentRegion = "australia"
$CurrencyName = "AUD"
$LanguageName = "3081"  # 3081 = English (Australia)
$SecurityGroupId = "4a6457ed-7d02-4faf-9a9e-c834e7dbed20"  # Optional: Specify a security group ID to restrict access

# Dataverse settings
$CreateDataverse = $false
#$DataverseName = "cajproddb"

# ================================
# AUTHENTICATE
# ================================
Write-Host "Authenticating to Power Platform..."
Add-PowerAppsAccount

# ================================
# CREATE ENVIRONMENT
# ================================
Write-Host "Creating Power Platform Environment..."
try {
    $environment = New-AdminPowerAppEnvironment `
        -DisplayName $EnvironmentName `
        -Location $EnvironmentRegion `
        -EnvironmentSku $EnvironmentSku `
        -CurrencyName $CurrencyName `
        -LanguageName $LanguageName `
        -ProvisionDatabase:$CreateDataverse `
        -SecurityGroupId $SecurityGroupId `
        -WaitUntilFinished $true

    Write-Host "Environment creation initiated."
    Write-Host "Environment ID:" $environment.EnvironmentName
}
catch {
    Write-Host "Error creating environment: $_" -ForegroundColor Red
    exit 1
}
# ================================
# OPTIONAL: WAIT FOR COMPLETION
# ================================
Write-Host "Waiting for environment to become ready..."

do {
    Start-Sleep -Seconds 30
    $status = Get-AdminPowerAppEnvironment -EnvironmentName $environment.EnvironmentName
    Write-Host "Current status:" $status.ProvisioningState
} while ($status.ProvisioningState -ne "Succeeded")

Write-Host "Environment successfully provisioned!"
