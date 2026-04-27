Import-Module -Name Microsoft.PowerApps.Administration.PowerShell

$targetEnvironment = "A Caj1"
$environments = Get-AdminPowerAppEnvironment | Where-Object { $_.DisplayName -eq $targetEnvironment } # Include default environment
Set-AdminPowerAppSharepointFormEnvironment -EnvironmentName $environments.EnvironmentName