Add-PowerAppsAccount 
$tenantSettings = Get-TenantSettings
$tenantSettings.powerPlatform.governance.enableDefaultEnvironmentRouting = $False
Set-TenantSettings -RequestBody $tenantSettings
