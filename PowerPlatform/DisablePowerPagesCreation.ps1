# Opens a prompt to collect credentials (Microsoft Entra account and password).
<#
.SYNOPSIS
    Disables the creation of Power Pages by non-admin users.
.DESCRIPTION
    This script disables the creation of Power Pages by non-admin users.
    Ensure user running the script has Power Platform Admin permissions.
 Refer to https://learn.microsoft.com/en-us/power-platform/admin/disable-creation-power-pages-non-admin-users for more information.
.PARAMETER RequestBody
    The request body to disable the creation of Power Pages by non-admin users.
#>


If (!(Get-Module -Name Microsoft.PowerApps.Administration.PowerShell)) {
    Write-Host "Installing Microsoft.PowerApps.Administration.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force
}
else {
    Write-Host "Microsoft.PowerApps.Administration.PowerShell module already installed..." -ForegroundColor Green
}
Add-PowerAppsAccount 

Set-TenantSettings -RequestBody @{ "disablePortalsCreationByNonAdminUsers" = $true }

Disconnect-Powe
