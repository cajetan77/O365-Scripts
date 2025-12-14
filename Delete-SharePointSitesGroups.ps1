<#

.SYNOPSIS
    Deletes a SharePoint Online Team Site permanently, including disconnecting it from its associated Office 365 Group if applicable.
.DESCRIPTION
    This script connects to the SharePoint Online Admin Center, checks if the specified Team Site is connected to an Office 365 Group, disconnects it if necessary, deletes the site, and then permanently removes it from the tenant's recycle bin.
.PARAMETER TenantName
    The name of the SharePoint Online tenant (e.g., "contoso" for contoso.sharepoint.com).  
.PARAMETER SiteUrl
    The URL of the SharePoint Online Team Site to be deleted (e.g., "https
://contoso.sharepoint.com/sites/TeamSite1").
.PARAMETER ClientId
    The Client ID of the Azure AD App used for authentication.
.PARAMETER TenantId
    The Tenant ID of the Azure AD App used for authentication.  
.PARAMETER Thumbprint
    The certificate thumbprint for the Azure AD App used for authentication.
#>

param(
   
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = ""   
)


# Load configuration from JSON file
$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$TenantName = $config.TenantName    
$Thumbprint = $config.ThumbPrint
$clientSecret = $config.ClientSecret

$adminUrl = "https://$TenantName-admin.sharepoint.com"

try {
    if ($clientSecret -ne $null ) {
        Connect-PnPOnline -Url $adminUrl  -ClientId $ClientId -ClientSecret $clientSecret
    }
    elseif ($Thumbprint -ne $null) {  
        Connect-PnPOnline -Url $adminUrl  -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
    }
    else {
        Connect-PnPOnline -Url $adminUrl -Interactive
    }
}
catch {
    Write-Host "Error connecting to SharePoint Online Admin Center: $_"
    exit
}

# --- (1) Check if the site is connected to an Office 365 Group ---
$tenantSite = Get-PnPTenantSite -Url $SiteUrl
$groupId = $tenantSite.GroupId

if ($groupId -ne [Guid]::Empty) {
    Write-Host "The site is connected to an Office 365 Group (GroupId: $groupId). Disconnecting the site from the group first..."
    try {
        Remove-PnPMicrosoft365Group -Identity $groupId
        Get-PnPDeletedMicrosoft365Group -Identity $groupId | Remove-PnPDeletedMicrosoft365Group 
    }
    catch {
        Write-Host "Error disconnecting the site from the Office 365 Group: $_"
        exit
    }
    $i = 0;
    do {

        $i++;
        $tenantSite = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
        $groupId = $tenantSite.GroupId
        Write-Host "Trying $i Waiting for the site to be disconnected from the group... Current GroupId: $groupId"
        Start-Sleep -Seconds 600
    
    
    } until (
        $groupId -eq [Guid]::Empty -or $groupId -eq $null )
    
}

else {
    Remove-PnPTenantSite -Url $SiteUrl -Force 
}
try {
    Clear-PnPTenantRecycleBinItem -Url $SiteUrl -Force -Wait
}
catch {
    Write-Host "Error permanently deleting the site from the tenant recycle bin: $_"
    exit
}

Disconnect-PnPOnline
Write-Host "Done. Site permanently deleted: $SiteUrl"
