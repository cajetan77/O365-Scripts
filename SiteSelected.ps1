<#

.SYNOPSIS
    This script connects to Microsoft Graph and grants the selected application permission to the site.
.DESCRIPTION
    This script connects to Microsoft Graph and grants the selected application permission to the site.
    Create two applications in Azure AD one with Sites.FullControl.All permission and one with Sites.Selected permission.
    Enable the "Grant admin consent to this API" option for the Sites.Selected application
    Copy the client ID and certificate thumbprint of the Sites.Selected application and paste it in the config.json file.
    The Sites.FullControl.All application is used to connect to the site and grant permission to the site for the Sites.Selected.All application.
    Once script is run, you can remove the Sites.FullControl.All application from the Azure AD application registration.
.PARAMETER clientId
    The Client ID of the Azure AD App used for authentication.
.PARAMETER clientSecret
    The Client Secret of the Azure AD App used for authentication.
.PARAMETER TenantId
    The Tenant ID of the Azure AD App used for authentication.
.PARAMETER selectedclientId
    The Client ID of the selected Azure AD App used for authentication.
.PARAMETER selectedclientSecret
    The Client Secret of the selected Azure AD App used for authentication.
.PARAMETER Url
    The URL of the site which needs to be controlled.
#>


$config = Get-Content -Raw -Path ".\config.json" | ConvertFrom-Json

$clientId = $config.AppId
[string]$TenantId = $config.TenantId
$tenantName = $config.TenantName
$selectedclientId = $config.selectedAppID #ID with Sites.Selected Graph
$selectedAppDisplayName = $config.selectedAppDisplayName
$thumbprint = $config.ThumbPrint

$Urls = @("https://$tenantName.sharepoint.com/sites/CDocs", "https://$tenantName.sharepoint.com/sites/M365Updates","https://$tenantName.sharepoint.com/sites/CajGeneric113113") #Sites which needs to be controlled through the Sites.Selected.All application



#Connect-PnPOnline -Url $Url -ClientId $clientId -ClientSecret $clientSecret 

Function Connect_PnPOnline {
    param(
        [string]$Url,
        [string]$clientId,
        [string]$thumbprint,
        [string]$TenantId
    )
    try {
        Connect-PnPOnline -Url $Url -ClientId $clientId -Thumbprint $thumbprint  -Tenant $TenantId
    }
    catch {
        Write-Host "Error connecting to PnP Online: $_" -ForegroundColor Red
        exit
    }
}

Function Connect_MgGraph {
    param(
        [string]$clientId,
        [string]$thumbprint,
        [string]$TenantId
    )
    
    Write-Host "Checking Microsoft Graph module versions..." -ForegroundColor Yellow
    
    # Ensure Microsoft.Graph.Authentication is installed and updated
    $authModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
    if (-not $authModule) {
        Write-Host "Installing Microsoft.Graph.Authentication module..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    else {
        Write-Host "Updating Microsoft.Graph.Authentication module..." -ForegroundColor Yellow
        Update-Module -Name Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
    }
    
    # Import required modules
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Users -Force -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Groups -Force -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Sites -Force -ErrorAction SilentlyContinue

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    #$SecureClientSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force

    # Create a PSCredential Object Using the Client ID and Secure Client Secret
    #$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $SecureClientSecret
    
    # Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
    try {
        #Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -ErrorAction Stop
        Connect-MgGraph -TenantId $TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        Write-Host "Attempting alternative connection method..." -ForegroundColor Yellow
        # Alternative: Use ClientSecret parameter directly
        Connect-MgGraph -TenantId $TenantId -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph using alternative method." -ForegroundColor Green
    }
}




foreach ($Url in $Urls) {

    Connect_MgGraph -clientId $clientId -thumbprint $thumbprint -TenantId $TenantId
    $SiteId = $Url.Split('//')[1].split("/")[0] + ":/sites/" + $Url.Split('//')[1].split("/")[2]
    $Site = Get-MgSite -SiteId $SiteId
    $Application = @{}
    $Application.Add("id", $selectedclientId)
    $Application.Add("displayName", "SitesSelectedAll")
    $RequestedRole = "write"
    [array]$Permissions = Get-MgSitePermission -SiteId $Site.Id 
    ForEach ($Permission in $Permissions) {
        $Data = Get-MgSitePermission -PermissionId $Permission.Id -SiteId $Site.Id -Property Id, Roles, GrantedToIdentitiesV2
        Write-Host ("{0} permission available to {1}" -f ($Data.Roles -join ","), $Data.GrantedToIdentitiesV2.Application.DisplayName)
    }

    $Status = New-MgSitePermission -SiteId $Site.Id -Roles $RequestedRole -GrantedToIdentities @{"application" = $Application }
    If ($Status.id) { 
        Write-Host ("{0} permission granted to site {1}" -f $RequestedRole, $Site.DisplayName )
    }

    Disconnect-MgGraph

    Connect_MgGraph -clientId $selectedclientId -thumbprint $thumbprint -TenantId $TenantId
    Get-MgSitePage -SiteId $Site.Id
    Disconnect-MgGraph

 
    Connect_PnPOnline -Url $Url -ClientId $clientId -Thumbprint $thumbprint  -Tenant $TenantId
    try {
        Write-Host "Granting permission to $Url" -ForegroundColor Yellow
        $sitesperm=Grant-PnPAzureADAppSitePermission  -AppId $selectedclientId -DisplayName $selectedAppDisplayName -Permission "write" -Site $Url
        Write-Host "Permission $($sitesperm.Roles)  granted to $Url  for $sitesperm.Apps.DisplayName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error granting permission to $Url $_" -ForegroundColor Red
    }

    Disconnect-PnPOnline

    Connect_PnPOnline -Url $Url -ClientId $selectedclientId -Thumbprint $thumbprint  -Tenant $TenantId
    $web = Get-PnpWeb
    if ($web) {
        Write-Host "Web found for $Url"
    }
    else {
        Write-Host "Web not found for $Url" -ForegroundColor Red
    }
}


