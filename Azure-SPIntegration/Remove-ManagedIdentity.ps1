# Grants the Function App managed identity SharePoint Sites.FullControl.All.
# Required for Connect-PnPOnline -ManagedIdentity (SharePoint API, NOT Microsoft Graph).

$RESOURCE_GROUP = 'SPO-Automation'
$FUNCTION_APP_NAME = 'func-secure-processor02'
$PermissionName = 'Sites.FullControl.All'

# Office 365 SharePoint Online
$SharePointResourceAppId = '00000003-0000-0ff1-ce00-000000000000'

Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome

$principalId = az functionapp identity show `
    --name $FUNCTION_APP_NAME `
    --resource-group $RESOURCE_GROUP `
    --query principalId `
    --output tsv

if ([string]::IsNullOrWhiteSpace($principalId)) {
    throw "Managed identity not found on '$FUNCTION_APP_NAME'. Enable system-assigned identity first."
}

Write-Host "Function App: $FUNCTION_APP_NAME"
Write-Host "Managed identity object id: $principalId"

$managedIdentitySp = Get-MgServicePrincipal -Filter "Id eq '$principalId'"
$sharePointSp = Get-MgServicePrincipal -Filter "AppId eq '$SharePointResourceAppId'"

$spOAuth2PermissionsGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $managedIdentitySp.Id -All

# Remove all delegated permissions
$spOAuth2PermissionsGrants | ForEach-Object {
    Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $_.Id
}

# Get all application permissions for the service principal
$spApplicationPermissions = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySp.Id

# Remove all app role assignments
$spApplicationPermissions | ForEach-Object {
    Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $_.PrincipalId -AppRoleAssignmentId $_.Id
}