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

$appRole = $sharePointSp.AppRoles |
    Where-Object { $_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains 'Application' } |
    Select-Object -First 1

if (-not $appRole) {
    throw "App role '$PermissionName' not found on SharePoint Online"
}

$existing = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $managedIdentitySp.Id -All |
    Where-Object { $_.ResourceId -eq $sharePointSp.Id -and $_.AppRoleId -eq $appRole.Id }

if ($existing) {
    Write-Host "Already assigned: SharePoint $PermissionName"
}
else {
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $managedIdentitySp.Id `
        -PrincipalId $managedIdentitySp.Id `
        -ResourceId $sharePointSp.Id `
        -AppRoleId $appRole.Id

    Write-Host "Assigned SharePoint $PermissionName to $principalId"
}

Write-Host ''
Write-Host 'Restart the Function App after assigning permissions, then wait a few minutes for tokens to refresh.'
