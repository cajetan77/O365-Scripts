# Grants Graph + SharePoint app roles to a system-assigned Managed Identity.
# Set $PrincipalId to the MI object (principal) id from Azure portal.

$PrincipalId = 'a99d2497-e8d9-4a56-9474-0be25075367b'

$GraphAppId = '00000003-0000-0000-c000-000000000000'      # Microsoft Graph
$SharePointAppId = '00000003-0000-0ff1-ce00-000000000000' # SharePoint Online

$Permissions = @(
    @{ AppId = $GraphAppId;      Name = 'User.Read.All' }
    @{ AppId = $GraphAppId;      Name = 'Group.Read.All' }
    @{ AppId = $GraphAppId;      Name = 'GroupMember.Read.All' }
    @{ AppId = $GraphAppId;      Name = 'AuditLog.Read.All' }
    @{ AppId = $GraphAppId;      Name = 'Organization.Read.All' }
    @{ AppId = $GraphAppId;      Name = 'Sites.ReadWrite.All' }
    @{ AppId = $SharePointAppId; Name = 'Sites.FullControl.All' }
)

Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome

$mi = Get-MgServicePrincipal -Filter "Id eq '$PrincipalId'"
if (-not $mi) { throw "Managed identity not found: $PrincipalId" }

Write-Host "Assigning permissions to $PrincipalId ($($mi.DisplayName))"

foreach ($permission in $Permissions) {
    $resource = Get-MgServicePrincipal -Filter "AppId eq '$($permission.AppId)'"
    $role = $resource.AppRoles |
        Where-Object { $_.Value -eq $permission.Name -and $_.AllowedMemberTypes -contains 'Application' } |
        Select-Object -First 1

    if (-not $role) {
        throw "Permission '$($permission.Name)' not found."
    }

    $exists = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $mi.Id -All |
        Where-Object { $_.ResourceId -eq $resource.Id -and $_.AppRoleId -eq $role.Id }

    if ($exists) {
        Write-Host "  Already assigned: $($permission.Name)"
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $mi.Id `
        -PrincipalId $mi.Id `
        -ResourceId $resource.Id `
        -AppRoleId $role.Id | Out-Null

    Write-Host "  Assigned: $($permission.Name)"
}

Write-Host 'Done. Wait a few minutes for token refresh before testing.'
