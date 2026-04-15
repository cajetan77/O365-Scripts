#PowerShell: Promote Co-owner → Owner
# 1) Install modules (run once)
#Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force
#Install-Module Microsoft.PowerApps.PowerShell -Scope CurrentUser -AllowClobber -Force

# 2) Sign in
Add-PowerAppsAccount

# 3) Find the app + environment (if you don’t already know them)
# Tip: filter by display name
$apps = Get-AdminPowerApp
#$newUser = "Bapty@cajesharepoint.onmicrosoft.com"
#$user = Get-AdminPowerAppsUserDetails -UserPrincipalName $newUser

$app = $apps | Where-Object { $_.DisplayName -eq "TestCaje" }

# 4) Promote the co-owner (set them as the new Owner)
$EnvironmentName = $app.EnvironmentName
$AppName = $app.AppName
# AppOwner is the target user's AAD ObjectId (recommended) or whatever identifier your tenant accepts
$NewOwnerObjectId = "c39ea9e8-2939-493f-96cd-9dc631a2cbb6" # Add  Service Account as Owner


$owners = Get-AdminPowerAppRoleAssignment -EnvironmentName $EnvironmentName -AppName $AppName | Where-Object { $_.RoleType -eq "Owner" } 



try {
    Set-AdminPowerAppOwner -EnvironmentName $EnvironmentName -AppName $AppName -AppOwner $NewOwnerObjectId
}
catch {
    Write-Host "Failed to promote co-owner to owner: $_"
    exit 1
}



Write-Host "Promoted co-owner with ObjectId $NewOwnerObjectId to Owner."


foreach ($owner in $owners) {
    Write-Host "Demoting owner  $($owner.PrincipalDisplayName) to Co-owner."  
    try {
        Set-AdminPowerAppRoleAssignment `
            -EnvironmentName $EnvironmentName `
            -AppName $AppName `
            -PrincipalType User `
            -PrincipalObjectId $owner.PrincipalObjectId `
            -RoleName CanEdit 
        Write-Host "Demoted owner  $($owner.PrincipalDisplayName) to Co-owner."        
    }
    catch {
        Write-Host "Failed to demote owner $($owner.PrincipalDisplayName): $_"
    }
    
}