# Get users from an Entra group and sync them into a Power Platform environment.

$GroupId = "3dbbf8b5-ad18-4c15-9ecb-f8c8418ac38e"          # Entra group Object ID
$EnvironmentName = "b8c68c04-3c46-e5b2-9c78-53225e2b22ce"  # Power Platform environment ID

Connect-MgGraph -Scopes "GroupMember.Read.All", "User.Read.All"
Add-PowerAppsAccount

$members = Get-MgGroupMember -GroupId $GroupId -All

foreach ($member in $members) {
    if ($member.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.user') { continue }

    $user = Get-MgUser -UserId $member.Id -Property DisplayName, UserPrincipalName
    Write-Host "Syncing $($user.DisplayName) ($($user.UserPrincipalName))..." -ForegroundColor Cyan

    $result = Add-AdminPowerAppsSyncUser -EnvironmentName $EnvironmentName -PrincipalObjectId $member.Id
    if ($result.Code -eq 200 -and -not $result.Error) {
        Write-Host "Synced $($user.DisplayName)" -ForegroundColor Green
    }
    else {
        $reason = if ($result.Error.message) { $result.Error.message } else { $result.Description }
        if ($reason -match 'userNotLicensed|does not have a valid active license') {
            $reason = "No valid Power Apps / Dynamics license"
        }
        Write-Host "Failed $($user.DisplayName): $reason" -ForegroundColor Red
    }
}

Write-Host "Done." -ForegroundColor Green
