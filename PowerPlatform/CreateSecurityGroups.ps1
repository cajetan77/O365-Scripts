<#
.SYNOPSIS
    Create security groups for Power Platform environments and set owners and members
.DESCRIPTION
    This script creates security groups for Power Platform environments.
.PARAMETER TicketNumber
    The ticket number for the security groups.
.PARAMETER GroupNamePrefix
    The prefix for the security group names.
.PARAMETER EnvName

#>
Connect-MgGraph -Scopes "User.ReadWrite.All", "GroupMember.ReadWrite.All", "Group.ReadWrite.All" -ErrorAction Stop
$ticketNumber = "326"
$groupNamePrefix = "CS_CAJ"
$envname = "Finance16"
$platformname = "PowerPlatform"
$owners = @("Lester Fernandes", "Sydelle Trindade")
$environments = @("Sandbox", "Prod")
$pproles = @("BasicUser", "SystemCustomizer", "SystemAdmin", "EnvMaker")
$BasicUser = @("Lester Fernandes", "Sydelle Trindade")
$SystemCustomizer = @("BIService")
$SystemAdmin = @("Flow Admin")
$EnvMaker = @("Test Flow")
$roleToUsers = @{
    BasicUser        = $BasicUser
    SystemCustomizer = $SystemCustomizer
    SystemAdmin      = $SystemAdmin
    EnvMaker         = $EnvMaker
}


foreach ($env in $environments) {
    foreach ($role in $pproles) {
        $securityGroupName = $groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env + "_" + $role
   
        $securityGroupDescription = "Security Group for " + $platformname + " - " + $envname + " - " + $env + " - " + $role + " - Ticket Number: " + $ticketNumber
        
        $existingGroup = Get-MgGroup -Filter "DisplayName eq '$securityGroupName'"
        if ($existingGroup) {
            Write-Host "Security Group already exists: $securityGroupName"
            continue
        }
        try {
            New-MgGroup -DisplayName $securityGroupName -Description $securityGroupDescription -MailEnabled:$false -SecurityEnabled:$true  -MailNickname $securityGroupName
            Write-Host "Security Group created: $securityGroupName"
        }
        catch {
            Write-Host "Error creating Security Group: $securityGroupName"
            Write-Host "Error: $_"
        }
    }
}

$ownerUsers = @()
foreach ($ownerName in $owners) {
    $escapedOwnerName = $ownerName -replace "'", "''"
    $ownerUser = Get-MgUser -Filter "displayName eq '$escapedOwnerName'"

    if (-not $ownerUser) {
        Write-Host "Owner user not found: $ownerName"
        continue
    }

    $ownerUsers += $ownerUser
}

foreach ($env in $environments) {
    $groupsToSetOwners = @()
    $groupsToSetOwners += ($groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env)

    foreach ($role in $pproles) {
        $groupsToSetOwners += ($groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env + "_" + $role)
    }

    foreach ($groupName in $groupsToSetOwners) {
        if ($groupName -eq ($groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env)) {
            Write-Host "Skipping owner assignment for group: $groupName"
            continue
        }

        $group = Get-MgGroup -Filter "DisplayName eq '$groupName'"

        if (-not $group) {
            Write-Host "Group not found for owner assignment: $groupName"
            continue
        }

        foreach ($ownerUser in $ownerUsers) {
            try {
                $existingOwner = Get-MgGroupOwner -GroupId $group.Id -All | Where-Object { $_.Id -eq $ownerUser.Id }

                if ($existingOwner) {
                    Write-Host "Owner already exists: $($ownerUser.DisplayName) in $groupName"
                    continue
                }

                $body = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($ownerUser.Id)"
                }
                New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter $body
                Write-Host "Added owner $($ownerUser.DisplayName) to $groupName"
            }
            catch {
                Write-Host "Error adding owner $($ownerUser.DisplayName) to $groupName"
                Write-Host "Error: $_"
            }
        }
    }
}

foreach ($env in $environments) {
    foreach ($role in $pproles) {
        $roleGroupName = $groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env + "_" + $role
        $roleGroup = Get-MgGroup -Filter "DisplayName eq '$roleGroupName'"

        if (-not $roleGroup) {
            Write-Host "Role Security Group not found: $roleGroupName"
            continue
        }

        $usersForRole = $roleToUsers[$role]
        if (-not $usersForRole) {
            Write-Host "No users configured for role: $role"
            continue
        }

        foreach ($userName in $usersForRole) {
            $escapedUserName = $userName -replace "'", "''"
            $user = Get-MgUser -Filter "displayName eq '$escapedUserName'"

            if (-not $user) {
                Write-Host "User not found: $userName"
                continue
            }

            try {
                $existingMember = Get-MgGroupMember -GroupId $roleGroup.Id -All | Where-Object { $_.Id -eq $user.Id }

                if ($existingMember) {
                    Write-Host "User already exists in group: $userName in $roleGroupName"
                    continue
                }

                New-MgGroupMember -GroupId $roleGroup.Id -DirectoryObjectId $user.Id
                Write-Host "Added $userName to $roleGroupName"
            }
            catch {
                Write-Host "Error adding $userName to $roleGroupName"
                Write-Host "Error: $_"
            }
        }
    }
}

foreach ($env in $environments) {
    $securityGroupName = $groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env 
    $securityGroupDescription = "Security Group for " + $platformname + " - " + $envname + " - " + $env + " - Ticket Number: " + $ticketNumber
    $existingGroup = Get-MgGroup -Filter "DisplayName eq '$securityGroupName'"
    if ($existingGroup) {
        Write-Host "Security Group already exists: $securityGroupName"
        continue
    }
    try {
        New-MgGroup -DisplayName $securityGroupName -Description $securityGroupDescription -MailEnabled:$false -SecurityEnabled:$true  -MailNickname $securityGroupName
        Write-Host "Security Group created: $securityGroupName"
    }
    catch {
        Write-Host "Error creating Security Group: $securityGroupName"
        Write-Host "Error: $_"
    }
}

foreach ($env in $environments) {
    $parentGroupName = $groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env
    $parentGroup = Get-MgGroup -Filter "DisplayName eq '$parentGroupName'"

    if (-not $parentGroup) {
        Write-Host "Parent Security Group not found: $parentGroupName"
        continue
    }

    foreach ($role in $pproles) {
        $childGroupName = $groupNamePrefix + "_" + $platformname + "_" + $envname + "_" + $env + "_" + $role
        $childGroup = Get-MgGroup -Filter "DisplayName eq '$childGroupName'"

        if (-not $childGroup) {
            Write-Host "Child Security Group not found: $childGroupName"
            continue
        }

        try {
            $existingMember = Get-MgGroupMember -GroupId $parentGroup.Id -All | Where-Object { $_.Id -eq $childGroup.Id }

            if ($existingMember) {
                Write-Host "Membership already exists: $childGroupName in $parentGroupName"
                continue
            }

            New-MgGroupMember -GroupId $parentGroup.Id -DirectoryObjectId $childGroup.Id
            Write-Host "Added $childGroupName to $parentGroupName"
        }
        catch {
            Write-Host "Error adding $childGroupName to $parentGroupName"
            Write-Host "Error: $_"
        }
    }
}