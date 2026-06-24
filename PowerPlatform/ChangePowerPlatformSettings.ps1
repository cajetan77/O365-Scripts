$logPath = Join-Path $PSScriptRoot ("ChangePowerPlatformSettings_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Status,
        [Parameter(Mandatory)]
        [string]$Message
    )

    $line = "{0:yyyy-MM-dd HH:mm:ss} | {1} | {2}" -f (Get-Date), $Status, $Message
    Add-Content -Path $logPath -Value $line

    switch ($Status) {
        'Success' { Write-Host $line -ForegroundColor Green }
        'Warning' { Write-Host $line -ForegroundColor Yellow }
        'Error' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line -ForegroundColor Cyan }
    }
}

Write-Log -Status 'Start' -Message "Log file: $logPath"

$systemAdminDirectFetchXml = "<fetch><entity name='role'><attribute name='roleid' /><attribute name='name' /><filter><condition attribute='name' operator='eq' value='System Administrator' /></filter><link-entity name='systemuserroles' from='roleid' to='roleid' intersect='true'><link-entity name='systemuser' from='systemuserid' to='systemuserid'><filter><condition attribute='systemuserid' operator='eq-userid' /></filter></link-entity></link-entity></entity></fetch>"

$systemAdminTeamFetchXml = "<fetch><entity name='role'><attribute name='roleid' /><attribute name='name' /><filter><condition attribute='name' operator='eq' value='System Administrator' /></filter><link-entity name='teamroles' from='roleid' to='roleid' intersect='true'><link-entity name='team' from='teamid' to='teamid'><link-entity name='teammembership' from='teamid' to='teamid' intersect='true'><link-entity name='systemuser' from='systemuserid' to='systemuserid'><filter><condition attribute='systemuserid' operator='eq-userid' /></filter></link-entity></link-entity></link-entity></link-entity></entity></fetch>"

function Test-PacOrgFetchHasSystemAdministrator {
    param(
        [Parameter(Mandatory)]
        [string]$EnvironmentId,
        [Parameter(Mandatory)]
        [string]$FetchXml
    )

    $output = pac org fetch --environment $EnvironmentId --xml $FetchXml 2>&1
    if ($LASTEXITCODE -ne 0) {
        return $null
    }

    return (($output | Out-String) -match '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\s+System Administrator')
}

function Get-SystemAdministratorAccess {
    param(
        [Parameter(Mandatory)]
        [string]$EnvironmentId
    )

    $viaDirect = Test-PacOrgFetchHasSystemAdministrator -EnvironmentId $EnvironmentId -FetchXml $systemAdminDirectFetchXml
    $viaTeam = Test-PacOrgFetchHasSystemAdministrator -EnvironmentId $EnvironmentId -FetchXml $systemAdminTeamFetchXml

    if ($null -eq $viaDirect -and $null -eq $viaTeam) {
        return $null
    }

    return [PSCustomObject]@{
        HasAccess           = ($viaDirect -eq $true) -or ($viaTeam -eq $true)
        ViaDirectAssignment = ($viaDirect -eq $true)
        ViaTeamOrGroup      = ($viaTeam -eq $true)
    }
}

function Get-SystemAdministratorSourceText {
    param(
        [Parameter(Mandatory)]
        [object]$Access
    )

    if ($Access.ViaDirectAssignment -and $Access.ViaTeamOrGroup) {
        return 'direct assignment and team/group membership'
    }

    if ($Access.ViaTeamOrGroup) {
        return 'team/group membership'
    }

    return 'direct assignment'
}

# Login first
pac auth create
if ($LASTEXITCODE -ne 0) {
    Write-Log -Status 'Error' -Message "pac auth create failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

# Get all environments
$envsJson = pac admin list --json
if ($LASTEXITCODE -ne 0) {
    Write-Log -Status 'Error' -Message "pac admin list failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

$envs = $envsJson | ConvertFrom-Json
Write-Log -Status 'Info' -Message "Found $($envs.Count) environment(s)"

$alreadySystemAdminEnvs = [System.Collections.Generic.List[string]]::new()
$addedSystemAdminEnvs = [System.Collections.Generic.List[string]]::new()

foreach ($env in $envs) {
    $envId = $env.EnvironmentId
    $envName = $env.DisplayName
    $envUrl = $env.EnvironmentUrl

    Write-Log -Status 'Info' -Message "Processing $envName | $envId | $envUrl"

    $adminAccess = Get-SystemAdministratorAccess -EnvironmentId $envId

    if ($adminAccess -and $adminAccess.HasAccess) {
        $alreadySystemAdminEnvs.Add($envName)
        $source = Get-SystemAdministratorSourceText -Access $adminAccess
        Write-Log -Status 'Info' -Message "System Administrator role already existed in $envName ($source)"
    }
    else {
        if ($null -eq $adminAccess) {
            Write-Log -Status 'Warning' -Message "Could not verify admin status in $envName; attempting self-elevate"
        }

        pac admin self-elevate --environment $envId
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Status 'Error' -Message "Self-elevate failed for $envName : exit code $LASTEXITCODE"
            continue
        }

        $addedSystemAdminEnvs.Add($envName)
        Write-Log -Status 'Success' -Message "System Administrator role added via self-elevate in $envName"
    }

    pac env update-settings --environment $envId --name iscomputeruseinmcsenabled --value false
    if ($LASTEXITCODE -ne 0) {
        Write-Log -Status 'Error' -Message "Setting update failed for $envName : exit code $LASTEXITCODE"
    }
    else {
        Write-Log -Status 'Success' -Message "Updated iscomputeruseinmcsenabled=false for $envName"
    }
}

if ($alreadySystemAdminEnvs.Count -gt 0) {
    Write-Log -Status 'Info' -Message "System Administrator already existed in: $($alreadySystemAdminEnvs -join ', ')"
}
else {
    Write-Log -Status 'Info' -Message 'System Administrator already existed in: none'
}

if ($addedSystemAdminEnvs.Count -gt 0) {
    Write-Log -Status 'Info' -Message "System Administrator added via self-elevate in: $($addedSystemAdminEnvs -join ', ')"
}
else {
    Write-Log -Status 'Info' -Message 'System Administrator added via self-elevate in: none'
}

Write-Log -Status 'Complete' -Message 'Finished processing all environments'
