# Disables Managed Environment for every environment in a named environment group:
# 1) Resolve the environment group by display name (Power Platform API)
# 2) Find member environments via Get-AdminPowerAppEnvironment (avoids Environments.Read API permission)
# 3) Remove each environment from that group
# 4) Set governance protectionLevel to Basic
# Logs success and errors to a log file.

$ErrorActionPreference = 'Stop'

# --- CONFIG ---
$EnvironmentGroupDisplayName = "Personal Dev"
# Optional: set explicitly to skip group lookup. Known id for Personal Dev:
# $EnvironmentGroupId = "46238328-6636-4b48-9d62-2005d8b79ce8"
$EnvironmentGroupId = $null
$ApiVersion = "2024-10-01"
$EmptyGroupId = "00000000-0000-0000-0000-000000000000"
$LogDirectory = $PSScriptRoot
$LogPath = Join-Path $LogDirectory ("Disable-ManagedEnvironment_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [ValidateSet('INFO', 'SUCCESS', 'ERROR', 'WARN')]
        [string] $Level = 'INFO'
    )

    $line = "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $LogPath -Value $line -Encoding UTF8

    switch ($Level) {
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        default   { Write-Host $line -ForegroundColor Cyan }
    }
}

function Get-PowerPlatformApiHeaders {
    $token = Get-JwtToken -Audience "https://api.powerplatform.com/"
    return @{
        Authorization = "Bearer $token"
        Accept        = "application/json"
    }
}

function Invoke-PowerPlatformApi {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Get', 'Post', 'Delete')]
        [string] $Method,

        [Parameter(Mandatory)]
        [string] $Uri
    )

    $headers = Get-PowerPlatformApiHeaders
    return Invoke-WebRequest -Method $Method -Uri $Uri -Headers $headers -UseBasicParsing
}

function Get-EnvironmentGroupByDisplayName {
    param(
        [Parameter(Mandatory)]
        [string] $DisplayName
    )

    $uri = "https://api.powerplatform.com/environmentmanagement/environmentGroups?api-version=$ApiVersion"
    Write-Log "Listing environment groups..."
    $response = Invoke-PowerPlatformApi -Method Get -Uri $uri
    $payload = $response.Content | ConvertFrom-Json
    $groups = @($payload.value)
    if (-not $groups -or $groups.Count -eq 0) {
        $groups = @($payload)
    }

    $match = $groups | Where-Object { $_.displayName -eq $DisplayName } | Select-Object -First 1
    if (-not $match) {
        throw "Environment group not found with display name '$DisplayName'."
    }

    return $match
}

function Get-EnvironmentGroupIdFromAdminObject {
    param($EnvironmentObject)

    $groupId = $null
    try {
        $groupId = $EnvironmentObject.Internal.properties.parentEnvironmentGroup.id
    }
    catch {
        $groupId = $null
    }

    if ([string]::IsNullOrWhiteSpace($groupId) -or $groupId -eq $EmptyGroupId) {
        return $null
    }

    return [string]$groupId
}

function Get-AdminEnvironmentsInGroup {
    param(
        [Parameter(Mandatory)]
        [string] $GroupId
    )

    Write-Log "Scanning environments with Get-AdminPowerAppEnvironment for group $GroupId..."
    $allEnvs = @(Get-AdminPowerAppEnvironment)
    Write-Log "Total environments returned by admin cmdlet: $($allEnvs.Count)"

    $matched = @(
        $allEnvs | Where-Object {
            (Get-EnvironmentGroupIdFromAdminObject -EnvironmentObject $_) -eq $GroupId
        }
    )

    return $matched
}

function Remove-EnvironmentFromGroup {
    param(
        [Parameter(Mandatory)]
        [string] $EnvironmentName,

        [Parameter(Mandatory)]
        [string] $GroupId
    )

    $uri = "https://api.powerplatform.com/environmentmanagement/environmentGroups/$GroupId/removeEnvironment/$EnvironmentName`?api-version=$ApiVersion"
    Write-Log "Calling remove-from-group API for $EnvironmentName"
    $response = Invoke-PowerPlatformApi -Method Post -Uri $uri
    Write-Log "Remove-from-group API status for ${EnvironmentName}: $($response.StatusCode)"
    return $response
}

function Disable-ManagedEnvironmentForOne {
    param(
        [Parameter(Mandatory)]
        [string] $EnvironmentName,

        [Parameter(Mandatory)]
        [string] $GroupId,

        [string] $DisplayName
    )

    $label = if ($DisplayName) { "$DisplayName ($EnvironmentName)" } else { $EnvironmentName }
    Write-Log "----- Processing environment: $label -----"

    # Step 1: Remove from environment group
    try {
        Write-Log "Removing $label from environment group $GroupId..."
        Remove-EnvironmentFromGroup -EnvironmentName $EnvironmentName -GroupId $GroupId | Out-Null
        Write-Log "Removed $label from environment group." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to remove $label from environment group: $_" -Level ERROR
        throw
    }

    # Step 2: Change governance to Basic
    try {
        Write-Log "Setting governance protectionLevel to Basic for $label..."
        $UpdatedGovernanceConfiguration = [pscustomobject]@{
            protectionLevel = "Basic"
        }

        Set-AdminPowerAppEnvironmentGovernanceConfiguration `
            -EnvironmentName $EnvironmentName `
            -UpdatedGovernanceConfiguration $UpdatedGovernanceConfiguration

        Write-Log "Set $label governance to Basic (unmanaged)." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to set governance to Basic for ${label}: $_" -Level ERROR
        throw
    }
}

# --- START ---
Write-Log "Log file: $LogPath"
Write-Log "Starting Disable Managed Environment for group: $EnvironmentGroupDisplayName"

try {
    Write-Log "Signing in with Add-PowerAppsAccount..."
    Add-PowerAppsAccount
    Write-Log "Signed in successfully." -Level SUCCESS
}
catch {
    Write-Log "Failed to sign in: $_" -Level ERROR
    throw
}

$successCount = 0
$errorCount = 0
$processed = @()

try {
    if ([string]::IsNullOrWhiteSpace($EnvironmentGroupId)) {
        $group = Get-EnvironmentGroupByDisplayName -DisplayName $EnvironmentGroupDisplayName
        $EnvironmentGroupId = [string]$group.id
        Write-Log "Found environment group '$($group.displayName)' with id $EnvironmentGroupId." -Level SUCCESS
    }
    else {
        Write-Log "Using configured environment group id: $EnvironmentGroupId"
    }

    # Use admin PowerShell instead of Environments API (avoids EnvironmentManagement.Environments.Read permission error)
    $environments = @(Get-AdminEnvironmentsInGroup -GroupId $EnvironmentGroupId)

    Write-Log "Environments found in group: $($environments.Count)"
    if ($environments.Count -eq 0) {
        Write-Log "No environments to process in group '$EnvironmentGroupDisplayName' ($EnvironmentGroupId)." -Level WARN
    }

    foreach ($env in $environments) {
        $envId = $env.EnvironmentName
        $envDisplayName = if ($env.DisplayName) { $env.DisplayName } else { $envId }

        try {
            Disable-ManagedEnvironmentForOne -EnvironmentName $envId -GroupId $EnvironmentGroupId -DisplayName $envDisplayName
            $successCount++
            $processed += [pscustomobject]@{ Environment = $envDisplayName; Id = $envId; Result = 'SUCCESS' }
        }
        catch {
            $errorCount++
            $processed += [pscustomobject]@{ Environment = $envDisplayName; Id = $envId; Result = "ERROR: $_" }
            Write-Log "Continuing with remaining environments after failure on $envDisplayName." -Level WARN
        }
    }
}
catch {
    Write-Log "Fatal error while resolving/processing environment group: $_" -Level ERROR
    throw
}

Write-Log "----- SUMMARY -----"
Write-Log "Succeeded: $successCount"
Write-Log "Failed: $errorCount"
foreach ($row in $processed) {
    $level = if ($row.Result -eq 'SUCCESS') { 'SUCCESS' } else { 'ERROR' }
    Write-Log "$($row.Environment) [$($row.Id)] -> $($row.Result)" -Level $level
}

Write-Log "Completed. Full log saved to: $LogPath" -Level SUCCESS
