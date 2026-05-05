[CmdletBinding()]
param(
    [string] $OutputDirectory = $PSScriptRoot,
    [switch] $SkipLogin
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name Microsoft.PowerApps.Administration.PowerShell)) {
    Write-Error "Install module first: Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force"
}

Import-Module -Name Microsoft.PowerApps.Administration.PowerShell

if (-not $SkipLogin) {
    Add-PowerAppsAccount
}

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

function ConvertTo-EnvironmentRow {
    param($EnvironmentObject)

    $row = [ordered] @{}
    foreach ($prop in $EnvironmentObject.PSObject.Properties) {
        $name = $prop.Name
        $value = $prop.Value
        if ($null -eq $value) {
            $row[$name] = $null
        }
        elseif ($value -is [string] -or $value -is [ValueType] -or $value -is [datetime]) {
            $row[$name] = $value
        }
        else {
            $row[$name] = $value | ConvertTo-Json -Depth 15 -Compress -ErrorAction SilentlyContinue
        }
    }

    return [pscustomobject] $row
}

function Get-SecurityGroupRestrictionInfo {
    param($EnvironmentObject)

    $groupIds = [System.Collections.Generic.List[string]]::new()
    $restricted = $false

    # Try known object properties first.
    $internal = $EnvironmentObject.PSObject.Properties['Internal']
    if ($internal -and $internal.Value) {
        $props = $internal.Value.PSObject.Properties['properties']
        if ($props -and $props.Value) {
            $connectedGroups = $props.Value.PSObject.Properties['connectedGroups']
            if ($connectedGroups -and $connectedGroups.Value) {
                foreach ($g in @($connectedGroups.Value)) {
                    $gid = $null
                    if ($g -is [string]) { $gid = $g }
                    elseif ($g.PSObject.Properties['id']) { $gid = [string] $g.id }
                    elseif ($g.PSObject.Properties['groupId']) { $gid = [string] $g.groupId }
                    if ($gid) {
                        $groupIds.Add($gid)
                        $restricted = $true
                    }
                }
            }
        }
    }

    # Fallback: look for common security-group fields in full JSON.
    $serialized = $EnvironmentObject | ConvertTo-Json -Depth 30 -Compress -ErrorAction SilentlyContinue
    if ($serialized) {
        foreach ($pattern in @(
                '"securityGroupId"\s*:\s*"([0-9a-fA-F-]{36})"',
                '"aadSecurityGroupId"\s*:\s*"([0-9a-fA-F-]{36})"',
                '"azureAdGroupId"\s*:\s*"([0-9a-fA-F-]{36})"',
                '"groupId"\s*:\s*"([0-9a-fA-F-]{36})"'
            )) {
            $regexMatches = [regex]::Matches($serialized, $pattern, 'IgnoreCase')
            foreach ($m in $regexMatches) {
                $gid = $m.Groups[1].Value
                if ($gid -and -not $groupIds.Contains($gid)) {
                    $groupIds.Add($gid)
                }
            }
        }
    }

    if ($groupIds.Count -gt 0) { $restricted = $true }

    [pscustomobject] @{
        RestrictedToSecurityGroup = $restricted
        SecurityGroupIds          = ($groupIds -join ';')
    }
}

function Get-ManagedAndDataverseStatus {
    param($EnvironmentObject)

    $isManaged = $null
    $managedProtectionLevel = ''
    $isDataverseEnabled = $false

    $serialized = $EnvironmentObject | ConvertTo-Json -Depth 30 -Compress -ErrorAction SilentlyContinue

    # Managed environment detection (authoritative source):
    # Internal.properties.governanceConfiguration
    $internalProp = $EnvironmentObject.PSObject.Properties['Internal']
    if ($internalProp -and $internalProp.Value) {
        $internalProperties = $internalProp.Value.PSObject.Properties['properties']
        if ($internalProperties -and $internalProperties.Value) {
            $govProp = $internalProperties.Value.PSObject.Properties['governanceConfiguration']
            if ($govProp) {
                $gov = $govProp.Value
                if ($null -eq $gov) {
                    $isManaged = $false
                }
                else {
                    $protectionProp = $gov.PSObject.Properties['protectionLevel']
                    if ($protectionProp -and $protectionProp.Value) {
                        $protection = [string]$protectionProp.Value
                        $managedProtectionLevel = $protection
                        if ($protection -eq 'Standard') {
                            $isManaged = $true
                        }
                        elseif ($protection -eq 'Basic') {
                            $isManaged = $false
                        }
                    }
                    else {
                        # governanceConfiguration exists but no explicit value.
                        $isManaged = $null
                    }
                }
            }
            else {
                $isManaged = $null
            }
        }
    }

    # Fallback only when Internal tree is absent from response.
    if ($null -eq $isManaged -and $serialized) {
        if ($serialized -match '"protectionLevel"\s*:\s*"Standard"') {
            $managedProtectionLevel = 'Standard'
            $isManaged = $true
        }
        elseif ($serialized -match '"protectionLevel"\s*:\s*"Basic"') {
            $managedProtectionLevel = 'Basic'
            $isManaged = $false
        }
        elseif ($serialized -match '"governanceConfiguration"\s*:\s*\{') {
            $isManaged = $null
        }
        elseif ($serialized -match '"isManaged"\s*:\s*true' -or $serialized -match '"managedEnvironment"\s*:\s*true') {
            $isManaged = $true
        }
        elseif ($serialized -match '"isManaged"\s*:\s*false' -or $serialized -match '"managedEnvironment"\s*:\s*false') {
            $isManaged = $false
        }
    }

    # Dataverse enabled detection
    if ($EnvironmentObject.PSObject.Properties['CommonDataServiceDatabaseProvisioningState']) {
        $provState = [string] $EnvironmentObject.CommonDataServiceDatabaseProvisioningState
        if ($provState -and $provState -ne 'Disabled' -and $provState -ne 'NotProvisioned') {
            $isDataverseEnabled = $true
        }
    }
    if (-not $isDataverseEnabled -and $EnvironmentObject.PSObject.Properties['OrganizationId']) {
        if ($EnvironmentObject.OrganizationId) {
            $isDataverseEnabled = $true
        }
    }
    if (-not $isDataverseEnabled -and $serialized) {
        if ($serialized -match '"linkedEnvironmentMetadata"\s*:\s*\{' -or $serialized -match '"instanceUrl"\s*:') {
            $isDataverseEnabled = $true
        }
    }

    
   

    [pscustomobject] @{
        ManagedProtectionLevel = $managedProtectionLevel
        IsManagedEnvironment   = if ($null -eq $isManaged) { '' } elseif ($isManaged) { 'Yes' } else { 'No' }
        DataverseEnabled       = if ($isDataverseEnabled) { 'Yes' } else { 'No' }
    }
}

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$csvPath = Join-Path $OutputDirectory "PowerPlatform-Environments-$timestamp.csv"

Write-Host "Retrieving Power Platform environments..." -ForegroundColor Cyan
$environments = @(Get-AdminPowerAppEnvironment -Capacity)



Write-Host "Environment count: $($environments.Count)" -ForegroundColor Green

$environments | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $fullJsonPath -Encoding utf8
Write-Host "Wrote full JSON: $fullJsonPath" -ForegroundColor Green

$rows = foreach ($environment in $environments) {
    $row = ConvertTo-EnvironmentRow -EnvironmentObject $environment
    $restrictionInfo = Get-SecurityGroupRestrictionInfo -EnvironmentObject $environment
    $statusInfo = Get-ManagedAndDataverseStatus -EnvironmentObject $environment

    $merged = [ordered] @{
        RestrictedToSecurityGroup = if ($restrictionInfo.RestrictedToSecurityGroup) { 'Yes' } else { '' }
        SecurityGroupIds          = if ($restrictionInfo.RestrictedToSecurityGroup) { $restrictionInfo.SecurityGroupIds } else { '' }
        ManagedProtectionLevel    = $statusInfo.ManagedProtectionLevel
        IsManagedEnvironment      = $statusInfo.IsManagedEnvironment
        DataverseEnabled          = $statusInfo.DataverseEnabled
    }
    foreach ($p in $row.PSObject.Properties) {
        $merged[$p.Name] = $p.Value
    }
    [pscustomobject] $merged
}

if (@($rows).Count -gt 0) {
    $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
    Write-Host "Wrote CSV: $csvPath" -ForegroundColor Green
}
else {
    Write-Warning "No environments returned; CSV not written."
}

Write-Host "Done." -ForegroundColor Cyan
