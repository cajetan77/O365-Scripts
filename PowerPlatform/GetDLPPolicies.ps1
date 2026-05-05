<#
.SYNOPSIS
    Exports Power Platform tenant data loss prevention (DLP) policies with full detail.

.DESCRIPTION
    Uses Get-AdminDlpPolicy for all policies, writes a deep JSON snapshot, flattened connector
    rows (Business / Non-Business / Blocked when present), environment scope rows, a per-policy
    details CSV (all properties; nested values as JSON in cells), and optional JSONL summary.
    When an Entra (AAD) tenant ID can be resolved, also calls Get-PowerAppDlpPolicyConnectorConfigurations
    and Get-PowerAppDlpPolicyExemptResources per policy (JSON + CSV).

.PARAMETER OutputDirectory
    Folder for output files. Defaults to the script directory.

.PARAMETER TenantId
    Optional Entra tenant GUID. If omitted, the script tries to resolve it from the default
    Power Platform environment (or its JSON metadata) so extended cmdlets can run.
#>
[CmdletBinding()]
param(
    [string] $OutputDirectory = $PSScriptRoot,
    [string] $TenantId
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name Microsoft.PowerApps.Administration.PowerShell)) {
    Write-Error "Install the admin module first: Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force"
}

Import-Module -Name Microsoft.PowerApps.Administration.PowerShell
Add-PowerAppsAccount

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

function Resolve-EntraTenantId {
    param([string] $ExplicitTenantId)
    if ($ExplicitTenantId -and $ExplicitTenantId.Trim() -match '^[0-9a-fA-F-]{36}$') {
        return $ExplicitTenantId.Trim()
    }
    $env = $null
    try {
        $env = Get-AdminPowerAppEnvironment -Default -ErrorAction Stop
    }
    catch {
        $env = Get-AdminPowerAppEnvironment -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $env) { return $null }
    foreach ($n in @('TenantId', 'AadTenantId', 'AzureTenantId')) {
        $p = $env.PSObject.Properties[$n]
        if ($p -and $p.Value -is [string] -and $p.Value -match '^[0-9a-fA-F-]{36}$') {
            return $p.Value
        }
    }
    $dump = $env | ConvertTo-Json -Depth 12 -Compress -ErrorAction SilentlyContinue
    if ($dump) {
        $m = [regex]::Match($dump, '"azureTenantId"\s*:\s*"([0-9a-fA-F-]{36})"', 'IgnoreCase')
        if ($m.Success) { return $m.Groups[1].Value }
        $m2 = [regex]::Match($dump, '"tenantId"\s*:\s*"([0-9a-fA-F-]{36})"', 'IgnoreCase')
        if ($m2.Success) { return $m2.Groups[1].Value }
    }
    return $null
}

function Get-PolicyKey {
    param($Policy)
    foreach ($n in @('PolicyName', 'Name', 'InternalName', 'Id')) {
        $p = $Policy.PSObject.Properties[$n]
        if ($p -and $p.Value) { return [string] $p.Value }
    }
    if ($Policy.DisplayName) { return [string] $Policy.DisplayName }
    return 'Unknown'
}

function ConvertTo-PolicyDetailOrdered {
    param($Policy)
    $lineObj = [ordered] @{
        PolicyKey = (Get-PolicyKey -Policy $Policy)
    }
    foreach ($prop in $Policy.PSObject.Properties) {
        $name = $prop.Name
        if ($name -in @('BusinessDataGroup', 'NonBusinessDataGroup', 'BlockedGroup', 'BlockedDataGroup', 'BlockedConnectors', 'Environments')) {
            $arr = @($prop.Value)
            $lineObj["${name}_Count"] = $arr.Count
            $lineObj["${name}_Json"] = ($prop.Value | ConvertTo-Json -Depth 12 -Compress -ErrorAction SilentlyContinue)
            continue
        }
        $v = $prop.Value
        if ($null -eq $v) {
            $lineObj[$name] = $null
        }
        elseif ($v -is [string] -or $v -is [valueType] -or $v -is [datetime]) {
            $lineObj[$name] = $v
        }
        else {
            $lineObj[$name] = ($v | ConvertTo-Json -Depth 8 -Compress -ErrorAction SilentlyContinue)
        }
    }
    return [pscustomobject] $lineObj
}

function Export-ObjectsToCsvUnifiedColumns {
    param(
        [System.Collections.IEnumerable] $Objects,
        [string] $LiteralPath
    )
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($o in $Objects) { if ($null -ne $o) { $list.Add($o) } }
    if ($list.Count -eq 0) {
        Write-Warning "No rows to export to $(Split-Path -Leaf $LiteralPath); file not written."
        return
    }
    $keyOrder = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $list) {
        foreach ($p in $r.PSObject.Properties) {
            if (-not $keyOrder.Contains($p.Name)) {
                $keyOrder.Add($p.Name)
            }
        }
    }
    $normalized = foreach ($r in $list) {
        $o = [ordered] @{}
        foreach ($k in $keyOrder) {
            $prop = $r.PSObject.Properties[$k]
            $o[$k] = if ($prop) { $prop.Value } else { $null }
        }
        [pscustomobject] $o
    }
    $normalized | Export-Csv -LiteralPath $LiteralPath -NoTypeInformation -Encoding utf8
}

function Add-ConnectorRows {
    param(
        $Policy,
        [string] $PolicyKey,
        [string] $PolicyDisplayName,
        [string] $GroupName,
        $Connectors,
        [System.Collections.Generic.List[object]] $Rows
    )
    if ($null -eq $Connectors) { return }
    foreach ($c in @($Connectors)) {
        if ($null -eq $c) { continue }
        $Rows.Add([pscustomobject] @{
                PolicyKey           = $PolicyKey
                PolicyDisplayName   = $PolicyDisplayName
                DataGroup           = $GroupName
                ConnectorId         = $c.id
                ConnectorName       = $c.name
                ConnectorType       = $c.type
                ConnectorJson       = ($c | ConvertTo-Json -Compress -Depth 8 -ErrorAction SilentlyContinue)
            })
    }
}

$stamp = Get-Date -Format 'yyyyMMddHHmmss'
$pathJsonFull = Join-Path $OutputDirectory "DLPPolicies-Full-$stamp.json"
$pathCsvConnectors = Join-Path $OutputDirectory "DLPPolicies-Connectors-$stamp.csv"
$pathCsvEnvironments = Join-Path $OutputDirectory "DLPPolicies-Environments-$stamp.csv"
$pathCsvPoliciesDetail = Join-Path $OutputDirectory "DLPPolicies-Policies-$stamp.csv"
$pathPolicySummaryJsonl = Join-Path $OutputDirectory "DLPPolicies-Summary-$stamp.jsonl"
$pathJsonExtended = Join-Path $OutputDirectory "DLPPolicies-ExtendedApis-$stamp.json"
$pathCsvExtended = Join-Path $OutputDirectory "DLPPolicies-ExtendedApis-$stamp.csv"

Write-Host "Retrieving all DLP policies..." -ForegroundColor Cyan
$policies = @(Get-AdminDlpPolicy)
Write-Host "Policy count: $($policies.Count)" -ForegroundColor Green

$policies | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $pathJsonFull -Encoding utf8
Write-Host "Wrote full snapshot: $pathJsonFull" -ForegroundColor Green

$connectorRows = [System.Collections.Generic.List[object]]::new()
$environmentRows = [System.Collections.Generic.List[object]]::new()
$policyDetailRows = [System.Collections.Generic.List[object]]::new()

foreach ($policy in $policies) {
    $pk = Get-PolicyKey -Policy $policy
    $pd = $null
    if ($policy.PSObject.Properties['DisplayName']) { $pd = $policy.DisplayName }

    Add-ConnectorRows -Policy $policy -PolicyKey $pk -PolicyDisplayName $pd -GroupName 'BusinessDataGroup' -Connectors $policy.BusinessDataGroup -Rows $connectorRows
    Add-ConnectorRows -Policy $policy -PolicyKey $pk -PolicyDisplayName $pd -GroupName 'NonBusinessDataGroup' -Connectors $policy.NonBusinessDataGroup -Rows $connectorRows

    foreach ($blockedProp in @('BlockedGroup', 'BlockedDataGroup', 'BlockedConnectors')) {
        if ($policy.PSObject.Properties[$blockedProp]) {
            Add-ConnectorRows -Policy $policy -PolicyKey $pk -PolicyDisplayName $pd -GroupName $blockedProp -Connectors $policy.$blockedProp -Rows $connectorRows
        }
    }

    $envList = $null
    if ($policy.PSObject.Properties['Environments']) { $envList = $policy.Environments }
    foreach ($e in @($envList)) {
        if ($null -eq $e) { continue }
        $environmentRows.Add([pscustomobject] @{
                PolicyKey         = $pk
                PolicyDisplayName = $pd
                EnvironmentJson   = ($e | ConvertTo-Json -Compress -Depth 10 -ErrorAction SilentlyContinue)
            })
    }

    $policyDetailRows.Add((ConvertTo-PolicyDetailOrdered -Policy $policy))
}

$connectorRows | Export-Csv -LiteralPath $pathCsvConnectors -NoTypeInformation -Encoding utf8
Write-Host "Wrote connectors (all policies / groups): $pathCsvConnectors" -ForegroundColor Green

$environmentRows | Export-Csv -LiteralPath $pathCsvEnvironments -NoTypeInformation -Encoding utf8
Write-Host "Wrote environment scope rows: $pathCsvEnvironments" -ForegroundColor Green

Export-ObjectsToCsvUnifiedColumns -Objects $policyDetailRows -LiteralPath $pathCsvPoliciesDetail
Write-Host "Wrote per-policy details (CSV): $pathCsvPoliciesDetail" -ForegroundColor Green

foreach ($row in $policyDetailRows) {
    ($row | ConvertTo-Json -Compress -Depth 12) | Add-Content -LiteralPath $pathPolicySummaryJsonl -Encoding utf8
}
Write-Host "Wrote per-policy summary (JSON lines): $pathPolicySummaryJsonl" -ForegroundColor Green

$resolvedTenant = Resolve-EntraTenantId -ExplicitTenantId $TenantId
if (-not $resolvedTenant) {
    Write-Warning "Could not resolve Entra tenant ID; skipping Get-PowerAppDlpPolicyConnectorConfigurations and Get-PowerAppDlpPolicyExemptResources. Pass -TenantId '<guid>' if you need those exports."
}
else {
    Write-Host "Using TenantId $resolvedTenant for extended DLP API exports." -ForegroundColor Cyan
    $extended = [System.Collections.Generic.List[object]]::new()
    foreach ($policy in $policies) {
        $pn = $null
        if ($policy.PSObject.Properties['PolicyName'] -and $policy.PolicyName) {
            $pn = [string] $policy.PolicyName
        }
        if (-not $pn) { $pn = Get-PolicyKey -Policy $policy }

        $block = [ordered] @{
            PolicyName   = $pn
            DisplayName  = $policy.DisplayName
            TenantIdUsed = $resolvedTenant
        }
        try {
            $block['ConnectorConfigurations'] = Get-PowerAppDlpPolicyConnectorConfigurations -TenantId $resolvedTenant -PolicyName $pn -ErrorAction Stop
        }
        catch {
            $block['ConnectorConfigurationsError'] = $_.Exception.Message
        }
        try {
            $block['ExemptResources'] = Get-PowerAppDlpPolicyExemptResources -TenantId $resolvedTenant -PolicyName $pn -ErrorAction Stop
        }
        catch {
            $block['ExemptResourcesError'] = $_.Exception.Message
        }
        $extended.Add([pscustomobject] $block)
    }
    $extended | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $pathJsonExtended -Encoding utf8
    Write-Host "Wrote extended API payload: $pathJsonExtended" -ForegroundColor Green

    $extendedCsvRows = foreach ($item in $extended) {
        $ccJson = $null
        $erJson = $null
        if ($item.PSObject.Properties['ConnectorConfigurations'] -and $null -ne $item.ConnectorConfigurations) {
            $ccJson = $item.ConnectorConfigurations | ConvertTo-Json -Depth 25 -Compress -ErrorAction SilentlyContinue
        }
        if ($item.PSObject.Properties['ExemptResources'] -and $null -ne $item.ExemptResources) {
            $erJson = $item.ExemptResources | ConvertTo-Json -Depth 25 -Compress -ErrorAction SilentlyContinue
        }
        [pscustomobject] @{
            PolicyName                     = $item.PolicyName
            DisplayName                    = $item.DisplayName
            TenantIdUsed                   = $item.TenantIdUsed
            ConnectorConfigurations_Json   = $ccJson
            ExemptResources_Json             = $erJson
            ConnectorConfigurationsError   = $item.ConnectorConfigurationsError
            ExemptResourcesError           = $item.ExemptResourcesError
        }
    }
    $extendedCsvRows | Export-Csv -LiteralPath $pathCsvExtended -NoTypeInformation -Encoding utf8
    Write-Host "Wrote extended API details (CSV): $pathCsvExtended" -ForegroundColor Green
}

Write-Host "Done." -ForegroundColor Cyan
