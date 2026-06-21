# --- SHAREPOINT CONFIG ---
$OutputDirectory = Get-Location
$ConfigPath = Join-Path $OutputDirectory 'config.json'
$ListName = 'Unused Flows'

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json

# 1. Install and authenticate
#Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force
#Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Force

Connect-MgGraph -ClientId $config.AppId -TenantId $config.TenantId -CertificateThumbprint $config.ThumbPrint -NoWelcome
if (-not (Get-Module -Name Microsoft.PowerApps.PowerShell)) {
    Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction Stop
}
Add-PowerAppsAccount

function Get-LastFlowRun {
    param(
        [Parameter(Mandatory)]
        [string]$EnvironmentName,
        [Parameter(Mandatory)]
        [string]$FlowName,
        [switch]$UsePnPFallback
    )

    try {
        $runs = @(Get-FlowRun -EnvironmentName $EnvironmentName -FlowName $FlowName -ErrorAction Stop)
        if ($runs.Count -gt 0) {
            $lastRun = $runs | Sort-Object { [datetime]$_.StartTime } -Descending | Select-Object -First 1
            return [PSCustomObject]@{
                StartTime = $lastRun.StartTime
                Status    = $lastRun.Status
            }
        }
    }
    catch {
        Write-Verbose "Get-FlowRun failed for '$FlowName': $_"
    }

    if ($UsePnPFallback) {
        try {
            $runs = @(Get-PnPFlowRun -Environment $EnvironmentName -Flow $FlowName -ErrorAction Stop)
            if ($runs.Count -gt 0) {
                $lastRun = $runs | Sort-Object {
                    $start = $_.Properties.StartTime
                    if (-not $start) { $start = $_.Properties.EndTime }
                    [datetime]$start
                } -Descending | Select-Object -First 1

                $startTime = $lastRun.Properties.StartTime
                if (-not $startTime) { $startTime = $lastRun.Properties.EndTime }

                return [PSCustomObject]@{
                    StartTime = $startTime
                    Status    = $lastRun.Properties.Status
                }
            }
        }
        catch {
            Write-Verbose "Get-PnPFlowRun failed for '$FlowName': $_"
        }
    }

    return $null
}

function ConvertTo-SharePointText {
    param([string]$Value)

    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    $cleaned = -join ($Value.ToCharArray() | ForEach-Object {
            $code = [int][char]$_
            if ($code -eq 0x9 -or $code -eq 0xA -or $code -eq 0xD -or ($code -ge 0x20 -and $code -le 0xD7FF) -or ($code -ge 0xE000 -and $code -le 0xFFFD)) {
                $_
            }
            elseif ($code -lt 0x20) {
                ' '
            }
        })

    return ($cleaned -replace '\s+', ' ').Trim()
}

$pnpConnected = $false
try {
    Connect-PnPOnline -Url $config.SiteUrl -ClientId $config.AppId -Tenant $config.TenantId -Thumbprint $config.ThumbPrint -ErrorAction Stop
    $pnpConnected = $true
}
catch {
    Write-Warning "PnP connection failed. Flow run fallback via Get-PnPFlowRun unavailable: $_"
}
# 2. Get all environments in the tenant
$DesiredEnvironment = ""

if ($DesiredEnvironment) {
    $Environments = Get-AdminPowerAppEnvironment | Where-Object { $_.DisplayName -eq $DesiredEnvironment }
}
else {
    $Environments = Get-AdminPowerAppEnvironment | Where-Object { $_.EnvironmentType -eq "Developer" -and $_.EnvironmentType -ne "Default" }
}
$OutputDirectory = Get-Location
$ReportPath = "$($OutputDirectory.Path)\Unused_Flows_Report.csv"

# Clear file if it exists from a previous run
if (Test-Path $ReportPath) { Remove-Item $ReportPath }

# 3. Loop through every environment
foreach ($Env in $Environments) {
    Write-Host "Scanning Environment: $($Env.DisplayName)" -ForegroundColor Cyan

    $environmentId = $Env.EnvironmentName

    
    # Get all flows in this specific environment
    $Flows = Get-AdminFlow -EnvironmentName $Env.EnvironmentName 
    # $Flows = Get-Flow -EnvironmentName $Env.EnvironmentName
    $i = 0
    foreach ($Flow in $Flows) {
        $i++;
        Write-Host "Scanning Flow: $($Flow.DisplayName) - $i of $($Flows.Count)" -ForegroundColor Cyan
        # Run history is limited to ~28 days; caller must be flow owner or have admin access.
        $RunHistory = Get-LastFlowRun -EnvironmentName $Env.EnvironmentName -FlowName $Flow.FlowName -UsePnPFallback:$pnpConnected
        if ($RunHistory) {
            Write-Host "Run History: $($RunHistory.StartTime) ($($RunHistory.Status))" -ForegroundColor Green
        }
        else {
            Write-Host "Run History: No runs in last 28 days" -ForegroundColor Yellow
        }

        $LastRunDate = if ($RunHistory) { $RunHistory.StartTime } else { "No runs in last 28 days" }
        $LastRunStatus = if ($RunHistory) { $RunHistory.Status } else { "N/A" }

        $FlowInfo = Get-AdminFlow -FlowName $Flow.FlowName -EnvironmentName $Env.EnvironmentName
        <#  $ConnectorDetails = @()
        $Connectors = $FlowInfo.Internal.properties.connectionReferences
        if ($Connectors) {
            $Connectors.PSObject.Properties | ForEach-Object {
                if ($_.Value.DisplayName) {
                    $ConnectorDetails += $_.Value.DisplayName
                }
            }
        }#>

        $owners = @()
        $coOwners = @()
        try {
            $permissions = Get-AdminFlowOwnerRole -EnvironmentName $Env.EnvironmentName -FlowName $Flow.FlowName -ErrorAction Stop
            $owners = @($permissions | Where-Object { $_.RoleType -eq 'Owner' } | ForEach-Object { $_.PrincipalObjectId })
            $coOwners = @($permissions | Where-Object { $_.RoleType -eq 'CanEdit' } | ForEach-Object { $_.PrincipalObjectId })
        }
        catch {
            Write-Host "Could not retrieve owners for $($Flow.DisplayName): $_" -ForegroundColor Yellow
        }

        [PSCustomObject]@{
            EnvironmentName = ConvertTo-SharePointText $Env.DisplayName
            EnvironmentType = ConvertTo-SharePointText $Env.EnvironmentType
            EnvironmentID   = $Env.EnvironmentName
            FlowName        = ConvertTo-SharePointText $Flow.DisplayName
            FlowID          = $Flow.FlowName
            FlowType        = ConvertTo-SharePointText $Flow.Internal.properties.flowType
            State           = ConvertTo-SharePointText $Flow.Internal.properties.state
            <#Connectors      = ($ConnectorDetails -join ' | ')#>
            Owner           = ($owners -join ' | ')
            CoOwner         = ($coOwners -join ' | ')
            CreatedBy       = $Flow.CreatedBy.userId
            SuspendedReason = ConvertTo-SharePointText $Flow.Internal.properties.flowSuspensionReason
            LastModified    = $Flow.LastModifiedTime
            LastRunDate     = $LastRunDate
            LastRunStatus   = ConvertTo-SharePointText $LastRunStatus
        } | Export-Csv -Path $ReportPath -Append -NoTypeInformation
    }
}




$csv = Import-Csv -Path $ReportPath
$userCache = @{}


function Get-MgIdentityInfo {
    param([string]$IdentityId)

    if ([string]::IsNullOrWhiteSpace($IdentityId)) { return $null }
    if ($IdentityId -match '@') {
        return @{ DisplayName = $IdentityId; UserPrincipalName = $IdentityId }
    }
    if ($userCache.ContainsKey($IdentityId)) { return $userCache[$IdentityId] }

    $info = @{ DisplayName = $IdentityId; UserPrincipalName = $null }
    try {
        $user = Get-MgUser -UserId $IdentityId -Property DisplayName, UserPrincipalName -ErrorAction Stop
        $info = @{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
        }
    }
    catch {
        try {
            $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $IdentityId -Property DisplayName -ErrorAction Stop
            $info.DisplayName = $servicePrincipal.DisplayName
        }
        catch {
            # keep object id as display fallback
        }
    }

    $userCache[$IdentityId] = $info
    return $info
}

function ConvertTo-SharePointDateTime {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ($Value -match '(?i)no runs|N/A') { return $null }

    $formats = @('dd-MM-yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm:ss', 'MM/dd/yyyy HH:mm:ss')
    foreach ($format in $formats) {
        try {
            return [datetime]::ParseExact($Value, $format, [cultureinfo]::InvariantCulture)
        }
        catch {
            # try next format
        }
    }

    [datetime]$parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [ref]$parsed)) { return $parsed }
    return $null
}

foreach ($item in $csv) {
    $createdByInfo = Get-MgIdentityInfo -IdentityId $item.CreatedBy
    if ($createdByInfo) {
        $item | Add-Member -NotePropertyName FlowCreatedByUpn -NotePropertyValue $createdByInfo.UserPrincipalName -Force
        if ($createdByInfo.DisplayName) { $item.CreatedBy = $createdByInfo.DisplayName }
    }

    $ownerUpn = $null
    $ownerNames = @()
    foreach ($ownerId in ($item.Owner -split ' \| ')) {
        $ownerId = $ownerId.Trim()
        if ([string]::IsNullOrWhiteSpace($ownerId)) { continue }

        $ownerInfo = Get-MgIdentityInfo -IdentityId $ownerId
        if (-not $ownerInfo) { continue }

        if ($ownerInfo.DisplayName) { $ownerNames += $ownerInfo.DisplayName }
        if (-not $ownerUpn -and $ownerInfo.UserPrincipalName) { $ownerUpn = $ownerInfo.UserPrincipalName }
    }
    $item | Add-Member -NotePropertyName OwnerUpn -NotePropertyValue $ownerUpn -Force
    if ($ownerNames.Count -gt 0) { $item.Owner = ($ownerNames -join ' | ') }

    $coOwnerNames = @()
    foreach ($coOwnerId in ($item.CoOwner -split ' \| ')) {
        $coOwnerId = $coOwnerId.Trim()
        if ([string]::IsNullOrWhiteSpace($coOwnerId)) { continue }

        $coOwnerInfo = Get-MgIdentityInfo -IdentityId $coOwnerId
        if ($coOwnerInfo -and $coOwnerInfo.DisplayName) {
            $coOwnerNames += $coOwnerInfo.DisplayName
        }
        else {
            $coOwnerNames += $coOwnerId
        }
    }
    if ($coOwnerNames.Count -gt 0) { $item.CoOwner = ConvertTo-SharePointText (($coOwnerNames -join ' | ')) }
}

$csv | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "User display names and UPNs resolved for SharePoint upload." -ForegroundColor Green

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Warning "Config not found at $ConfigPath. Skipping SharePoint upload."
}
else {
    if (-not $pnpConnected) {
        Connect-PnPOnline -Url $config.SiteUrl -ClientId $config.AppId -Tenant $config.TenantId -Thumbprint $config.ThumbPrint
        $pnpConnected = $true
    }

    $existingItemsByFlowId = @{}
    Get-PnPListItem -List $ListName -Fields 'FlowID' -PageSize 5000 | ForEach-Object {
        $flowId = $_.FieldValues.FlowID
        if (-not [string]::IsNullOrWhiteSpace($flowId)) {
            $existingItemsByFlowId[$flowId] = $_.Id
        }
    }

    $added = 0
    $updated = 0
    foreach ($item in $csv) {
        $values = @{
            Title           = ConvertTo-SharePointText $item.FlowName
            EnvironmentName = ConvertTo-SharePointText $item.EnvironmentName
            EnvironmentType = ConvertTo-SharePointText $item.EnvironmentType
            EnvironmentID   = $item.EnvironmentID
            FlowName        = ConvertTo-SharePointText $item.FlowName
            FlowID          = $item.FlowID
            FlowType        = ConvertTo-SharePointText $item.FlowType
            State           = ConvertTo-SharePointText $item.State
            CoOwner         = ConvertTo-SharePointText $item.CoOwner
            SuspendedReason = ConvertTo-SharePointText $item.SuspendedReason
            LastRunStatus   = ConvertTo-SharePointText $item.LastRunStatus
        }

        if (-not [string]::IsNullOrWhiteSpace($item.FlowCreatedByUpn)) {
            $values['FlowCreatedBy'] = $item.FlowCreatedByUpn
        }
        if (-not [string]::IsNullOrWhiteSpace($item.OwnerUpn)) {
            $values['Owner'] = $item.OwnerUpn
        }

        $lastModified = ConvertTo-SharePointDateTime -Value $item.LastModified
        if ($lastModified) { $values['LastModified'] = $lastModified }

        $lastRunDate = ConvertTo-SharePointDateTime -Value $item.LastRunDate
        if ($lastRunDate) { $values['LastRunDate'] = $lastRunDate }

        try {
            if ($existingItemsByFlowId.ContainsKey($item.FlowID)) {
                Set-PnPListItem -List $ListName -Identity $existingItemsByFlowId[$item.FlowID] -Values $values -ErrorAction Stop
                $updated++
            }
            else {
                $newItem = Add-PnPListItem -List $ListName -Values $values -ErrorAction Stop
                $existingItemsByFlowId[$item.FlowID] = $newItem.Id
                $added++
            }
        }
        catch {
            Write-Warning "Failed to sync '$($item.FlowName)' to '$ListName': $_"
        }
    }

    Write-Host "Synced $($csv.Count) flows to SharePoint list '$ListName' ($added added, $updated updated)." -ForegroundColor Green
}



