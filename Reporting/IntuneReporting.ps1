<#
.SYNOPSIS
    Exports Entra devices with RegisteredOwners and compliance status.

.DESCRIPTION
    Uses Get-MgDevice -All, resolves RegisteredOwners via Get-MgUser,
    and exports DisplayName, owners, and IsCompliant.

.NOTES
    Application permissions (admin consent):
      Device.Read.All
      User.Read.All
#>
[CmdletBinding()]
Param(
    [int]$BatchSize = 20,
    [int]$BatchDelayMs = 200,
    [int]$GraphRetries = 5,
    [int]$GraphRetryDelaySeconds = 10,

    [string]$TenantId = '764b46e8-d798-4ed3-87db-ae55ed7b0432',
    [string]$ClientId = 'dc223b11-5ab5-4a33-988a-3474b25eb9be',
    [string]$ClientSecret = '',
    [string]$Thumbprint = '800DB610ED947E9251A199ABFEA40AED1738128E',
    [string]$ExportPath = '.\Intune-Report'
)

$ErrorActionPreference = 'Stop'
$runStart = Get-Date

function Write-ReportLog {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function Connect-MgGraphApp {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    Write-ReportLog 'Connecting to Microsoft Graph...'
    if ($Thumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    }
    elseif ($ClientSecret) {
        $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential ([PSCredential]::new($ClientId, $secureSecret)) -NoWelcome
    }
    else {
        throw 'Provide -Thumbprint or -ClientSecret.'
    }

    Set-MgRequestContext -Retries $GraphRetries -RetryDelay $GraphRetryDelaySeconds | Out-Null
}

function Invoke-MgGraphBatch {
    param([array]$Requests, [int]$MaxRetries = 6)

    $body = @{ requests = $Requests } | ConvertTo-Json -Depth 6
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $result = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body
            return @($result.responses)
        }
        catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(60, 5 * $attempt))
                continue
            }
            throw
        }
    }
}

function Invoke-MgGraphBatchPages {
    param(
        [array]$AllRequests,
        [int]$Size = 20,
        [int]$DelayMs = 200,
        [string]$Activity = 'Graph batch'
    )

    $allResponses = [System.Collections.Generic.List[object]]::new()
    if ($AllRequests.Count -eq 0) { return @() }

    $totalBatches = [Math]::Ceiling($AllRequests.Count / [double]$Size)
    for ($i = 0; $i -lt $AllRequests.Count; $i += $Size) {
        $batchNum = [int]($i / $Size) + 1
        $end = [Math]::Min($i + $Size - 1, $AllRequests.Count - 1)
        $chunk = @($AllRequests[$i..$end])

        Write-Progress -Activity $Activity -Status "Batch $batchNum of $totalBatches" `
            -PercentComplete (($batchNum / $totalBatches) * 100)

        $allResponses.AddRange((Invoke-MgGraphBatch -Requests $chunk))

        if ($DelayMs -gt 0 -and ($i + $Size) -lt $AllRequests.Count) {
            Start-Sleep -Milliseconds $DelayMs
        }
    }

    Write-Progress -Activity $Activity -Completed
    return @($allResponses)
}

function Get-DeviceRegisteredOwnerIds {
    param(
        [array]$Devices,
        [int]$Size,
        [int]$DelayMs
    )

    $map = @{}
    foreach ($d in $Devices) {
        $map[$d.Id] = [System.Collections.Generic.List[string]]::new()
    }

    $requests = for ($i = 0; $i -lt $Devices.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/devices/$($Devices[$i].Id)/registeredOwners?`$select=id" }
    }

    Write-ReportLog "Loading RegisteredOwners for $($Devices.Count) device(s)..."
    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'RegisteredOwners'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $deviceId = $Devices[$index].Id
        if ($response.status -eq 200) {
            foreach ($obj in @($response.body.value)) {
                if ($obj.id) { $map[$deviceId].Add([string]$obj.id) }
            }
        }
    }

    return $map
}

function Import-UsersById {
    param(
        [string[]]$UserIds,
        [int]$Size,
        [int]$DelayMs
    )

    $cache = @{}
    $unique = @($UserIds | Where-Object { $_ } | Select-Object -Unique)
    if ($unique.Count -eq 0) { return $cache }

    Write-ReportLog "Resolving $($unique.Count) RegisteredOwner(s) with Get-MgUser..."
    $requests = for ($i = 0; $i -lt $unique.Count; $i++) {
        @{
            id     = "$i"
            method = 'GET'
            url    = "/users/$($unique[$i])?`$select=id,displayName,userPrincipalName"
        }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Get-MgUser'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $userId = $unique[$index]
        if ($response.status -eq 200) {
            $cache[$userId] = [PSCustomObject]@{
                DisplayName       = $response.body.displayName
                UserPrincipalName = $response.body.userPrincipalName
            }
        }
    }

    return $cache
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Data | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

try {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $exportRoot = [System.IO.Path]::GetFullPath($ExportPath)
    if (-not (Test-Path -LiteralPath $exportRoot)) {
        New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
    }

    $devicesPath = Join-Path $exportRoot "Devices-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "DeviceSummary-$timestamp.csv"
    $script:LogPath = Join-Path $exportRoot "RunLog-$timestamp.log"

    Write-ReportLog "Starting device report (TenantId: $TenantId)"
    Connect-MgGraphApp

    Write-ReportLog 'Retrieving devices with Get-MgDevice -All...'
    $allDevices = @(Get-MgDevice -All -Property Id, DeviceId, DisplayName, IsCompliant, AccountEnabled, OperatingSystem)
    Write-ReportLog "Get-MgDevice returned $($allDevices.Count) device(s)."

    $ownerMap = Get-DeviceRegisteredOwnerIds -Devices $allDevices -Size $BatchSize -DelayMs $BatchDelayMs

    $allOwnerIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $ownerMap.GetEnumerator()) {
        foreach ($id in $entry.Value) { [void]$allOwnerIds.Add($id) }
    }

    $userCache = Import-UsersById -UserIds @($allOwnerIds) -Size $BatchSize -DelayMs $BatchDelayMs
    Write-ReportLog "Resolved $($userCache.Count) owner(s)."

    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $compliantCount = 0
    $nonCompliantCount = 0
    $unknownComplianceCount = 0
    $noOwnerCount = 0
    $processed = 0

    foreach ($d in $allDevices) {
        $processed++
        if ($processed % 1000 -eq 0 -or $processed -eq $allDevices.Count) {
            Write-Progress -Activity 'Building report' -Status "$processed / $($allDevices.Count)" `
                -PercentComplete (($processed / [Math]::Max($allDevices.Count, 1)) * 100)
        }

        $ownerIds = @($ownerMap[$d.Id])
        $owners = @($ownerIds | ForEach-Object { $userCache[$_] } | Where-Object { $_ })
        $ownerNames = @($owners | ForEach-Object { $_.DisplayName })
        $ownerUpns = @($owners | ForEach-Object { $_.UserPrincipalName })

        $compliance = if ($null -eq $d.IsCompliant) {
            'Unknown'
        }
        elseif ($d.IsCompliant) {
            'Compliant'
        }
        else {
            'NonCompliant'
        }

        if ($compliance -eq 'Compliant') { $compliantCount++ }
        elseif ($compliance -eq 'NonCompliant') { $nonCompliantCount++ }
        else { $unknownComplianceCount++ }

        if ($owners.Count -eq 0) { $noOwnerCount++ }

        $deviceRows.Add([PSCustomObject]@{
            DisplayName        = $d.DisplayName
            DeviceId           = $d.DeviceId
            ObjectId           = $d.Id
            OperatingSystem    = $d.OperatingSystem
            AccountEnabled     = $d.AccountEnabled
            IsCompliant        = $d.IsCompliant
            ComplianceStatus   = $compliance
            RegisteredOwners   = ($ownerNames -join '; ')
            RegisteredOwnerUPN = ($ownerUpns -join '; ')
        })
    }

    Write-Progress -Activity 'Building report' -Completed

    $summaryRows = @(
        [PSCustomObject]@{ Metric = 'TotalDevices'; Value = $deviceRows.Count }
        [PSCustomObject]@{ Metric = 'Compliant'; Value = $compliantCount }
        [PSCustomObject]@{ Metric = 'NonCompliant'; Value = $nonCompliantCount }
        [PSCustomObject]@{ Metric = 'ComplianceUnknown'; Value = $unknownComplianceCount }
        [PSCustomObject]@{ Metric = 'NoRegisteredOwner'; Value = $noOwnerCount }
        [PSCustomObject]@{ Metric = 'ResolvedOwners'; Value = $userCache.Count }
    )

    Export-ReportCsv -Data $deviceRows -Path $devicesPath
    Export-ReportCsv -Data $summaryRows -Path $summaryPath

    $duration = (Get-Date) - $runStart
    Write-ReportLog 'Report complete.'
    Write-ReportLog "  Devices : $($deviceRows.Count) -> $devicesPath"
    Write-ReportLog "  Summary : $($summaryRows.Count) -> $summaryPath"
    Write-ReportLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)
}
catch {
    Write-ReportLog "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
