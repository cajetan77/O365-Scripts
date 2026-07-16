<#
.SYNOPSIS
    Exports a production-ready Microsoft Intune device report to CSV.

.DESCRIPTION
    Read-only report for managed devices. Uses Microsoft Graph v1.0 only.

    Outputs:
      - Devices        : device inventory
      - DeviceSummary  : counts by OS, compliance, and stale sync
      - RunLog         : run metadata (duration, counts, paths)

    Optional -IncludePolicies adds compliance and configuration profile lists.

    Designed for large tenants (10k+ devices): single paginated list API, automatic
    retries on throttling, progress during retrieval.

.NOTES
    Application permissions (admin consent):
      DeviceManagementManagedDevices.Read.All
      DeviceManagementConfiguration.Read.All  (only with -IncludePolicies)

    Authentication: prefer certificate (-Thumbprint) over client secret in production.
#>
[CmdletBinding()]
Param(
    [switch]$IncludePolicies,
    [switch]$IncludeUnassignedDevices,
    [int]$StaleSyncDays = 30,
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
    Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop

    Write-ReportLog 'Connecting to Microsoft Graph...'
    if ($Thumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    }
    elseif ($ClientSecret) {
        $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential ([PSCredential]::new($ClientId, $secureSecret)) -NoWelcome
    }
    else {
        throw 'Provide -Thumbprint (recommended) or -ClientSecret for app authentication.'
    }

    Set-MgRequestContext -Retries $GraphRetries -RetryDelay $GraphRetryDelaySeconds | Out-Null
}

function Get-AllManagedDevices {
    param([int]$PageSize = 999)

    $devices = [System.Collections.Generic.List[object]]::new()
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=$PageSize"

    do {
        $page = Invoke-MgGraphRequest -Uri $uri -Method GET
        if ($page.value) {
            $devices.AddRange(@($page.value))
        }

        Write-Progress -Activity 'Retrieving managed devices' `
            -Status "$($devices.Count) device(s) retrieved" `
            -PercentComplete $(if ($page.'@odata.nextLink') { -1 } else { 100 })

        $uri = $page.'@odata.nextLink'
    } while ($uri)

    Write-Progress -Activity 'Retrieving managed devices' -Completed
    return @($devices)
}

function Test-DeviceIncluded {
    param(
        $Device,
        [switch]$AllowUnassigned
    )

    if ($AllowUnassigned) {
        return $true
    }
    return [bool]$Device.userPrincipalName
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        [string]$Path
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

    $devicesPath = Join-Path $exportRoot "Intune-Devices-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "Intune-DeviceSummary-$timestamp.csv"
    $runLogPath = Join-Path $exportRoot "Intune-RunLog-$timestamp.csv"
    $script:LogPath = Join-Path $exportRoot "Intune-RunLog-$timestamp.log"

    Write-ReportLog "Starting Intune report (TenantId: $TenantId)"
    Connect-MgGraphApp

    Write-ReportLog 'Retrieving managed devices from Graph v1.0...'
    $allDevices = Get-AllManagedDevices
    Write-ReportLog "Retrieved $($allDevices.Count) managed device(s) from Intune."

    $devices = @($allDevices | Where-Object { Test-DeviceIncluded -Device $_ -AllowUnassigned:$IncludeUnassignedDevices })
    if (-not $IncludeUnassignedDevices) {
        Write-ReportLog "$($devices.Count) device(s) with an assigned user."
    }

    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $osCounts = @{}
    $complianceCounts = @{}
    $nonCompliantCount = 0
    $staleSyncCount = 0
    $processed = 0

    foreach ($d in $devices) {
        $processed++
        if ($processed % 1000 -eq 0 -or $processed -eq $devices.Count) {
            Write-Progress -Activity 'Building report rows' -Status "$processed / $($devices.Count)" `
                -PercentComplete (($processed / [Math]::Max($devices.Count, 1)) * 100)
        }

        $lastSync = if ($d.lastSyncDateTime) { [datetime]$d.lastSyncDateTime } else { $null }
        $daysSinceSync = if ($lastSync) { [int](New-TimeSpan -Start $lastSync -End (Get-Date)).TotalDays } else { $null }

        $compliance = if ($d.complianceState) { [string]$d.complianceState } else { 'unknown' }
        $os = if ($d.operatingSystem) { [string]$d.operatingSystem } else { 'Unknown' }

        if ($compliance -eq 'noncompliant') { $nonCompliantCount++ }
        if ($daysSinceSync -ne $null -and $daysSinceSync -gt $StaleSyncDays) { $staleSyncCount++ }

        if (-not $osCounts.ContainsKey($os)) { $osCounts[$os] = 0 }
        if (-not $complianceCounts.ContainsKey($compliance)) { $complianceCounts[$compliance] = 0 }
        $osCounts[$os]++
        $complianceCounts[$compliance]++

        $deviceRows.Add([PSCustomObject]@{
            DeviceName        = $d.deviceName
            UserPrincipalName = $d.userPrincipalName
            UserDisplayName   = $d.userDisplayName
            OperatingSystem   = $d.operatingSystem
            OSVersion         = $d.osVersion
            ComplianceState   = $d.complianceState
            ManagementAgent   = $d.managementAgent
            EnrollmentType    = $d.deviceEnrollmentType
            IsEncrypted       = $d.isEncrypted
            EnrolledDateTime  = $d.enrolledDateTime
            LastSyncDateTime  = $d.lastSyncDateTime
            DaysSinceLastSync = $daysSinceSync
            Manufacturer      = $d.manufacturer
            Model             = $d.model
            SerialNumber      = $d.serialNumber
            AzureADDeviceId   = $d.azureADDeviceId
            DeviceId          = $d.id
        })
    }

    Write-Progress -Activity 'Building report rows' -Completed

    $summaryRows = [System.Collections.Generic.List[object]]::new()
    $summaryRows.Add([PSCustomObject]@{ Metric = 'ReportGeneratedUtc'; Value = (Get-Date).ToUniversalTime().ToString('o') })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TenantId'; Value = $TenantId })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TotalDevicesInTenant'; Value = $allDevices.Count })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TotalDevicesInReport'; Value = $deviceRows.Count })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'NonCompliant'; Value = $nonCompliantCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = "StaleSyncOver${StaleSyncDays}Days"; Value = $staleSyncCount })

    foreach ($entry in ($osCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "OS:$($entry.Key)"; Value = $entry.Value })
    }
    foreach ($entry in ($complianceCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "Compliance:$($entry.Key)"; Value = $entry.Value })
    }

    Export-ReportCsv -Data $deviceRows -Path $devicesPath
    Export-ReportCsv -Data $summaryRows -Path $summaryPath

    $outputFiles = @(
        [PSCustomObject]@{ FileType = 'Devices'; RowCount = $deviceRows.Count; Path = $devicesPath }
        [PSCustomObject]@{ FileType = 'DeviceSummary'; RowCount = $summaryRows.Count; Path = $summaryPath }
    )

    if ($IncludePolicies) {
        Write-ReportLog 'Retrieving compliance and configuration policies...'
        $compliancePath = Join-Path $exportRoot "Intune-CompliancePolicies-$timestamp.csv"
        $configPath = Join-Path $exportRoot "Intune-ConfigurationProfiles-$timestamp.csv"

        $compliance = @(Get-MgDeviceManagementDeviceCompliancePolicy -All)
        $config = @(Get-MgDeviceManagementDeviceConfiguration -All)

        Export-ReportCsv -Data ($compliance | Select-Object DisplayName, Id, Description, CreatedDateTime, LastModifiedDateTime, Version) -Path $compliancePath
        Export-ReportCsv -Data ($config | Select-Object DisplayName, Id, Description, CreatedDateTime, LastModifiedDateTime, Version) -Path $configPath

        $outputFiles += [PSCustomObject]@{ FileType = 'CompliancePolicies'; RowCount = $compliance.Count; Path = $compliancePath }
        $outputFiles += [PSCustomObject]@{ FileType = 'ConfigurationProfiles'; RowCount = $config.Count; Path = $configPath }

        Write-ReportLog "Compliance policies: $($compliance.Count)"
        Write-ReportLog "Configuration profiles: $($config.Count)"
    }

    $duration = (Get-Date) - $runStart
    $runMeta = [PSCustomObject]@{
        ReportGeneratedUtc = (Get-Date).ToUniversalTime().ToString('o')
        TenantId           = $TenantId
        DevicesInTenant    = $allDevices.Count
        DevicesInReport    = $deviceRows.Count
        DurationMinutes    = [Math]::Round($duration.TotalMinutes, 2)
        Status             = 'Success'
    }
    Export-ReportCsv -Data $runMeta -Path $runLogPath
    Export-ReportCsv -Data $outputFiles -Path (Join-Path $exportRoot "Intune-OutputFiles-$timestamp.csv")

    Write-ReportLog 'Report complete.'
    foreach ($file in $outputFiles) {
        Write-ReportLog ("  {0,-22} {1,6} rows -> {2}" -f $file.FileType, $file.RowCount, $file.Path)
    }
    Write-ReportLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)

    if ($allDevices.Count -eq 0) {
        Write-ReportLog 'WARNING: No managed devices returned. Verify Intune licensing, enrollment, and API permissions.'
    }
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
