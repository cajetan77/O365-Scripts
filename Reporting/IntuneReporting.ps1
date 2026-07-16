<#
.SYNOPSIS
    Exports a production-ready device report using Get-MgDevice (Entra ID).

.DESCRIPTION
    Read-only report of all directory devices via Microsoft Graph v1.0 Get-MgDevice.
    This returns Entra-registered/joined devices (typically more complete than Intune
    managedDevices alone).

    Optionally enriches rows with Intune managed-device details when available
    (compliance, last sync, serial, encryption).

    Outputs:
      - Devices        : device inventory
      - DeviceSummary  : counts by OS, trust type, managed/compliant
      - RunLog         : run metadata

.NOTES
    Application permissions (admin consent):
      Device.Read.All                                   (required for Get-MgDevice)
      DeviceManagementManagedDevices.Read.All           (optional enrichment)
      DeviceManagementConfiguration.Read.All            (only with -IncludePolicies)
#>
[CmdletBinding()]
Param(
    [switch]$IncludePolicies,
    [switch]$SkipIntuneEnrichment,
    [switch]$EnabledDevicesOnly,
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
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue

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

function Get-AllEntraDevices {
    $properties = @(
        'id', 'deviceId', 'displayName', 'accountEnabled', 'operatingSystem', 'operatingSystemVersion',
        'trustType', 'isManaged', 'isCompliant', 'profileType', 'deviceOwnership', 'enrollmentType',
        'managementType', 'manufacturer', 'model', 'approximateLastSignInDateTime', 'registrationDateTime',
        'createdDateTime', 'deviceCategory', 'systemLabels'
    )

    Write-ReportLog 'Retrieving Entra devices with Get-MgDevice -All...'
    $devices = @(Get-MgDevice -All -Property $properties)
    Write-Progress -Activity 'Retrieving Entra devices' -Completed
    return $devices
}

function Get-IntuneManagedDeviceLookup {
    param([int]$PageSize = 999)

    $lookup = @{}
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=$PageSize&`$select=id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,complianceState,managementAgent,deviceEnrollmentType,isEncrypted,enrolledDateTime,lastSyncDateTime,manufacturer,model,serialNumber,azureADDeviceId"

    try {
        do {
            $page = Invoke-MgGraphRequest -Uri $uri -Method GET
            foreach ($md in @($page.value)) {
                if ($md.azureADDeviceId) {
                    $lookup[$md.azureADDeviceId.ToString()] = $md
                }
            }
            Write-Progress -Activity 'Retrieving Intune managed devices' `
                -Status "$($lookup.Count) managed device(s)" `
                -PercentComplete $(if ($page.'@odata.nextLink') { -1 } else { 100 })
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        Write-Progress -Activity 'Retrieving Intune managed devices' -Completed
    }
    catch {
        Write-ReportLog "WARNING: Intune enrichment skipped: $($_.Exception.Message)"
    }

    return $lookup
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

    $devicesPath = Join-Path $exportRoot "Intune-Devices-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "Intune-DeviceSummary-$timestamp.csv"
    $runLogPath = Join-Path $exportRoot "Intune-RunLog-$timestamp.csv"
    $script:LogPath = Join-Path $exportRoot "Intune-RunLog-$timestamp.log"

    Write-ReportLog "Starting device report (TenantId: $TenantId)"
    Connect-MgGraphApp

    $allDevices = Get-AllEntraDevices
    Write-ReportLog "Retrieved $($allDevices.Count) Entra device(s) via Get-MgDevice."

    if ($EnabledDevicesOnly) {
        $allDevices = @($allDevices | Where-Object { $_.AccountEnabled })
        Write-ReportLog "$($allDevices.Count) enabled device(s) after filter."
    }

    $intuneLookup = @{}
    if (-not $SkipIntuneEnrichment) {
        Write-ReportLog 'Enriching with Intune managedDevices (matched by AzureADDeviceId)...'
        $intuneLookup = Get-IntuneManagedDeviceLookup
        Write-ReportLog "Loaded $($intuneLookup.Count) Intune managed device(s) for enrichment."
    }

    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $osCounts = @{}
    $trustCounts = @{}
    $complianceCounts = @{}
    $managedInIntuneCount = 0
    $nonCompliantCount = 0
    $staleCount = 0
    $processed = 0

    foreach ($d in $allDevices) {
        $processed++
        if ($processed % 1000 -eq 0 -or $processed -eq $allDevices.Count) {
            Write-Progress -Activity 'Building report rows' -Status "$processed / $($allDevices.Count)" `
                -PercentComplete (($processed / [Math]::Max($allDevices.Count, 1)) * 100)
        }

        $deviceIdKey = if ($d.DeviceId) { $d.DeviceId.ToString() } else { '' }
        $intune = if ($deviceIdKey -and $intuneLookup.ContainsKey($deviceIdKey)) { $intuneLookup[$deviceIdKey] } else { $null }
        if ($intune) { $managedInIntuneCount++ }

        $lastActivity = $null
        if ($intune -and $intune.lastSyncDateTime) {
            $lastActivity = [datetime]$intune.lastSyncDateTime
        }
        elseif ($d.ApproximateLastSignInDateTime) {
            $lastActivity = [datetime]$d.ApproximateLastSignInDateTime
        }

        $daysSinceActivity = if ($lastActivity) {
            [int](New-TimeSpan -Start $lastActivity -End (Get-Date)).TotalDays
        }
        else { $null }

        $compliance = if ($intune -and $intune.complianceState) {
            [string]$intune.complianceState
        }
        elseif ($null -ne $d.IsCompliant) {
            if ($d.IsCompliant) { 'compliant' } else { 'noncompliant' }
        }
        else { 'unknown' }

        $os = if ($d.OperatingSystem) { [string]$d.OperatingSystem } else { 'Unknown' }
        $trust = if ($d.TrustType) { [string]$d.TrustType } else { 'unknown' }

        if ($compliance -eq 'noncompliant') { $nonCompliantCount++ }
        if ($daysSinceActivity -ne $null -and $daysSinceActivity -gt $StaleSyncDays) { $staleCount++ }

        if (-not $osCounts.ContainsKey($os)) { $osCounts[$os] = 0 }
        if (-not $trustCounts.ContainsKey($trust)) { $trustCounts[$trust] = 0 }
        if (-not $complianceCounts.ContainsKey($compliance)) { $complianceCounts[$compliance] = 0 }
        $osCounts[$os]++
        $trustCounts[$trust]++
        $complianceCounts[$compliance]++

        $deviceRows.Add([PSCustomObject]@{
            DeviceName                 = $d.DisplayName
            AzureADDeviceId            = $d.DeviceId
            ObjectId                   = $d.Id
            AccountEnabled             = $d.AccountEnabled
            OperatingSystem            = $d.OperatingSystem
            OSVersion                  = $d.OperatingSystemVersion
            TrustType                  = $d.TrustType
            IsManagedInEntra           = $d.IsManaged
            IsCompliantInEntra         = $d.IsCompliant
            ProfileType                = $d.ProfileType
            DeviceOwnership            = $d.DeviceOwnership
            EnrollmentType             = $d.EnrollmentType
            ManagementType             = $d.ManagementType
            Manufacturer               = $(if ($intune -and $intune.manufacturer) { $intune.manufacturer } else { $d.Manufacturer })
            Model                      = $(if ($intune -and $intune.model) { $intune.model } else { $d.Model })
            RegistrationDateTime       = $d.RegistrationDateTime
            ApproximateLastSignIn      = $d.ApproximateLastSignInDateTime
            DaysSinceLastActivity      = $daysSinceActivity
            InIntune                   = [bool]$intune
            UserPrincipalName          = $(if ($intune) { $intune.userPrincipalName } else { $null })
            UserDisplayName            = $(if ($intune) { $intune.userDisplayName } else { $null })
            IntuneComplianceState      = $(if ($intune) { $intune.complianceState } else { $null })
            IntuneManagementAgent      = $(if ($intune) { $intune.managementAgent } else { $null })
            IntuneEnrollmentType       = $(if ($intune) { $intune.deviceEnrollmentType } else { $null })
            IsEncrypted                = $(if ($intune) { $intune.isEncrypted } else { $null })
            IntuneEnrolledDateTime     = $(if ($intune) { $intune.enrolledDateTime } else { $null })
            IntuneLastSyncDateTime     = $(if ($intune) { $intune.lastSyncDateTime } else { $null })
            SerialNumber               = $(if ($intune) { $intune.serialNumber } else { $null })
            IntuneManagedDeviceId      = $(if ($intune) { $intune.id } else { $null })
            ComplianceState            = $compliance
        })
    }

    Write-Progress -Activity 'Building report rows' -Completed

    $summaryRows = [System.Collections.Generic.List[object]]::new()
    $summaryRows.Add([PSCustomObject]@{ Metric = 'ReportGeneratedUtc'; Value = (Get-Date).ToUniversalTime().ToString('o') })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TenantId'; Value = $TenantId })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'Source'; Value = 'Get-MgDevice (+ optional Intune enrichment)' })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TotalEntraDevices'; Value = $deviceRows.Count })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'MatchedInIntune'; Value = $managedInIntuneCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'NotInIntune'; Value = ($deviceRows.Count - $managedInIntuneCount) })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'NonCompliant'; Value = $nonCompliantCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = "StaleOver${StaleSyncDays}Days"; Value = $staleCount })

    foreach ($entry in ($osCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "OS:$($entry.Key)"; Value = $entry.Value })
    }
    foreach ($entry in ($trustCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "TrustType:$($entry.Key)"; Value = $entry.Value })
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
    }

    $duration = (Get-Date) - $runStart
    $runMeta = [PSCustomObject]@{
        ReportGeneratedUtc = (Get-Date).ToUniversalTime().ToString('o')
        TenantId           = $TenantId
        DevicesInReport    = $deviceRows.Count
        MatchedInIntune    = $managedInIntuneCount
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

    if ($deviceRows.Count -eq 0) {
        Write-ReportLog 'WARNING: No devices returned. Grant Device.Read.All and verify admin consent.'
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
