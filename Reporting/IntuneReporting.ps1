<#
.SYNOPSIS
    Exports all Entra ID devices using Get-MgDevice.

.DESCRIPTION
    Read-only inventory of directory devices via Microsoft Graph v1.0.
    Uses Get-MgDevice -All only (no Intune managedDevices calls).

.NOTES
    Application permission required (admin consent):
      Device.Read.All
#>
[CmdletBinding()]
Param(
    [switch]$EnabledDevicesOnly,
    [int]$StaleDays = 30,
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
    $runLogPath = Join-Path $exportRoot "RunLog-$timestamp.csv"
    $script:LogPath = Join-Path $exportRoot "RunLog-$timestamp.log"

    Write-ReportLog "Starting device report (TenantId: $TenantId)"
    Connect-MgGraphApp

    Write-ReportLog 'Retrieving devices with Get-MgDevice -All...'
    $allDevices = @(Get-MgDevice -All)
    Write-ReportLog "Get-MgDevice returned $($allDevices.Count) device(s)."

    if ($EnabledDevicesOnly) {
        $allDevices = @($allDevices | Where-Object { $_.AccountEnabled -eq $true })
        Write-ReportLog "Enabled devices: $($allDevices.Count)"
    }

    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $osCounts = @{}
    $trustCounts = @{}
    $staleCount = 0
    $disabledCount = 0
    $processed = 0

    foreach ($d in $allDevices) {
        $processed++
        if ($processed % 1000 -eq 0 -or $processed -eq $allDevices.Count) {
            Write-Progress -Activity 'Building report' -Status "$processed / $($allDevices.Count)" `
                -PercentComplete (($processed / [Math]::Max($allDevices.Count, 1)) * 100)
        }

        $lastSignIn = $null
        if ($d.ApproximateLastSignInDateTime) {
            $lastSignIn = [datetime]$d.ApproximateLastSignInDateTime
        }

        $daysSinceSignIn = if ($lastSignIn) {
            [int](New-TimeSpan -Start $lastSignIn -End (Get-Date)).TotalDays
        }
        else { $null }

        $os = if ($d.OperatingSystem) { [string]$d.OperatingSystem } else { 'Unknown' }
        $trust = if ($d.TrustType) { [string]$d.TrustType } else { 'unknown' }

        if ($d.AccountEnabled -eq $false) { $disabledCount++ }
        if ($daysSinceSignIn -ne $null -and $daysSinceSignIn -gt $StaleDays) { $staleCount++ }

        if (-not $osCounts.ContainsKey($os)) { $osCounts[$os] = 0 }
        if (-not $trustCounts.ContainsKey($trust)) { $trustCounts[$trust] = 0 }
        $osCounts[$os]++
        $trustCounts[$trust]++

        $deviceRows.Add([PSCustomObject]@{
            DisplayName                  = $d.DisplayName
            DeviceId                     = $d.DeviceId
            ObjectId                     = $d.Id
            AccountEnabled               = $d.AccountEnabled
            OperatingSystem              = $d.OperatingSystem
            OperatingSystemVersion       = $d.OperatingSystemVersion
            TrustType                    = $d.TrustType
            IsManaged                    = $d.IsManaged
            IsCompliant                  = $d.IsCompliant
            ProfileType                  = $d.ProfileType
            DeviceOwnership              = $d.DeviceOwnership
            EnrollmentType               = $d.EnrollmentType
            ManagementType               = $d.ManagementType
            Manufacturer                 = $d.Manufacturer
            Model                        = $d.Model
            ApproximateLastSignInDateTime = $d.ApproximateLastSignInDateTime
            DaysSinceLastSignIn          = $daysSinceSignIn
            RegistrationDateTime         = $d.RegistrationDateTime
            CreatedDateTime              = $d.CreatedDateTime
            DeviceCategory               = $d.DeviceCategory
        })
    }

    Write-Progress -Activity 'Building report' -Completed

    $summaryRows = [System.Collections.Generic.List[object]]::new()
    $summaryRows.Add([PSCustomObject]@{ Metric = 'ReportGeneratedUtc'; Value = (Get-Date).ToUniversalTime().ToString('o') })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TenantId'; Value = $TenantId })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'Source'; Value = 'Get-MgDevice -All' })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TotalDevices'; Value = $deviceRows.Count })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'DisabledDevices'; Value = $disabledCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = "StaleOver${StaleDays}Days"; Value = $staleCount })

    foreach ($entry in ($osCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "OS:$($entry.Key)"; Value = $entry.Value })
    }
    foreach ($entry in ($trustCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "TrustType:$($entry.Key)"; Value = $entry.Value })
    }

    Export-ReportCsv -Data $deviceRows -Path $devicesPath
    Export-ReportCsv -Data $summaryRows -Path $summaryPath

    $duration = (Get-Date) - $runStart
    $runMeta = [PSCustomObject]@{
        ReportGeneratedUtc = (Get-Date).ToUniversalTime().ToString('o')
        TenantId           = $TenantId
        TotalDevices       = $deviceRows.Count
        DurationMinutes    = [Math]::Round($duration.TotalMinutes, 2)
        Status             = 'Success'
        DevicesCsv         = $devicesPath
        SummaryCsv         = $summaryPath
    }
    Export-ReportCsv -Data $runMeta -Path $runLogPath

    Write-ReportLog 'Report complete.'
    Write-ReportLog "  Devices : $($deviceRows.Count) -> $devicesPath"
    Write-ReportLog "  Summary : $($summaryRows.Count) -> $summaryPath"
    Write-ReportLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)

    if ($deviceRows.Count -eq 0) {
        Write-ReportLog 'WARNING: Get-MgDevice returned 0 devices. Confirm Device.Read.All is granted with admin consent.'
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
